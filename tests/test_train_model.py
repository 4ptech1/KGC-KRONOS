"""The tuning and export scripts (train/qlora_train.py, train/export_model.py; D128), as far as they can be tested without
a graphics card or a model: the dataset checks, the loss mask, the guards, the plan, the commands, the record.

Nothing here loads, downloads or runs a model, and nothing here may: the heavy steps are replaced by stand-ins.
"""
import contextlib
import io
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "train"))
sys.path.insert(0, str(ROOT / "compiler"))

import export_model as em  # noqa: E402
import make_dataset as md  # noqa: E402
import qlora_train as qt  # noqa: E402
from vcl_corpus import load  # noqa: E402


class FakeTokenizer:
    """A chat tokenizer in miniature: a character is a token; the template wraps each message."""

    def apply_chat_template(self, messages, tokenize=False, add_generation_prompt=False):
        text = "".join(f"<{m['role']}>{m['content']}</>" for m in messages)
        return text + ("<assistant>" if add_generation_prompt else "")

    def __call__(self, text, add_special_tokens=False):
        return {"input_ids": [ord(c) for c in text]}


class OddTokenizer(FakeTokenizer):
    def apply_chat_template(self, messages, tokenize=False, add_generation_prompt=False):
        text = super().apply_chat_template(messages, tokenize, add_generation_prompt)
        return text.replace("<system>", "<S>") if add_generation_prompt else text   # the prompt differs alone and whole


def make_dataset(folder, repairs=None):
    pairs = load("verified")
    rows = repairs if repairs is not None else []
    Path(folder).mkdir(parents=True, exist_ok=True)
    path = Path(folder) / "repairs.jsonl"
    path.write_text("".join(json.dumps(r) + "\n" for r in rows), encoding="utf-8")
    train, val, manifest = md.build(path)
    md.write_all(folder, train, val, manifest)
    return len(pairs)


class Libraries(unittest.TestCase):
    def test_importing_the_scripts_loads_no_model_library(self):
        code = ("import sys; sys.path[:0] = [%r, %r]; import qlora_train, export_model; "
                "bad = {'torch', 'transformers', 'peft', 'bitsandbytes', 'accelerate'} & set(sys.modules); "
                "print(sorted(bad)); sys.exit(1 if bad else 0)" % (str(ROOT / "train"), str(ROOT / "compiler")))
        done = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
        self.assertEqual(done.returncode, 0, done.stdout + done.stderr)

    def test_a_real_run_without_the_libraries_stops_in_plain_words(self):
        cfg = qt.Config("b", "d", "o")
        with mock.patch.object(qt, "missing_libraries", return_value=["torch"]):
            with self.assertRaises(qt.TrainError) as caught:
                qt.run_training(cfg)
        self.assertIn("new PC", str(caught.exception))


class Dataset(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.data = Path(self.tmp.name) / "data"
        make_dataset(self.data)

    def tearDown(self):
        self.tmp.cleanup()

    def retarget(self, file, edit):
        """Edit a dataset file and keep the manifest's hash right, so only the edit is under test."""
        path = self.data / file
        path.write_text(edit(path.read_text(encoding="utf-8")), encoding="utf-8")
        manifest = json.loads((self.data / "manifest.json").read_text(encoding="utf-8"))
        manifest["files"][file] = qt.sha256_file(path)
        (self.data / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")

    def test_a_whole_current_dataset_passes(self):
        train, val, manifest = qt.check_dataset(self.data)
        self.assertTrue(train and val)
        self.assertEqual(manifest["prompt_sha256"], md.digest(md.SYSTEM))

    def test_a_missing_file_is_named(self):
        (self.data / "val.jsonl").unlink()
        with self.assertRaisesRegex(qt.TrainError, "val.jsonl is missing"):
            qt.check_dataset(self.data)

    def test_a_changed_file_no_longer_matches_its_hash(self):
        with (self.data / "train.jsonl").open("a", encoding="utf-8") as f:
            f.write("\n")
        with self.assertRaisesRegex(qt.TrainError, "SHA-256"):
            qt.check_dataset(self.data)

    def test_a_dataset_made_with_another_prompt_is_refused(self):
        manifest = json.loads((self.data / "manifest.json").read_text(encoding="utf-8"))
        manifest["prompt_sha256"] = "0" * 64
        (self.data / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
        with self.assertRaisesRegex(qt.TrainError, "another prompt"):
            qt.check_dataset(self.data)

    def test_a_held_out_item_is_never_trained_on(self):
        held = sorted(qt.held_out_ids())[0]
        row = json.loads((self.data / "train.jsonl").read_text(encoding="utf-8").split("\n")[0])
        row["id"] = held + "/write0"
        self.retarget("train.jsonl", lambda t: t + json.dumps(row) + "\n")
        with self.assertRaisesRegex(qt.TrainError, "held-out"):
            qt.check_dataset(self.data)

    def test_a_row_that_is_not_system_user_assistant_is_refused(self):
        row = {"id": "c001/write0", "messages": [{"role": "user", "content": "x"}]}
        self.retarget("train.jsonl", lambda t: t + json.dumps(row) + "\n")
        with self.assertRaisesRegex(qt.TrainError, "not system, user, assistant"):
            qt.check_dataset(self.data)


class Encoding(unittest.TestCase):
    EXAMPLE = {"id": "c001/write0", "messages": [{"role": "system", "content": "S"}, {"role": "user", "content": "ask"},
                                                 {"role": "assistant", "content": "VCL 0.6"}]}

    def test_the_loss_is_on_the_answer_only(self):
        e = qt.encode_example(FakeTokenizer(), self.EXAMPLE, 1000)
        text = "".join(chr(i) for i in e["input_ids"])
        answer = "VCL 0.6</>"
        self.assertTrue(text.endswith(answer))
        kept = [chr(i) for i, label in zip(e["input_ids"], e["labels"]) if label != -100]
        self.assertEqual("".join(kept), answer)
        self.assertEqual(len(e["input_ids"]), len(e["labels"]))
        self.assertEqual(e["attention_mask"], [1] * len(e["input_ids"]))
        # the labels that are kept are the tokens themselves
        self.assertTrue(all(label == i for i, label in zip(e["input_ids"], e["labels"]) if label != -100))

    def test_an_example_over_the_limit_is_skipped_not_cut(self):
        full = len(qt.encode_example(FakeTokenizer(), self.EXAMPLE, 1000)["input_ids"])
        self.assertIsNotNone(qt.encode_example(FakeTokenizer(), self.EXAMPLE, full))
        self.assertIsNone(qt.encode_example(FakeTokenizer(), self.EXAMPLE, full - 1))

    def test_a_template_that_changes_the_prompt_is_refused(self):
        with self.assertRaisesRegex(qt.TrainError, "chat template"):
            qt.encode_example(OddTokenizer(), self.EXAMPLE, 1000)

    def test_mask_labels(self):
        self.assertEqual(qt.mask_labels([1, 2], [1, 2, 3, 4]), ([1, 2, 3, 4], [-100, -100, 3, 4]))
        with self.assertRaises(qt.TrainError):
            qt.mask_labels([1, 9], [1, 2, 3])

    def test_the_length_guess_errs_high(self):
        short = {"messages": [{"content": "x" * 100}]}
        long = {"messages": [{"content": "x" * 30000}]}
        self.assertEqual(qt.too_long([short, long], 8192), 1)
        self.assertEqual(qt.too_long([short, long], 20000), 0)


class Guards(unittest.TestCase):
    def test_nothing_is_written_inside_the_repo(self):
        for inside in (ROOT / "dist" / "run", ROOT / "train" / "runs", ROOT):
            with self.assertRaises(qt.TrainError, msg=str(inside)):
                qt.refuse_inside_repo(inside)
        with tempfile.TemporaryDirectory() as outside:
            qt.refuse_inside_repo(Path(outside) / "run")

    def test_the_reserve_stays_free(self):
        with tempfile.TemporaryDirectory() as d:
            base, out = Path(d) / "base", Path(d) / "out"
            base.mkdir()
            (base / "config.json").write_text("{}")
            data = Path(d) / "data"
            make_dataset(data)
            cfg = qt.Config(str(base), str(data), str(out))
            with mock.patch.object(qt, "free_gb", return_value=100.0), mock.patch.object(qt, "need_gb", return_value=20.0):
                self.assertEqual(qt.plan(cfg)[1], [])
            with mock.patch.object(qt, "free_gb", return_value=45.0), mock.patch.object(qt, "need_gb", return_value=20.0):
                self.assertTrue(any("reserve" in p for p in qt.plan(cfg)[1]))


class Command(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        d = Path(self.tmp.name)
        (d / "base").mkdir()
        (d / "base" / "config.json").write_text("{}")
        make_dataset(d / "data")
        self.args = ["--base", str(d / "base"), "--data", str(d / "data"), "--out", str(d / "runs" / "r1")]
        self.free = mock.patch.object(qt, "free_gb", return_value=500.0)
        self.free.start()

    def tearDown(self):
        self.free.stop()
        self.tmp.cleanup()

    def run_main(self, extra=()):
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = qt.main(self.args + list(extra))
        return code, out.getvalue()

    def test_without_go_it_only_plans(self):
        with mock.patch.object(qt, "run_training") as trained:
            code, text = self.run_main()
        self.assertEqual(code, 0)
        self.assertIn("dry run", text)
        trained.assert_not_called()
        self.assertFalse((Path(self.tmp.name) / "runs").exists())

    def test_with_go_it_trains_with_the_settings_given(self):
        with mock.patch.object(qt, "run_training") as trained:
            code, _ = self.run_main(["--go", "--smoke", "--epochs", "1", "--max-seq-len", "4096", "--no-4bit"])
        self.assertEqual(code, 0)
        cfg = trained.call_args[0][0]
        self.assertEqual((cfg.epochs, cfg.max_seq_len, cfg.smoke, cfg.load_4bit), (1.0, 4096, True, False))
        self.assertEqual((cfg.lora_r, cfg.grad_accum), (16, 16))

    def test_a_problem_stops_it_before_anything_runs(self):
        self.args[self.args.index("--base") + 1] = str(Path(self.tmp.name) / "missing")
        with mock.patch.object(qt, "run_training") as trained:
            code, text = self.run_main(["--go"])
        self.assertEqual(code, 1)
        self.assertIn("PROBLEM", text)
        trained.assert_not_called()

    def test_a_failed_run_says_stop(self):
        with mock.patch.object(qt, "run_training", side_effect=qt.TrainError("no card")):
            code, text = self.run_main(["--go"])
        self.assertEqual(code, 2)
        self.assertIn("STOP: no card", text)

    def test_the_run_record_can_be_written(self):
        manifest = {"files": {"train.jsonl": "a"}, "prompt_sha256": "p", "pairs": {"train": [], "val": []}}
        record = qt.run_record(qt.Config("b", "d", "o"), manifest, {"train": 1, "val": 1}, {"torch": "x"},
                               {"eval_loss": 0.5}, "t0", "t1")
        again = json.loads(json.dumps(record))
        self.assertEqual(again["config"]["lora_r"], 16)
        self.assertEqual(again["config"]["target_modules"], list(qt.Config("b", "d", "o").target_modules))


class Export(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        d = Path(self.tmp.name)
        self.base = d / "base"
        (self.base / ".cache").mkdir(parents=True)
        (self.base / "config.json").write_text("{}")
        (self.base / "model.safetensors").write_bytes(b"weights")
        (self.base / "README.md").write_text("model card")
        (self.base / "LICENSE").write_text("Apache-2.0")
        (self.base / ".cache" / "junk").write_text("x")
        self.run_dir = d / "runs" / "r1"
        (self.run_dir / "adapter").mkdir(parents=True)
        (self.run_dir / "adapter" / "adapter_config.json").write_text("{}")
        self.run_record = {"config": {"base": str(self.base)}, "examples": {"train": 5, "val": 1},
                           "losses": {"eval_loss": 0.4}, "libraries": {"torch": "x"},
                           "dataset": {"files": {"train.jsonl": "a"}, "prompt_sha256": "p"}, "smoke": False}
        (self.run_dir / "run.json").write_text(json.dumps(self.run_record))
        self.llama = d / "llama.cpp"
        (self.llama / "build" / "bin").mkdir(parents=True)
        (self.llama / "convert_hf_to_gguf.py").write_text("")
        self.cfg = em.Export(base=str(self.base), adapter=str(self.run_dir / "adapter"), out=str(d / "models" / "kronos-v1"),
                             llama_cpp=str(self.llama), base_name="Some-Coder-7B", base_version="1.0",
                             licence="Apache-2.0")
        self.cfg.quantize_bin.write_text("")
        self.free = mock.patch.object(em, "free_gb", return_value=500.0)
        self.free.start()

    def tearDown(self):
        self.free.stop()
        self.tmp.cleanup()

    def problems(self, **change):
        return em.check(em.Export(**{**self.cfg.__dict__, **change}))

    def test_a_complete_setup_has_no_problems(self):
        self.assertEqual(em.check(self.cfg), [])

    def test_each_thing_that_would_stop_it_is_named(self):
        cases = {
            "out inside the repo": ({"out": str(ROOT / "dist" / "m")}, "inside the repo"),
            "a quantisation that does not exist": ({"quant": "Q9"}, "--quant"),
            "no base folder": ({"base": str(self.base) + "x"}, "does not exist"),
            "an adapter from another base": ({"base": str(self.llama)}, "tuned from"),
            "no licence": ({"licence": " "}, "licence"),
            "no llama.cpp": ({"llama_cpp": str(self.llama / "nowhere")}, "llama.cpp is built"),
        }
        for what, (change, words) in cases.items():
            with self.subTest(what):
                self.assertTrue(any(words in p for p in self.problems(**change)), self.problems(**change))

    def test_the_reserve_stays_free(self):
        with mock.patch.object(em, "free_gb", return_value=30.0):
            self.assertTrue(any("reserve" in p or "must stay free" in p for p in em.check(self.cfg)))

    def test_the_commands_are_lists_never_shell_strings(self):
        plan = em.steps(self.cfg)
        self.assertEqual([s["name"] for s in plan], ["merge", "convert", "quantize"])
        for step in plan[1:]:
            self.assertTrue(all(isinstance(a, str) for a in step["argv"]))
        self.assertEqual(plan[2]["argv"][-1], "Q4_K_M")
        self.assertEqual(self.cfg.gguf.name, "kronos-v1-Q4_K_M.gguf")

    def test_every_base_file_is_hashed_but_not_the_cache(self):
        files = em.base_files(self.base)
        self.assertEqual(sorted(files), ["LICENSE", "README.md", "config.json", "model.safetensors"])
        self.assertEqual(files["model.safetensors"], qt.sha256_file(self.base / "model.safetensors"))

    def test_the_record_holds_what_d102_asks_for(self):
        record = em.model_record(self.cfg, self.run_record, {"LICENSE": "h"}, "g" * 64, "r" * 64)
        base = record["base_model"]
        self.assertEqual((base["name"], base["version"], base["licence"]), ("Some-Coder-7B", "1.0", "Apache-2.0"))
        self.assertEqual(record["file"]["sha256"], "g" * 64)
        self.assertEqual(record["prompt_sha256"], md.digest(md.SYSTEM))
        self.assertFalse(record["trained_on_real_data"])
        real = em.model_record(em.Export(**{**self.cfg.__dict__, "real_data": True}), self.run_record, {}, "g", "r")
        self.assertTrue(real["trained_on_real_data"])
        self.assertIn("never leaves this PC", real["note"])
        json.dumps(record)

    def test_a_whole_export_with_stand_in_commands(self):
        """The orchestration, with the model steps replaced: every step runs, in order, and the record is written."""
        calls = []

        def fake_run(argv, **kwargs):
            calls.append(Path(argv[1]).name if argv[0] == sys.executable else Path(argv[0]).name)
            target = Path(argv[argv.index("--outfile") + 1]) if "--outfile" in argv else Path(argv[2])
            target.write_bytes(b"gguf")
            return subprocess.CompletedProcess(argv, 0)

        with mock.patch.object(em.subprocess, "run", side_effect=fake_run), \
                mock.patch.object(em, "merge_adapter", side_effect=lambda cfg: calls.append("merge")):
            record = em.run(self.cfg)
        self.assertEqual(calls[0], "merge")
        self.assertEqual(calls[1], "convert_hf_to_gguf.py")
        self.assertTrue(calls[2].startswith("llama-quantize"))
        out = Path(self.cfg.out)
        self.assertTrue((out / "kronos-v1-Q4_K_M.gguf").is_file())
        self.assertEqual((out / "vcl.gbnf").read_bytes(), (ROOT / "train" / "vcl.gbnf").read_bytes())
        self.assertFalse((out / "work").exists(), "the 16-bit copies are removed")
        written = json.loads((out / "model.json").read_text(encoding="utf-8"))
        self.assertEqual(written["file"]["sha256"], qt.sha256_file(out / "kronos-v1-Q4_K_M.gguf"))
        self.assertEqual(written["grammar_sha256"], qt.sha256_file(out / "vcl.gbnf"))
        self.assertEqual(record["base_model"]["files_sha256"], em.base_files(self.base))

    def test_a_failing_step_stops_and_says_which(self):
        with mock.patch.object(em, "merge_adapter"), \
                mock.patch.object(em.subprocess, "run", return_value=subprocess.CompletedProcess([], 3)):
            with self.assertRaisesRegex(qt.TrainError, "convert failed"):
                em.run(self.cfg)
        self.assertFalse((Path(self.cfg.out) / "model.json").exists())

    def test_without_go_it_only_plans(self):
        args = ["--base", self.cfg.base, "--adapter", self.cfg.adapter, "--out", self.cfg.out, "--llama-cpp",
                self.cfg.llama_cpp, "--base-name", "n", "--base-version", "1", "--licence", "MIT"]
        out = io.StringIO()
        with contextlib.redirect_stdout(out), mock.patch.object(em, "run") as ran:
            code = em.main(args)
        self.assertEqual(code, 0)
        self.assertIn("dry run", out.getvalue())
        ran.assert_not_called()


if __name__ == "__main__":
    unittest.main()
