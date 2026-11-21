"""The prompts the tuned model is trained on and, from March 2027, run with (D124).

They are short and fixed on purpose: the model learns VCL from its tuning, so it needs none of the long Orchestrator
prompt (`vcl/orchestrator_prompt.txt`), which is for Claude and Codex. The text a trained model sees must be the text
it is trained on, so KRONOS's model adapter (March 2027) builds its prompts with these two functions, never its own.

The model's whole reply is the recipe: no code fence, no explanation (the grammar-locked generator, `docs/ai-plan.md`,
allows nothing else). A repair is one turn, not a conversation: the request, the recipe that did not pass and the
toolchain's messages go in, and the whole corrected recipe comes out, so the loss is on the recipe alone.
"""

SYSTEM = ("You write VCL recipes for KGC KRONOS, an Excel automation tool for a pressure-vessel fabrication shop. "
          "Reply with the recipe only: no explanation and no code fence.")


def write_prompt(request):
    """The messages for 'write the recipe for this request'."""
    return [{"role": "system", "content": SYSTEM},
            {"role": "user", "content": request.strip()}]


def repair_prompt(request, broken, messages):
    """The messages for 'this recipe did not pass; here is what the toolchain said': messages are the lines
    `kronos_engine.feedback` gives, as vclc prints them."""
    body = ("Request:\n" + request.strip() + "\n\nRecipe that did not pass:\n" + broken.rstrip("\n") +
            "\n\nThe toolchain's messages:\n" + "\n".join(messages) + "\n\nWrite the whole corrected recipe.")
    return [{"role": "system", "content": SYSTEM},
            {"role": "user", "content": body}]
