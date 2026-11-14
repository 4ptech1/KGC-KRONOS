// KRONOS Offline (D86): run compiler/kronos_web.py under Pyodide in Node.js, from a built KRONOS-Offline/app folder,
// the way web/worker.js does in the browser: the same Pyodide files, the same lark wheel, the same toolchain files
// from the manifest. tests/test_offline.py compares every answer with normal Python's, as one string.
//
//   node tests/offline_pyodide.js <KRONOS-Offline/app>  < cases.json  > answers.json
//
// cases.json: [{"name": ..., "text": ..., "today": "YYYY-MM-DD"}]. answers.json: {"python": "3.11.3",
// "pyodide": "0.24.1", "answers": {name: {"json": <kronos_web.check's text>, "steps": [...]}}}.
'use strict'
const { readFileSync } = require('fs')
const { join } = require('path')

const app = process.argv[2]
const { loadPyodide } = require(join(app, 'pyodide', 'pyodide.js'))

async function main() {
  const pyodide = await loadPyodide({ indexURL: join(app, 'pyodide') + '/' })
  const manifest = JSON.parse(readFileSync(join(app, 'toolchain', 'manifest.json'), 'utf-8'))
  const site = pyodide.runPython('import site; site.getsitepackages()[0]')
  for (const wheel of manifest.wheels) {   // Node's Buffer is refused here: it must be a plain Uint8Array
    pyodide.unpackArchive(new Uint8Array(readFileSync(join(app, 'wheels', wheel))), 'wheel', { extractDir: site })
  }
  for (const file of manifest.files) {
    const path = `/kronos/${file}`
    pyodide.FS.mkdirTree(path.slice(0, path.lastIndexOf('/')))
    pyodide.FS.writeFile(path, new Uint8Array(readFileSync(join(app, 'toolchain', file))))
  }
  pyodide.runPython("import sys; sys.path.insert(0, '/kronos/compiler')")
  const kronos = pyodide.pyimport('kronos_web')
  const cases = JSON.parse(readFileSync(0, 'utf-8'))
  const answers = {}
  for (const item of cases) {
    const steps = []
    const json = kronos.check(item.text, item.today, (name) => steps.push(String(name)))
    answers[item.name] = { json, steps }
  }
  const python = pyodide.runPython('import sys; sys.version.split()[0]')
  process.stdout.write(JSON.stringify({ python, pyodide: pyodide.version, answers }))
}

main().catch((error) => {
  process.stderr.write(String((error && error.stack) || error) + '\n')
  process.exit(1)
})
