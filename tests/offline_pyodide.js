// KRONOS Offline (D86): run compiler/kronos_web.py under Pyodide in Node.js, from a built KRONOS-Offline/app folder,
// the way web/worker.js does in the browser: the same Pyodide files, the same lark wheel, the same toolchain files
// from the manifest. tests/test_offline.py compares every answer with normal Python's, as one string.
//
//   node tests/offline_pyodide.js <KRONOS-Offline/app>  < cases.json  > answers.json
//
// cases.json: [{"name": ..., "text": ..., "today": "YYYY-MM-DD"}], and for an uploaded workbook (D87) cases with a
// "kind": "describe" ({"data": <base64>, "file": "orders.xlsx"}), "arguments", or "apply" (also "args": {...}).
// answers.json: {"python": "3.11.3", "pyodide": "0.24.1", "answers": {name: {"json": <kronos_web's text>,
// "steps": [...], "data": <base64 of the changed copy, for apply>}}}.
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
    const step = (name) => steps.push(String(name))
    if (item.kind === 'describe') {
      answers[item.name] = { json: kronos.describe_book(new Uint8Array(Buffer.from(item.data, 'base64')), item.file), steps }
    } else if (item.kind === 'arguments') {
      answers[item.name] = { json: kronos.arguments(item.text), steps }
    } else if (item.kind === 'apply') {
      const answer = kronos.apply_changes(item.text, new Uint8Array(Buffer.from(item.data, 'base64')), item.file,
        JSON.stringify(item.args || {}), item.today, step)
      const parts = answer.toJs()
      answer.destroy()
      answers[item.name] = { json: parts[0], steps, data: Buffer.from(parts[1]).toString('base64') }
    } else {
      answers[item.name] = { json: kronos.check(item.text, item.today, step), steps }
    }
  }
  const python = pyodide.runPython('import sys; sys.version.split()[0]')
  process.stdout.write(JSON.stringify({ python, pyodide: pyodide.version, answers }))
}

main().catch((error) => {
  process.stderr.write(String((error && error.stack) || error) + '\n')
  process.exit(1)
})
