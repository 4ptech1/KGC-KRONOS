// KRONOS Offline's checker (D86): Python in the browser (Pyodide 0.24.1) running the KRONOS toolchain, in a
// classic Web Worker so the page never freezes. Everything loads from this computer's own little server:
// Pyodide from pyodide/, the lark wheel from wheels/, and the repo's own compiler/, vcl/ and vba/ files from
// toolchain/, listed in toolchain/manifest.json. The toolchain answers through compiler/kronos_web.py, the same
// module the tests call in normal Python.
//
// Messages from the page, each with an id; the answer comes back with the same id as {type: 'answer', id, json} (and
// data, an ArrayBuffer, for a changed workbook), or as {type: 'error', id, message, details, fatal}:
//   {type: 'words'}                                   KRONOS's words and the Orchestrator prompt (kronos_web.words)
//   {type: 'reset', today}                            a fresh wizard (kronos_web.wizard_reset)
//   {type: 'wizard', action, wid, args}               one wizard action (kronos_web.wizard_call, D89)
//   {type: 'describe', data, name}                    an uploaded workbook, described (D87)
//   {type: 'recipes', data, name}                     the recipes in a file (kronos_web.read_recipe_file, D89)
//   {type: 'arguments', text}                         a recipe's arguments
//   {type: 'apply', texts, data, name, args, today, read}   recipes made on one copy of the workbook (apply_changes_many);
//                                     read: {data, name} of the workbook a recipe only reads (D120), or null
//   {type: 'fill', text, template, templateName, args, today, book}   a DOCUMENT fills a Word template (fill_document, D121)
//   {type: 'check', text, today}                      check one recipe (kronos_web.check)
// While working it sends {type: 'step', id, name}; before that, {type: 'status', text, short} and {type: 'ready', ...}.

// pyodide.asm.js is loaded here, before loadPyodide starts its big downloads (the Python library and the .wasm, 9 MB
// each). Pyodide would otherwise load it itself, synchronously, while those downloads wait unread, and a server that
// answers one request at a time (BlueForce's PowerShell server, which the work PC uses) would wait for them to be
// read before sending pyodide.asm.js: neither side could move (D86). Pyodide skips its own load when it is here.
importScripts('pyodide/pyodide.js', 'pyodide/pyodide.asm.js')

const ROOT = '/kronos'   // compiler/resources.py finds vcl/ and vba/ beside compiler/
let kronos = null
let failed = null

async function get(path) {
  const response = await fetch(path, { cache: 'no-cache' })
  if (!response.ok) throw new Error(`could not load ${path} (HTTP ${response.status})`)
  return response
}

async function bytes(path) {
  return new Uint8Array(await (await get(path)).arrayBuffer())
}

function status(text, short) {
  postMessage({ type: 'status', text, short })
}

async function start() {
  status('Starting Python', 'starting')
  const pyodide = await loadPyodide({ indexURL: new URL('pyodide/', self.location.href).href })
  status('Loading the KRONOS toolchain', 'loading the KRONOS toolchain')
  const manifest = await (await get('toolchain/manifest.json')).json()
  const site = pyodide.runPython('import site; site.getsitepackages()[0]')
  for (const wheel of manifest.wheels) {
    pyodide.unpackArchive(await bytes('wheels/' + wheel), 'wheel', { extractDir: site })
  }
  for (const file of manifest.files) {
    const path = `${ROOT}/${file}`
    pyodide.FS.mkdirTree(path.slice(0, path.lastIndexOf('/')))
    pyodide.FS.writeFile(path, await bytes('toolchain/' + file))
  }
  pyodide.runPython(`import sys; sys.path.insert(0, ${JSON.stringify(ROOT + '/compiler')})`)
  kronos = pyodide.pyimport('kronos_web')
  const python = pyodide.runPython('import sys; sys.version.split()[0]')
  postMessage({ type: 'ready', python, pyodide: pyodide.version, files: manifest.files.length })
}

function words(error) {
  const text = String((error && error.message) || error)
  const lines = text.trim().split('\n')
  return { message: lines[lines.length - 1], details: text }
}

const started = start().catch((error) => {
  failed = words(error)
  postMessage({ type: 'error', id: 0, fatal: true, message: failed.message, details: failed.details })
})

const JOBS = ['words', 'reset', 'wizard', 'describe', 'recipes', 'arguments', 'apply', 'check', 'fill', 'find', 'adapt']

self.onmessage = async (event) => {
  const job = event.data
  if (!job || JOBS.indexOf(job.type) < 0) return
  await started
  if (failed) {
    postMessage({ type: 'error', id: job.id, fatal: true, message: failed.message, details: failed.details })
    return
  }
  try {
    const step = (name) => postMessage({ type: 'step', id: job.id, name: String(name) })
    const answer = (json) => postMessage({ type: 'answer', id: job.id, json })
    if (job.type === 'words') {
      answer(kronos.words())
    } else if (job.type === 'reset') {
      answer(kronos.wizard_reset('/kronos-data', job.today || null))
    } else if (job.type === 'wizard') {
      answer(kronos.wizard_call(job.action, job.wid, JSON.stringify(job.args || []), step))
    } else if (job.type === 'describe') {
      step('reading')   // openpyxl, the workbook reader, is imported the first time a workbook is used
      answer(kronos.describe_book(new Uint8Array(job.data), job.name))
    } else if (job.type === 'recipes') {
      step('opening')
      answer(kronos.read_recipe_file(new Uint8Array(job.data), job.name))
    } else if (job.type === 'arguments') {
      answer(kronos.arguments(job.text))
    } else if (job.type === 'check') {
      answer(kronos.check(job.text, job.today || null, step))
    } else if (job.type === 'find') {   // D129: the recipe finder
      answer(kronos.find(job.request, JSON.stringify(job.headings || {})))
    } else if (job.type === 'adapt') {
      answer(job.propose ? kronos.propose(job.id, JSON.stringify(job.headings || {}))
        : kronos.adapt(job.id, JSON.stringify(job.headings || {}), JSON.stringify(job.accepted || []), job.today || null))
    } else if (job.type === 'fill') {   // D121: a DOCUMENT recipe on the person's Word template
      const result = kronos.fill_document(job.text, new Uint8Array(job.template), job.templateName, JSON.stringify(job.args || {}),
        job.today || null, job.book ? new Uint8Array(job.book.data) : null, job.book ? job.book.name : '', step)
      const parts = result.toJs()
      result.destroy()
      const data = parts[1].buffer.slice(parts[1].byteOffset, parts[1].byteOffset + parts[1].byteLength)
      postMessage({ type: 'answer', id: job.id, json: parts[0], data }, [data])
    } else {
      const result = kronos.apply_changes_many(JSON.stringify(job.texts), new Uint8Array(job.data), job.name,
        JSON.stringify(job.args || []), job.today || null, step,
        job.read ? new Uint8Array(job.read.data) : null, job.read ? job.read.name : '')
      const parts = result.toJs()   // [the answer's JSON text, the copy's bytes]
      result.destroy()
      const data = parts[1].buffer.slice(parts[1].byteOffset, parts[1].byteOffset + parts[1].byteLength)
      postMessage({ type: 'answer', id: job.id, json: parts[0], data }, [data])
    }
  } catch (error) {
    const problem = words(error)
    postMessage({ type: 'error', id: job.id, fatal: false, message: problem.message, details: problem.details })
  }
}
