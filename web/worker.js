// KRONOS Offline's checker (D86): Python in the browser (Pyodide 0.24.1) running the KRONOS toolchain, in a
// classic Web Worker so the page never freezes. Everything loads from this computer's own little server:
// Pyodide from pyodide/, the lark wheel from wheels/, and the repo's own compiler/, vcl/ and vba/ files from
// toolchain/, listed in toolchain/manifest.json. The toolchain answers through compiler/kronos_web.py, the same
// module the tests call in normal Python.
//
// Messages from the page: {type: 'check', id, text, today}, and for an uploaded workbook (D87, compiler/kronos_book.py)
// {type: 'describe', id, data, name}, {type: 'arguments', id, text} and {type: 'apply', id, text, data, name, args, today}
// (data is the workbook's ArrayBuffer). To the page: {type: 'status', text, short}, {type: 'ready', python, pyodide, files},
// {type: 'step', id, name}, {type: 'result', id, json}, {type: 'described', id, json}, {type: 'arguments', id, json},
// {type: 'applied', id, json, data} (data: the changed copy's ArrayBuffer, empty when there is none) and
// {type: 'error', id, message, details, fatal}.

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

self.onmessage = async (event) => {
  const job = event.data
  if (!job || ['check', 'describe', 'arguments', 'apply'].indexOf(job.type) < 0) return
  await started
  if (failed) {
    postMessage({ type: 'error', id: job.id, fatal: true, message: failed.message, details: failed.details })
    return
  }
  try {
    const step = (name) => postMessage({ type: 'step', id: job.id, name: String(name) })
    if (job.type === 'check') {
      const json = kronos.check(job.text, job.today || null, step)
      postMessage({ type: 'result', id: job.id, json })
    } else if (job.type === 'describe') {
      step('reading')   // openpyxl, the workbook reader, is imported the first time a workbook is used
      postMessage({ type: 'described', id: job.id, json: kronos.describe_book(new Uint8Array(job.data), job.name) })
    } else if (job.type === 'arguments') {
      postMessage({ type: 'arguments', id: job.id, json: kronos.arguments(job.text) })
    } else {
      const answer = kronos.apply_changes(job.text, new Uint8Array(job.data), job.name, JSON.stringify(job.args || {}),
        job.today || null, step)
      const parts = answer.toJs()   // [the answer's JSON text, the copy's bytes]
      answer.destroy()
      const data = parts[1].buffer.slice(parts[1].byteOffset, parts[1].byteOffset + parts[1].byteLength)
      postMessage({ type: 'applied', id: job.id, json: parts[0], data }, [data])
    }
  } catch (error) {
    const problem = words(error)
    postMessage({ type: 'error', id: job.id, fatal: false, message: problem.message, details: problem.details })
  }
}
