// KRONOS Offline's History (D145): the runs of this browser on this PC, newest first, kept in IndexedDB. Every read and
// write is in try/catch: where the browser keeps nothing (a private window, blocked site data) the list lasts until the
// page is closed and says so. A run is a record of what was asked and which recipe was used, never a file: no workbook
// and no Word file is stored. Plain JavaScript that Chrome 109 and Firefox 115 run.
(function () {
  'use strict'

  const DB = 'kronos-offline'
  const STORE = 'runs'
  const KEEP = 200   // the oldest runs beyond this are dropped
  const memory = []  // used only where IndexedDB is not there
  let mode = 'saved'

  function open() {
    return new Promise((resolve, reject) => {
      let request
      try { request = window.indexedDB.open(DB, 1) } catch (e) { reject(e); return }
      if (!request) { reject(new Error('no IndexedDB')); return }
      request.onupgradeneeded = () => { request.result.createObjectStore(STORE, { keyPath: 'id', autoIncrement: true }) }
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => reject(request.error)
      request.onblocked = () => reject(new Error('blocked'))
    })
  }

  function done(transaction, value) {
    return new Promise((resolve, reject) => {
      transaction.oncomplete = () => resolve(value)
      transaction.onerror = () => reject(transaction.error)
      transaction.onabort = () => reject(transaction.error)
    })
  }

  // all(): every run, newest first
  async function all() {
    try {
      const db = await open()
      const tx = db.transaction(STORE, 'readonly')
      const request = tx.objectStore(STORE).getAll()
      const rows = await new Promise((resolve, reject) => {
        request.onsuccess = () => resolve(request.result)
        request.onerror = () => reject(request.error)
      })
      db.close()
      mode = 'saved'
      return rows.sort((a, b) => b.id - a.id)
    } catch (e) {
      mode = 'memory'
      return memory.slice().reverse()
    }
  }

  // add(run): keeps one run: {when, kind, title, request, recipe, files}
  async function add(run) {
    const record = Object.assign({ when: new Date().toISOString() }, run)
    try {
      const db = await open()
      const tx = db.transaction(STORE, 'readwrite')
      tx.objectStore(STORE).add(record)
      await done(tx)
      const rows = await all()
      if (rows.length > KEEP) {
        const db2 = await open()
        const tx2 = db2.transaction(STORE, 'readwrite')
        rows.slice(KEEP).forEach((r) => tx2.objectStore(STORE).delete(r.id))
        await done(tx2)
        db2.close()
      }
      db.close()
      mode = 'saved'
    } catch (e) {
      mode = 'memory'
      record.id = memory.length + 1
      memory.push(record)
    }
  }

  async function clearAll() {
    try {
      const db = await open()
      const tx = db.transaction(STORE, 'readwrite')
      tx.objectStore(STORE).clear()
      await done(tx)
      db.close()
    } catch (e) { /* nothing was kept */ }
    memory.length = 0
  }

  // text(run): the run record a person can download, as plain text
  function text(run) {
    const lines = [`KRONOS Offline run, ${run.when}`, `What was done: ${run.kind}`, `Recipe: ${run.title}`]
    if (run.files && run.files.length) lines.push(`Files used: ${run.files.join(', ')}`)
    if (run.request) lines.push('', 'The request:', run.request)
    if (run.recipe) lines.push('', 'The recipe:', run.recipe)
    return lines.join('\n') + '\n'
  }

  window.kronosHistory = { all, add, clear: clearAll, text, mode: () => mode }
})()
