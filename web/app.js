// KRONOS Offline's one screen (D86, D87): choose a workbook (optional), paste a recipe, check it in the browser, see
// PROVEN or the problems, and, with a workbook chosen, make the recipe's changes on a copy and download it. The work
// is done in worker.js; Cancel stops that worker and starts a fresh one. Plain JavaScript that Chrome 109 and
// Firefox 115 run, with no framework and no build step.
(function () {
  'use strict'

  const oldBrowser = document.getElementById('old-browser')   // index.html hides it where KRONOS can run
  if (oldBrowser) oldBrowser.parentNode.removeChild(oldBrowser)

  const $ = (id) => document.getElementById(id)
  const recipe = $('recipe')
  const checkButton = $('check')
  const exampleButton = $('load-example')
  const working = $('working')
  const workingStep = $('working-step')
  const note = $('note')
  const result = $('result')
  const applied = $('applied')

  const EXAMPLE = 'toolchain/vcl/examples/example1_rev3.vcl'   // the prompt's Example 1, the same file KRONOS.exe uses
  const STEP_WORDS = {
    checking: 'Checking the recipe against the rules',
    testing: 'Running its practice tests',
    building: 'Building the macro and checking its VBA',
    reading: 'Reading your workbook',
    running: 'Making the changes on a copy',
    writing: 'Writing the changed copy',
  }
  const MAX_BOOK = 25 * 1024 * 1024   // a bigger workbook may not fit in the browser's memory
  const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  const XLSM = 'application/vnd.ms-excel.sheet.macroEnabled.12'
  // For the tests only: ?today=YYYY-MM-DD pins the date the practice tests use, as `vclc --today` does.
  const pinned = /[?&]today=(\d{4}-\d{2}-\d{2})(&|$)/.exec(window.location.search)
  const TODAY = pinned ? pinned[1] : null

  let worker = null
  let ready = false
  let fatal = null
  let job = null          // {id, kind, message, sent, quiet}: the one thing the worker is doing
  let jobs = 0
  let book = null         // the chosen workbook: {name, bytes (ArrayBuffer), info (kronos_book.describe)}
  let proven = null       // the recipe last proven: {text, name, args, files}
  let copy = null         // the changed copy ready to download: {data (ArrayBuffer), name, type}

  // ---------- the worker ----------

  function startWorker() {
    ready = false
    fatal = null
    $('python').textContent = 'starting'
    worker = new Worker('worker.js')
    worker.onmessage = (event) => heard(event.data)
    worker.onerror = (event) => {
      event.preventDefault()
      heard({ type: 'error', id: 0, fatal: true, message: event.message || 'worker.js could not run', details: '' })
    }
  }

  // Start one job. words is what the working panel says first; a quiet job shows no panel.
  function begin(kind, message, words, quiet) {
    jobs += 1
    job = { id: jobs, kind, message, sent: false, quiet: !!quiet }
    if (!quiet) showWorking(true, ready ? words : 'Starting Python (the first time takes a few seconds)')
    if (fatal) {
      const problem = fatal
      job = null
      showWorking(false)
      problemHere(kind, problem)
    } else if (ready) {
      send()
    }
  }

  function send() {
    job.sent = true
    const message = Object.assign({ id: job.id }, job.message)
    if (job.kind === 'describe' || job.kind === 'apply') message.data = book.bytes.slice(0)   // the page keeps its own
    worker.postMessage(message)
  }

  function heard(message) {
    if (message.type === 'status') {
      $('python').textContent = message.short
      if (job && !job.quiet) workingStep.textContent = message.text
    } else if (message.type === 'ready') {
      ready = true
      $('python').textContent = `${message.python} in this browser (Pyodide ${message.pyodide}), ready`
      if (job && !job.sent) send()
    } else if (message.type === 'step') {
      if (job && message.id === job.id && !job.quiet) workingStep.textContent = STEP_WORDS[message.name] || message.name
    } else if (message.type === 'error') {
      if (message.fatal) {
        fatal = message
        $('python').textContent = 'could not start'
      }
      if (!job && !message.fatal) return
      const kind = job ? job.kind : 'check'
      job = null
      showWorking(false)
      problemHere(kind, message)
    } else if (job && message.id === job.id) {
      const kind = job.kind
      const text = job.message.text
      job = null
      if (kind !== 'arguments') showWorking(false)
      if (message.type === 'result') {
        window.kronosLastResult = message.json   // what the tests compare with normal Python's answer
        showChecked(JSON.parse(message.json), text)
      } else if (message.type === 'described') {
        window.kronosLastDescribed = message.json
        showBook(JSON.parse(message.json))
      } else if (message.type === 'arguments') {
        gotArguments(JSON.parse(message.json))
      } else if (message.type === 'applied') {
        window.kronosLastApplied = message.json
        showApplied(JSON.parse(message.json), message.data)
      }
    }
  }

  function problemHere(kind, message) {
    if (kind === 'describe') {
      book = null
      showBook(null)
      bookSay(`KRONOS could not read this workbook: ${message.message}`)
    } else if (kind === 'apply') {
      render(applied, problemParts(message, true))
    } else if (kind === 'check') {
      render(result, problemParts(message, false))
    }
  }

  // ---------- the buttons ----------

  function showWorking(on, text) {
    working.hidden = !on
    checkButton.disabled = on
    exampleButton.disabled = on
    $('apply').disabled = on
    recipe.readOnly = on
    if (text) workingStep.textContent = text
    if (on) $('cancel').focus()
  }

  function say(text) {
    note.textContent = text
    note.hidden = !text
  }

  function bookSay(text) {
    $('book-note').textContent = text
    $('book-note').hidden = !text
  }

  function applySay(text) {
    $('apply-note').textContent = text
    $('apply-note').hidden = !text
  }

  function check() {
    const text = recipe.value
    say('')
    if (!text.trim()) {
      say('Paste a recipe first, or click Load the practice example.')
      recipe.focus()
      return
    }
    if (fatal) {
      render(result, problemParts(fatal, false))
      return
    }
    result.hidden = true
    applied.hidden = true
    proven = null
    showApplyCard()
    begin('check', { type: 'check', text, today: TODAY }, STEP_WORDS.checking)
  }

  function cancel() {
    const kind = job ? job.kind : 'check'
    worker.terminate()
    job = null
    showWorking(false)
    if (kind === 'describe') {
      book = null
      showBook(null)
      bookSay('Cancelled. Nothing was changed.')
    } else if (kind === 'apply') {
      applySay('Cancelled. Nothing was changed.')
    } else {
      say('Cancelled. Nothing was changed.')
    }
    startWorker()
    checkButton.focus()
  }

  async function loadExample() {
    say('')
    try {
      const response = await fetch(EXAMPLE, { cache: 'no-cache' })
      if (!response.ok) throw new Error(`HTTP ${response.status}`)
      recipe.value = await response.text()
      result.hidden = true
      say('The practice example is in the box. Click Check the recipe.')
      checkButton.focus()
    } catch (error) {
      say(`I couldn't load the practice example (${error.message}). Start KRONOS again from its folder.`)
    }
  }

  // ---------- small helpers ----------

  function el(tag, attrs, children) {
    const node = document.createElement(tag)
    Object.keys(attrs || {}).forEach((key) => node.setAttribute(key, attrs[key]))
    ;[].concat(children || []).forEach((child) => {
      node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child)
    })
    return node
  }

  function table(headings, widths, rows, numbers) {
    const cols = el('colgroup', {}, widths.map((w) => el('col', w ? { class: w } : {})))
    const head = el('thead', {}, el('tr', {}, headings.map((h, i) => el('th', Object.assign({ scope: 'col' },
      numbers && numbers.indexOf(i) >= 0 ? { class: 'num' } : {}), h))))
    const body = el('tbody', {}, rows.map((row) => el('tr', {}, row.map((cell, i) => {
      const attrs = numbers && numbers.indexOf(i) >= 0 ? { class: 'num' } : {}
      return cell && cell.nodeType ? el('td', attrs, cell) : el('td', attrs, String(cell))
    }))))
    return el('table', {}, [cols, head, body])
  }

  function banner(kind, label, words, id) {
    return el('div', { class: `banner ${kind}`, id: id || 'banner' }, [el('strong', {}, label), ` ${words}`])
  }

  function render(into, parts) {
    while (into.firstChild) into.removeChild(into.firstChild)
    parts.forEach((part) => into.appendChild(part))
    into.hidden = false
  }

  function capital(text) {
    return text ? text.charAt(0).toUpperCase() + text.slice(1) : ''
  }

  function problemParts(message, applying) {
    const parts = [banner('problem', 'KRONOS HIT A PROBLEM', message.fatal
      ? 'Its checker could not start in this browser.' : (applying ? 'The copy stopped with an error.' : 'The check stopped with an error.'))]
    parts.push(el('p', { class: 'details' }, message.message))
    parts.push(el('p', { class: 'details muted' }, message.fatal
      ? 'Close the black window, start KRONOS again from its folder, and use Chrome, Edge or Firefox.'
      : 'Nothing was changed. Try again; if it happens again, send this message to Claude.'))
    if (message.details && message.details !== message.message) parts.push(el('pre', { class: 'text' }, message.details))
    return parts
  }

  function size(bytes) {
    return bytes >= 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`
  }

  // ---------- the workbook ----------

  async function chooseFile(file) {
    bookSay('')
    if (!file) return
    const lower = file.name.toLowerCase()
    if (/\.xls$/.test(lower)) {
      bookSay('This is an old .xls workbook. In Excel choose File, Save As, and save it as .xlsx; then choose that.')
      return
    }
    if (!/\.xls[xm]$/.test(lower)) {
      bookSay('KRONOS works on Excel workbooks saved as .xlsx or .xlsm.')
      return
    }
    if (file.size > MAX_BOOK) {
      bookSay(`This workbook is ${size(file.size)}. KRONOS can work on files up to ${size(MAX_BOOK)}.`)
      return
    }
    let bytes
    try {
      bytes = await file.arrayBuffer()
    } catch (error) {
      bookSay(`KRONOS could not read the file (${error.message}).`)
      return
    }
    book = { name: file.name, bytes, info: null }
    copy = null
    applied.hidden = true
    begin('describe', { type: 'describe', name: file.name }, STEP_WORDS.reading)
  }

  function showBook(info) {
    if (info && !info.ok) {
      bookSay(info.message)
      info = null
      book = null
    }
    $('book-info').hidden = !info
    $('remove-book').hidden = !book
    $('choose-book').textContent = book ? 'Choose another workbook…' : 'Choose a workbook…'
    if (!info) {
      showApplyCard()
      return
    }
    book.info = info
    $('book-name').textContent = `${book.name} (${size(book.bytes.byteLength)})`
    const holder = $('book-sheets')
    while (holder.firstChild) holder.removeChild(holder.firstChild)
    holder.appendChild(table(['Sheet', 'Rows', 'Columns', 'Headings (row 1)'], ['col-where', 'col-short', 'col-short', ''],
      info.sheets.map((s) => [s.name, s.rows, s.cols, s.headings.filter(Boolean).join(', ')]), [1, 2]))
    const warn = $('book-warn')
    while (warn.firstChild) warn.removeChild(warn.firstChild)
    if (info.features.length) {
      warn.appendChild(el('p', {}, [el('strong', {}, 'In this workbook: '), `${info.features.join(', ')}. KRONOS tells you what it cannot ` +
        'keep before you download a copy.']))
    }
    info.notes.forEach((text) => warn.appendChild(el('p', {}, text)))
    warn.hidden = !warn.firstChild
    $('book-text').value = info.text
    showApplyCard()
  }

  function removeBook() {
    book = null
    copy = null
    $('book-file').value = ''
    bookSay('')
    applied.hidden = true
    showBook(null)
  }

  function copyDescription() {
    const box = $('book-text')
    const fallback = () => {
      box.focus()
      box.select()
      bookSay('Press Ctrl+C to copy the selected text.')
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(box.value).then(() => bookSay('Copied. Paste it into your AI chat.'), fallback)
    } else {
      fallback()
    }
  }

  // ---------- the check's result ----------

  const WHERE = { check: 'The recipe', test: 'The practice tests', build: 'Building the macro', lint: 'The macro' }

  function where(finding) {
    if (finding.stage === 'check' && finding.line) return `Line ${finding.line}`
    if (finding.stage === 'lint' && finding.line) return `The macro, line ${finding.line}`
    return WHERE[finding.stage] || finding.stage
  }

  function problemRows(data) {
    const rows = []
    data.findings.filter((f) => f.severity === 'error').forEach((f) => rows.push([where(f), f.message, f.code]))
    data.tests.filter((t) => !t.passed).forEach((t) => {
      const failed = t.failed.length ? t.failed : [{ message: `RESULT=${t.result} ERROR_COUNT=${t.errors}`, code: '' }]
      failed.forEach((f) => rows.push([`TEST on line ${t.line}`, f.message, f.code]))
    })
    return rows
  }

  function showChecked(data, text) {
    const parts = []
    if (data.reply === 'question') {
      parts.push(banner('not-proven', 'A QUESTION', 'The AI asked a question instead of writing a recipe.'))
      parts.push(el('div', { class: 'question' }, data.question))
      parts.push(el('p', { class: 'details' }, 'Answer it in the AI chat, then paste the AI\'s new reply here.'))
    } else if (data.reply === 'none') {
      parts.push(banner('not-proven', 'NOT A RECIPE', `${capital(data.words)}.`))
      parts.push(el('p', { class: 'details' }, data.problems.join(' ')))
      parts.push(el('p', { class: 'details muted' }, 'Copy the AI\'s whole reply, with its vcl code block, and paste it here.'))
    } else if (data.proven) {
      parts.push(banner('proven', 'PROVEN', `${capital(data.words)}.`))
      const lines = data.bas.split('\r\n').length - 1
      parts.push(el('p', { class: 'details' }, `${data.kind === 'BUILD' ? 'The set of macros' : 'The macro'} ${data.name} ` +
        `is built: the module ${data.module}, ${lines} lines of VBA with KRONOS's helpers, lint-clean.`))
      parts.push(el('h2', {}, 'Its practice tests'))
      parts.push(table(['TEST', 'Result', 'Errors', 'Passed'], ['col-short', '', 'col-short', 'col-short'],
        data.tests.map((t) => [`Line ${t.line}`, t.result === null ? '(none)' : t.result, t.errors,
          el('span', { class: t.passed ? 'yes' : 'no' }, t.passed ? 'Yes' : 'No')])))
      const warnings = data.findings.filter((f) => f.severity === 'warning')
      if (warnings.length) {
        parts.push(el('h2', {}, 'Warnings (the recipe still passes)'))
        parts.push(table(['Where', 'Warning', 'Code'], ['col-where', '', 'col-code'], warnings.map((f) => [where(f), f.message, f.code])))
      }
      parts.push(el('p', { class: 'details muted' }, book
        ? 'Below you can make this recipe\'s changes on a copy of your workbook.'
        : 'To make this recipe\'s changes on a workbook of yours, choose the workbook at the top of the page. ' +
          'Asking the AI and saving the macro come in a later version.'))
    } else {
      parts.push(banner('not-proven', 'NOT PROVEN', `${capital(data.words)}.`))
      const rows = problemRows(data)
      if (rows.length) {
        parts.push(el('h2', {}, rows.length === 1 ? 'The problem' : `The problems (${rows.length})`))
        parts.push(table(['Where', 'Problem', 'Code'], ['col-where', '', 'col-code'], rows))
      }
      parts.push(el('h2', {}, 'What the AI would be told'))
      parts.push(el('pre', { class: 'text', id: 'problems' }, data.problems.join('\n')))
    }
    render(result, parts)
    if (data.reply === 'spec' && data.proven) {
      proven = { text, name: data.name, args: null, files: false }
      begin('arguments', { type: 'arguments', text }, '', true)
    } else {
      proven = null
      showApplyCard()
    }
  }

  // ---------- making the changes on a copy ----------

  function gotArguments(data) {
    if (!proven) return
    if (!data.ok) {
      proven = null
      showApplyCard()
      return
    }
    proven.args = data.args
    proven.files = data.files
    showApplyCard()
  }

  function showApplyCard() {
    const card = $('apply-card')
    if (!(proven && proven.args && book && book.info)) {
      card.hidden = true
      return
    }
    card.hidden = false
    applySay('')
    $('apply-lead').textContent = `KRONOS will run ${proven.name} on a copy of ${book.name} and give you the changed copy. ` +
      'Your original file is not touched.'
    const fields = $('apply-fields')
    while (fields.firstChild) fields.removeChild(fields.firstChild)
    $('apply').hidden = proven.files
    if (proven.files) {
      fields.appendChild(el('p', { class: 'details' }, 'This recipe works with files or folders, which KRONOS cannot do on an ' +
        'uploaded workbook. Use the macro it builds instead.'))
      return
    }
    proven.args.forEach((arg, index) => {
      const id = `arg-${index}`
      let input
      if (arg.type === 'Worksheet') {
        input = el('select', { id, class: 'box', 'data-arg': arg.name })
        const names = book.info.sheets.map((s) => s.name)
        const chosen = names.filter((n) => n.toLowerCase() === arg.default.toLowerCase())[0]
        if (!chosen) input.appendChild(el('option', { value: '' }, 'Choose a sheet…'))
        names.forEach((n) => {
          const option = el('option', { value: n }, n)
          if (n === chosen) option.setAttribute('selected', 'selected')
          input.appendChild(option)
        })
      } else if (arg.type === 'Workbook') {
        input = el('div', { class: 'static', 'data-arg': arg.name }, 'the whole workbook')
      } else {
        input = el('input', { id, class: 'box', type: 'text', 'data-arg': arg.name, value: arg.default })
      }
      const label = arg.type === 'Worksheet' ? `Sheet for ${arg.name}` : (arg.type === 'Workbook' ? arg.name : `${arg.name} (${arg.type})`)
      fields.appendChild(el('div', { class: 'field' }, [el('label', { for: id }, label), input]))
    })
  }

  function apply() {
    applySay('')
    if (!proven || !book || fatal) return
    const args = {}
    let missing = null
    ;[].forEach.call($('apply-fields').querySelectorAll('[data-arg]'), (input) => {
      if (input.tagName === 'DIV') return
      args[input.getAttribute('data-arg')] = input.value
      if (input.tagName === 'SELECT' && !input.value) missing = input.getAttribute('data-arg')
    })
    if (missing) {
      applySay(`Choose a sheet for ${missing} first.`)
      return
    }
    applied.hidden = true
    copy = null
    begin('apply', { type: 'apply', text: proven.text, name: book.name, args, today: TODAY }, STEP_WORDS.checking)
  }

  function stamp() {
    const d = new Date()
    const two = (n) => (n < 10 ? '0' : '') + n
    return `${d.getFullYear()}-${two(d.getMonth() + 1)}-${two(d.getDate())} ${two(d.getHours())}-${two(d.getMinutes())}`
  }

  const KINDS = { lost: 'Lost: ', 'saved again': 'Saved again: ', moved: 'Check: ', load: 'Note: ' }

  function showApplied(data, bytes) {
    const parts = []
    if (!data.ok) {
      parts.push(banner('not-proven', 'NOT CHANGED', 'KRONOS made no copy.', 'applied-banner'))
      parts.push(el('p', { class: 'details' }, data.message))
      render(applied, parts)
      return
    }
    copy = { data: bytes, name: `${data.stem} (KRONOS ${stamp()}).${data.ext}`, type: data.ext === 'xlsm' ? XLSM : XLSX }
    parts.push(banner('proven', 'CHANGED COPY READY', data.message, 'applied-banner'))
    if (data.result !== null) parts.push(el('p', { class: 'details' }, `The recipe returned: ${data.result}`))
    data.messages.forEach((m) => parts.push(el('p', { class: 'details muted' }, m)))
    const status = { changed: 'Changed', unchanged: 'Not changed', new: 'New sheet' }
    parts.push(el('h2', {}, 'What changed'))
    parts.push(table(['Sheet', 'What happened', 'Cells changed', 'Rows deleted', 'Rows added', 'Rows moved', 'Fills changed', 'Rows hidden'],
      ['col-where', 'col-short', '', '', '', '', '', ''],
      data.sheets.map((s) => [s.name, status[s.status] || s.status, s.cells, s.rows_deleted, s.rows_added, s.rows_moved, s.fills, s.hidden]),
      [2, 3, 4, 5, 6, 7]))
    if (data.changes.length) {
      const more = data.changes_total > data.changes.length ? ` (${data.changes_total} cells changed in all)` : ''
      parts.push(el('details', {}, [el('summary', {}, `Show the first ${data.changes.length} changes${more}`),
        table(['Sheet', 'Cell', 'Before', 'After'], ['col-where', 'col-short', '', ''], data.changes)]))
    }
    if (data.warnings.length) {
      const box = el('div', { class: 'warn', id: 'warnings' }, [el('p', {}, el('strong', {}, 'Before you download, please read:'))])
      data.warnings.forEach((w) => box.appendChild(el('p', {}, [el('strong', {}, KINDS[w.kind] || ''), w.text])))
      parts.push(box)
      parts.push(el('label', { class: 'checkline' }, [el('input', { type: 'checkbox', id: 'understand' }),
        'I understand, and I want the changed copy']))
    }
    parts.push(el('div', { class: 'buttons' }, el('button', { type: 'button', id: 'download' }, 'Download the changed copy')))
    parts.push(el('p', { class: 'details muted' }, 'Open the copy in Excel and look it over before you use it. Your original file is untouched.'))
    render(applied, parts)
    const download = $('download')
    const understand = $('understand')
    download.disabled = !!understand
    if (understand) understand.addEventListener('change', () => { download.disabled = !understand.checked })
    download.addEventListener('click', saveCopy)
  }

  function saveCopy() {
    if (!copy) return
    const url = URL.createObjectURL(new Blob([copy.data], { type: copy.type }))
    const link = el('a', { href: url, download: copy.name })
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.setTimeout(() => URL.revokeObjectURL(url), 30000)
  }

  // ---------- the browser line ----------

  function browserName() {
    const ua = navigator.userAgent
    const found = /Edg\/(\d+)/.exec(ua) || /Firefox\/(\d+)/.exec(ua) || /Chrome\/(\d+)/.exec(ua)
    if (!found) return ua
    const name = found[0].indexOf('Edg') === 0 ? 'Edge' : found[0].indexOf('Firefox') === 0 ? 'Firefox' : 'Chrome'
    return `${name} ${found[1]}`
  }

  $('browser').textContent = browserName()
  checkButton.addEventListener('click', check)
  exampleButton.addEventListener('click', loadExample)
  $('cancel').addEventListener('click', cancel)
  $('choose-book').addEventListener('click', () => $('book-file').click())
  $('book-file').addEventListener('change', (event) => {
    const file = event.target.files && event.target.files[0]
    event.target.value = ''   // so that choosing the same file again still fires
    chooseFile(file)
  })
  $('remove-book').addEventListener('click', removeBook)
  $('copy-book').addEventListener('click', copyDescription)
  $('apply').addEventListener('click', apply)
  startWorker()
})()
