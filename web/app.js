// KRONOS Offline, version 1's one screen (D86): paste a recipe, check it in the browser, see PROVEN or the
// problems. The checking runs in worker.js; Cancel stops that worker and starts a fresh one. Plain JavaScript
// that Chrome 109 and Firefox 115 run, with no framework and no build step.
(function () {
  'use strict'

  const oldBrowser = document.getElementById('old-browser')   // stays on screen only where this script can't run
  if (oldBrowser) oldBrowser.parentNode.removeChild(oldBrowser)

  const $ = (id) => document.getElementById(id)
  const recipe = $('recipe')
  const checkButton = $('check')
  const exampleButton = $('load-example')
  const working = $('working')
  const workingStep = $('working-step')
  const note = $('note')
  const result = $('result')

  const EXAMPLE = 'toolchain/vcl/examples/example1_rev3.vcl'   // the prompt's Example 1, the same file KRONOS.exe uses
  const STEP_WORDS = {
    checking: 'Checking the recipe against the rules',
    testing: 'Running its practice tests',
    building: 'Building the macro and checking its VBA',
  }
  // For the tests only: ?today=YYYY-MM-DD pins the date the practice tests use, as `vclc --today` does.
  const pinned = /[?&]today=(\d{4}-\d{2}-\d{2})(&|$)/.exec(window.location.search)
  const TODAY = pinned ? pinned[1] : null

  let worker = null
  let ready = false
  let fatal = null
  let job = null
  let jobs = 0

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

  function send() {
    job.sent = true
    worker.postMessage({ type: 'check', id: job.id, text: job.text, today: TODAY })
  }

  function heard(message) {
    if (message.type === 'status') {
      $('python').textContent = message.text.toLowerCase()
      if (job) workingStep.textContent = message.text
    } else if (message.type === 'ready') {
      ready = true
      $('python').textContent = `${message.python} in this browser (Pyodide ${message.pyodide}), ready`
      if (job && !job.sent) send()
    } else if (message.type === 'step') {
      if (job && message.id === job.id) workingStep.textContent = STEP_WORDS[message.name] || message.name
    } else if (message.type === 'result') {
      if (!job || message.id !== job.id) return
      job = null
      window.kronosLastResult = message.json   // what the tests compare with normal Python's answer
      showWorking(false)
      show(JSON.parse(message.json))
    } else if (message.type === 'error') {
      if (message.fatal) {
        fatal = message
        $('python').textContent = 'could not start'
      }
      if (!job && !message.fatal) return
      job = null
      showWorking(false)
      showProblem(message)
    }
  }

  // ---------- the buttons ----------

  function showWorking(on, text) {
    working.hidden = !on
    checkButton.disabled = on
    exampleButton.disabled = on
    recipe.readOnly = on
    if (text) workingStep.textContent = text
    if (on) $('cancel').focus()
  }

  function say(text) {
    note.textContent = text
    note.hidden = !text
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
      showProblem(fatal)
      return
    }
    result.hidden = true
    jobs += 1
    job = { id: jobs, text, sent: false }
    showWorking(true, ready ? STEP_WORDS.checking : 'Starting Python (the first time takes a few seconds)')
    if (ready) send()
  }

  function cancel() {
    worker.terminate()
    job = null
    showWorking(false)
    say('Cancelled. Nothing was changed.')
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

  // ---------- the result ----------

  function el(tag, attrs, children) {
    const node = document.createElement(tag)
    Object.keys(attrs || {}).forEach((key) => node.setAttribute(key, attrs[key]))
    ;[].concat(children || []).forEach((child) => {
      node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child)
    })
    return node
  }

  function table(headings, widths, rows) {
    const cols = el('colgroup', {}, widths.map((w) => el('col', w ? { class: w } : {})))
    const head = el('thead', {}, el('tr', {}, headings.map((h) => el('th', { scope: 'col' }, h))))
    const body = el('tbody', {}, rows.map((row) => el('tr', {}, row.map((cell) => {
      return cell && cell.nodeType ? el('td', {}, cell) : el('td', {}, String(cell))
    }))))
    return el('table', {}, [cols, head, body])
  }

  function banner(kind, label, words) {
    return el('div', { class: `banner ${kind}`, id: 'banner' }, [el('strong', {}, label), ` ${words}`])
  }

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

  function show(data) {
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
      parts.push(el('p', { class: 'details muted' }, 'This first version checks recipes only. Asking the AI and saving the macro come in the next version.'))
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
    render(parts)
  }

  function showProblem(message) {
    const parts = [banner('problem', 'KRONOS HIT A PROBLEM', message.fatal
      ? 'Its checker could not start in this browser.' : 'The check stopped with an error.')]
    parts.push(el('p', { class: 'details' }, message.message))
    parts.push(el('p', { class: 'details muted' }, message.fatal
      ? 'Close the black window, start KRONOS again from its folder, and use Chrome, Edge or Firefox.'
      : 'Nothing was changed. Try again; if it happens again, send this message to Claude.'))
    if (message.details && message.details !== message.message) parts.push(el('pre', { class: 'text' }, message.details))
    render(parts)
  }

  function render(parts) {
    while (result.firstChild) result.removeChild(result.firstChild)
    parts.forEach((part) => result.appendChild(part))
    result.hidden = false
  }

  function capital(text) {
    return text ? text.charAt(0).toUpperCase() + text.slice(1) : ''
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
  startWorker()
})()
