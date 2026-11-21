// KRONOS Offline's wizard (D86, D87, D89): 1 Ask, 2 Questions, 3 Review, 4 Save, as KRONOS.exe has them, with the AI
// step done by copy and paste. The work is done by worker.js (Python in the browser), one job at a time; Cancel stops
// that worker, starts a fresh one and rebuilds the requests from the actions this page kept. Plain JavaScript that
// Chrome 109 and Firefox 115 run, with no framework and no build step.
(function () {
  'use strict'

  const oldBrowser = document.getElementById('old-browser')   // index.html hides it where KRONOS can run
  if (oldBrowser) oldBrowser.parentNode.removeChild(oldBrowser)

  const $ = (id) => document.getElementById(id)
  const working = $('working')
  const workingStep = $('working-step')

  const EXAMPLE = 'toolchain/vcl/examples/example1_rev3.vcl'   // the prompt's Example 1, the same file KRONOS.exe uses
  const STEP_WORDS = {
    checking: 'Checking the recipe against the rules',
    testing: 'Running its practice tests',
    building: 'Building the macro and checking its VBA',
    reading: 'Reading your workbook',
    opening: 'Reading the recipe file',
    running: 'Making the changes on a copy',
    writing: 'Writing the changed copy',
    filling: 'Filling your Word template',
  }
  const MAX_BOOK = 25 * 1024 * 1024   // a bigger workbook may not fit in the browser's memory
  const MAX_RECIPE_FILE = 2 * 1024 * 1024
  const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  const XLSM = 'application/vnd.ms-excel.sheet.macroEnabled.12'
  const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  const MAX_TEMPLATE = 2 * 1024 * 1024   // a Word template may be up to 2 MB (vcl_docx.TEMPLATE_LIMIT)
  const PROMPT_KEY = 'kronos.orchestratorPrompt'
  // For the tests only: ?today=YYYY-MM-DD pins the date the practice tests use, as `vclc --today` does.
  const pinned = /[?&]today=(\d{4}-\d{2}-\d{2})(&|$)/.exec(window.location.search)
  const TODAY = pinned ? pinned[1] : null

  // ---------- the worker, one job at a time ----------

  let worker = null
  let ready = false
  let fatal = null
  let queue = []          // jobs not sent yet
  let job = null          // the job the worker is doing
  let jobs = 0

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

  // Ask the worker for something; words is what the working panel says (none for a quiet job). Resolves with
  // {json, data}; rejects with {message, details, fatal} or {cancelled: true}.
  function call(message, words) {
    return new Promise((resolve, reject) => {
      jobs += 1
      queue.push({ id: jobs, message, words: words || '', resolve, reject })
      next()
    })
  }

  function next() {
    if (job || !queue.length) return
    if (fatal) {
      const failed = queue
      queue = []
      failed.forEach((j) => j.reject(fatal))
      return
    }
    job = queue.shift()
    showWorking(!!job.words, ready ? job.words : 'Starting Python (the first time takes a few seconds)')
    if (ready) send()
  }

  function send() {
    job.sent = true
    const message = Object.assign({ id: job.id }, job.message)
    const transfer = message.data ? [message.data] : []
    worker.postMessage(message, transfer)
  }

  function heard(message) {
    if (message.type === 'status') {
      $('python').textContent = message.short
      if (job && job.words) workingStep.textContent = message.text
    } else if (message.type === 'ready') {
      ready = true
      $('python').textContent = `${message.python} in this browser (Pyodide ${message.pyodide}), ready`
      if (job && !job.sent) {
        if (job.words) workingStep.textContent = job.words
        send()
      }
    } else if (message.type === 'step') {
      if (job && message.id === job.id && job.words) workingStep.textContent = STEP_WORDS[message.name] || message.name
    } else if (message.type === 'error') {
      if (message.fatal) {
        fatal = message
        $('python').textContent = 'could not start'
      }
      if (!job) return
      const done = job
      job = null
      showWorking(false)
      done.reject(message)
      next()
    } else if (message.type === 'answer' && job && message.id === job.id) {
      const done = job
      job = null
      showWorking(false)
      done.resolve({ json: JSON.parse(message.json), data: message.data })
      next()
    }
  }

  function showWorking(on, text) {
    working.hidden = !on
    if (text) workingStep.textContent = text
    if (on) $('cancel').focus()
  }

  let cancels = 0   // counts Cancel clicks, so work that was waiting for a file knows it was cancelled

  function cancel() {
    cancels += 1
    worker.terminate()
    const stopped = [].concat(job ? [job] : [], queue)
    job = null
    queue = []
    showWorking(false)
    stopped.forEach((j) => j.reject({ cancelled: true }))
    if (!stopped.some((j) => j.words)) {   // nothing on the panel was working yet: say so where the person is
      say({ ask: 'file-note', question: 'question-note', review: 'review-note', save: 'apply-note' }[state.step],
        'Cancelled. Nothing was changed.')
    }
    startWorker()
    restore()
  }

  // ---------- the wizard's state ----------

  const state = {
    words: null,        // KRONOS's words, the Orchestrator prompt and its SHA-256 (kronos_web.words)
    step: 'ask',
    book: null,         // {name, bytes, info}
    ai: null,           // the id of the request's run, and what it was asked with: {id, request, book}
    shown: [],          // the run ids Review and Save show, in order
    views: {},          // id -> the wizard's view of that run
    messages: {},       // id -> {orchestrator, instructions}: the texts to copy, fetched while the click can wait
    ticks: {},          // id -> the assumptions are ticked
    actions: [],        // what this page asked the wizard to do, to rebuild it after Cancel
    copy: null,         // the changed workbook ready to download
    readBook: null,     // {name, bytes}: the workbook a recipe only reads (D120); never changed, never sent anywhere
    templates: {},      // id -> {name, bytes}: the Word template a DOCUMENT recipe fills (D121)
    filled: {},         // id -> {data, name}: the filled Word file ready to download
  }
  let runs = 0
  window.kronosViews = state.views

  function newId(kind) {
    runs += 1
    return `${kind}${runs}`
  }

  async function wizard(action, wid, args, words) {
    const answer = await call({ type: 'wizard', action, wid, args }, words)
    if (!answer.json.ok) throw { message: answer.json.message, plain: true }
    if (answer.json.view) state.views[wid] = answer.json.view
    if (['ask', 'recipe', 'paste', 'answer', 'problems'].indexOf(action) >= 0) state.actions.push([wid, action].concat(args))
    return answer.json
  }

  // After Cancel: a fresh wizard, then every request rebuilt from the actions kept.
  async function restore() {
    try {
      await call({ type: 'reset', today: TODAY })
      if (!state.words) loadWords()   // Cancel came before KRONOS's words did
      if (!state.actions.length) return
      const answer = await call({ type: 'wizard', action: 'restore', wid: '', args: [state.actions] },
        'Putting your requests back as they were')
      if (answer.json.ok) Object.assign(state.views, answer.json.views)
      state.messages = {}
      for (const id of Object.keys(state.views)) {
        if (state.views[id].status === 'waiting') await fetchMessages(id)
      }
      render()
    } catch (problem) {
      if (!problem.cancelled) say('ask-note', `KRONOS could not put your requests back: ${problem.message}. Start a new request.`)
    }
  }

  // ---------- small helpers ----------

  function el(tag, attrs, children) {
    const node = document.createElement(tag)
    Object.keys(attrs || {}).forEach((key) => {
      if (attrs[key] !== null && attrs[key] !== undefined && attrs[key] !== false) node.setAttribute(key, attrs[key])
    })
    ;[].concat(children || []).forEach((child) => {
      if (child === null || child === undefined || child === false) return
      node.appendChild(typeof child === 'string' || typeof child === 'number' ? document.createTextNode(String(child)) : child)
    })
    return node
  }

  function clear(node) {
    while (node.firstChild) node.removeChild(node.firstChild)
    return node
  }

  function table(headings, widths, rows, numbers) {
    const num = (i) => (numbers && numbers.indexOf(i) >= 0 ? { class: 'num' } : {})
    const cols = el('colgroup', {}, widths.map((w) => el('col', w ? { class: w } : {})))
    const head = el('thead', {}, el('tr', {}, headings.map((h, i) => el('th', Object.assign({ scope: 'col' }, num(i)), h))))
    const body = el('tbody', {}, rows.map((row) => el('tr', {}, row.map((cell, i) => el('td', num(i),
      cell && cell.nodeType ? cell : String(cell))))))
    return el('table', {}, [cols, head, body])
  }

  function banner(kind, label, words, id) {
    return el('div', { class: `banner ${kind}`, id }, [el('strong', {}, label), ` ${words}`])
  }

  function say(id, text) {
    $(id).textContent = text || ''
    $(id).hidden = !text
  }

  function problemText(problem) {
    if (problem.cancelled) return 'Cancelled. Nothing was changed.'
    if (problem.plain) return problem.message
    return problem.fatal ? `KRONOS's checker could not start in this browser: ${problem.message}. Close the black window, ` +
      'start KRONOS again from its folder, and use Chrome, Edge or Firefox.'
      : `KRONOS hit a problem: ${problem.message}. Nothing was changed; try again, and if it happens again, send this to Claude.`
  }

  function size(bytes) {
    return bytes >= 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`
  }

  function download(body, name, type) {
    const url = URL.createObjectURL(new Blob([body], { type }))
    const link = el('a', { href: url, download: name })
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.setTimeout(() => URL.revokeObjectURL(url), 30000)
  }

  // Copy inside the click itself (the clipboard needs the click); a box with the text selected is the fallback.
  function copyText(text, box, note) {
    window.kronosLastCopied = text
    const fallback = () => {
      box.value = text
      box.hidden = false
      box.focus()
      box.select()
      say(note, 'Press Ctrl+C to copy the selected text, then paste it into your AI chat.')
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => {
        box.hidden = true
        say(note, 'Copied. Paste it into your AI chat, then paste the AI\'s whole reply below.')
      }, fallback)
    } else {
      fallback()
    }
  }

  // ---------- the steps ----------

  const ORDER = ['ask', 'question', 'review', 'save']

  function go(step) {
    state.step = step
    ORDER.forEach((name) => { $(`step-${name}`).hidden = name !== step })
    ;[].forEach.call(document.querySelectorAll('.step'), (node) => {
      const at = ORDER.indexOf(node.getAttribute('data-step'))
      node.className = 'step' + (at === ORDER.indexOf(step) ? ' current' : at < ORDER.indexOf(step) ? ' done' : '')
    })
    render()
    window.scrollTo(0, 0)
  }

  function render() {
    if (state.step === 'ask') renderAsk()
    else if (state.step === 'question') renderQuestion()
    else if (state.step === 'review') renderReview()
    else renderSave()
  }

  // ---------- the panel that sends a message to the AI and takes its reply ----------

  // id: the run; where: the element to fill; label: what the copy buttons say they copy.
  function aiPanel(where, id, label) {
    clear(where)
    const texts = state.messages[id]
    if (!texts) return
    const note = `${where.id}-note`
    const box = el('textarea', { class: 'box copy-box', rows: '8', readonly: 'readonly', hidden: 'hidden' })
    const reply = el('textarea', { class: 'box reply', rows: '8', spellcheck: 'false', id: `${where.id}-reply` })
    const orchestrator = el('button', { type: 'button', class: 'copy-orchestrator' }, `Copy ${label} for the Orchestrator`)
    const instructions = el('button', { type: 'button', class: 'copy-instructions' }, `Copy ${label} with instructions (ChatGPT or a new chat)`)
    orchestrator.addEventListener('click', () => copyText(texts.orchestrator, box, note))
    instructions.addEventListener('click', () => copyText(texts.instructions, box, note))
    const check = el('button', { type: 'button', class: 'check-reply' }, 'Check the reply')
    check.addEventListener('click', () => pasteReply(id, reply, note))
    where.appendChild(el('div', { class: 'ai-panel' }, [
      el('h2', {}, 'Send this to your AI chat'),
      el('p', { class: 'muted small-gap' }, 'Copy the message, paste it into your AI chat, then copy the AI\'s whole ' +
        'reply and paste it here. Use "for the Orchestrator" in your claude.ai Orchestrator project (it already has ' +
        'KRONOS\'s instructions); use "with instructions" in ChatGPT or any new chat.'),
      el('div', { class: 'buttons' }, [orchestrator, instructions]),
      box,
      el('p', { class: 'note', id: note, hidden: 'hidden' }),
      el('label', { class: 'label', for: reply.id }, 'Paste the AI\'s whole reply'),
      reply,
      el('div', { class: 'buttons' }, check),
    ]))
  }

  async function fetchMessages(id) {
    const instructions = await wizard('message', id, [true])
    const orchestrator = await wizard('message', id, [false])
    state.messages[id] = { instructions: instructions.text, orchestrator: orchestrator.text }
  }

  async function pasteReply(id, box, note) {
    const text = box.value
    if (!text.trim()) {
      say(note, 'Paste the AI\'s whole reply first.')
      box.focus()
      return
    }
    try {
      await wizard('paste', id, [text], STEP_WORDS.checking)
      await settle(id)
    } catch (problem) {
      say(note, problemText(problem))
    }
  }

  // Where a run goes after the AI's reply: a question, more waiting (after a repair), or Review.
  async function settle(id) {
    const view = state.views[id]
    delete state.messages[id]
    if (view.status === 'waiting') await fetchMessages(id)
    if (view.status === 'question') {
      $('answer').value = ''
      go('question')
    } else {
      if (state.shown.indexOf(id) < 0) state.shown = [id]
      go('review')
    }
  }

  // ---------- 1 Ask ----------

  function renderAsk() {
    const words = state.words
    if (words) {
      $('request').placeholder = `For example: ${words.example}`
      const tips = clear($('tips'))
      words.tips.forEach((tip) => tips.appendChild(el('li', {}, tip)))
    }
    $('include-line').hidden = !(state.book && state.book.info)
    const panel = $('ai-panel-ask')
    const ai = state.ai
    if (ai && state.views[ai.id] && state.views[ai.id].status === 'waiting' && state.messages[ai.id]) {
      aiPanel(panel, ai.id, 'the request')
    } else {
      clear(panel)
      const write = el('button', { type: 'button', id: 'write-message' }, 'Write the message for the AI')
      write.addEventListener('click', writeMessage)
      panel.appendChild(el('div', { class: 'buttons' }, write))
    }
    promptNote()
  }

  function promptNote() {
    const old = $('prompt-note')
    if (old) old.parentNode.removeChild(old)
    const words = state.words
    if (!words) return
    let known = null
    try { known = window.localStorage.getItem(PROMPT_KEY) } catch (e) { known = null }
    if (known === words.prompt_sha256) return
    const box = el('textarea', { class: 'box copy-box', rows: '6', readonly: 'readonly', hidden: 'hidden' })
    const copy = el('button', { type: 'button', id: 'copy-prompt' }, 'Copy the Orchestrator instructions')
    copy.addEventListener('click', () => copyText(words.prompt, box, 'prompt-note-text'))
    const done = el('button', { type: 'button', id: 'prompt-done' }, 'I have updated them')
    done.addEventListener('click', () => {
      try { window.localStorage.setItem(PROMPT_KEY, words.prompt_sha256) } catch (e) { /* this browser keeps nothing */ }
      promptNote()
    })
    $('ask-card').insertBefore(el('div', { class: 'prompt-note', id: 'prompt-note' }, [
      el('strong', {}, 'If you use the claude.ai Orchestrator project: '),
      'KRONOS doesn\'t know whether its instructions are the newest. Copy them, open your Orchestrator project on ' +
        'claude.ai, open its project instructions, replace all of the text with what you copied, and save. Then click ' +
        '"I have updated them". (ChatGPT or a new chat needs nothing: use "with instructions".)',
      el('div', { class: 'buttons' }, [copy, done]),
      box,
      el('p', { class: 'note', id: 'prompt-note-text', hidden: 'hidden' }),
    ]), $('ai-panel-ask'))
  }

  async function writeMessage() {
    say('ask-note', '')
    const request = $('request').value
    const book = state.book && state.book.info && $('include-book').checked ? $('book-text').value : null
    const id = newId('a')
    try {
      await wizard('ask', id, [request, book], 'Writing the message')
      state.ai = { id, request, book }
      await fetchMessages(id)
      renderAsk()
    } catch (problem) {
      say('ask-note', problemText(problem))
    }
  }

  // ---------- the workbook ----------

  // What is wrong with a workbook file the person chose, in words, or '' when KRONOS can use it.
  function bookProblem(file) {
    const lower = file.name.toLowerCase()
    if (/\.xls$/.test(lower)) return 'This is an old .xls workbook. In Excel choose File, Save As, and save it as .xlsx; then choose that.'
    if (!/\.xls[xm]$/.test(lower)) return 'KRONOS works on Excel workbooks saved as .xlsx or .xlsm.'
    if (file.size > MAX_BOOK) return `This workbook is ${size(file.size)}. KRONOS can work on files up to ${size(MAX_BOOK)}.`
    return ''
  }

  async function chooseBook(file) {
    say('book-note', '')
    if (!file) return
    if (bookProblem(file)) {
      say('book-note', bookProblem(file))
      return
    }
    let bytes
    try {
      bytes = await file.arrayBuffer()
    } catch (error) {
      say('book-note', `KRONOS could not read the file (${error.message}).`)
      return
    }
    try {
      const answer = await call({ type: 'describe', data: bytes.slice(0), name: file.name }, STEP_WORDS.reading)
      window.kronosLastDescribed = JSON.stringify(answer.json)
      if (!answer.json.ok) {
        say('book-note', answer.json.message)
        return
      }
      state.book = { name: file.name, bytes, info: answer.json }
      state.copy = null
      showBook()
    } catch (problem) {
      say('book-note', problemText(problem))
    }
  }

  function showBook() {
    const book = state.book
    $('book-info').hidden = !book
    $('remove-book').hidden = !book
    $('choose-book').textContent = book ? 'Choose another workbook…' : 'Choose a workbook…'
    if (!book) {
      render()
      return
    }
    const info = book.info
    $('book-name').textContent = `${book.name} (${size(book.bytes.byteLength)})`
    clear($('book-sheets')).appendChild(table(['Sheet', 'Rows', 'Columns', 'Headings (row 1)'],
      ['col-where', 'col-short', 'col-short', ''],
      info.sheets.map((s) => [s.name, s.rows, s.cols, s.headings.filter(Boolean).join(', ')]), [1, 2]))
    const warn = clear($('book-warn'))
    if (info.features.length) {
      warn.appendChild(el('p', {}, [el('strong', {}, 'In this workbook: '), `${info.features.join(', ')}. KRONOS tells you ` +
        'what it cannot keep before you download a copy.']))
    }
    info.notes.forEach((text) => warn.appendChild(el('p', {}, text)))
    warn.hidden = !warn.firstChild
    $('book-text').value = info.text
    render()
  }

  function removeBook() {
    state.book = null
    state.copy = null
    $('book-file').value = ''
    say('book-note', '')
    showBook()
  }

  // ---------- recipe files ----------

  async function openRecipes(file) {
    say('file-note', '')
    clear($('recipe-list')).hidden = true
    if (!file) return
    if (file.size > MAX_RECIPE_FILE) {
      say('file-note', `This file is ${size(file.size)}. KRONOS reads recipe files up to ${size(MAX_RECIPE_FILE)}.`)
      return
    }
    try {
      const bytes = await file.arrayBuffer()
      const answer = await call({ type: 'recipes', data: bytes, name: file.name }, STEP_WORDS.opening)
      window.kronosLastRecipes = JSON.stringify(answer.json)
      const found = answer.json
      if (!found.ok) {
        say('file-note', found.message)
        return
      }
      if (found.recipes.length === 1) {
        await checkRecipes(found.recipes, file.name)
        return
      }
      showRecipeList(found, file.name)
    } catch (problem) {
      say('file-note', problemText(problem))
    }
  }

  function showRecipeList(found, fileName) {
    const list = clear($('recipe-list'))
    list.hidden = false
    list.appendChild(el('p', { class: 'details' }, `${found.message} Choose the ones to check:`))
    const boxes = found.recipes.map((recipe, i) => {
      const box = el('input', { type: 'checkbox', id: `recipe-${i}` })
      const lines = recipe.text.split('\n').length - 1
      list.appendChild(el('label', { class: 'recipe-row', for: `recipe-${i}` }, [box,
        el('span', { class: 'what' }, [`${i + 1}. ${recipe.title}`, el('br'), el('span', { class: 'small-muted' }, `${lines} lines`)])]))
      return box
    })
    const all = el('button', { type: 'button', id: 'choose-all' }, 'Choose all')
    all.addEventListener('click', () => boxes.forEach((b) => { b.checked = true }))
    const check = el('button', { type: 'button', id: 'check-chosen' }, 'Check the chosen recipes')
    check.addEventListener('click', () => {
      const chosen = found.recipes.filter((r, i) => boxes[i].checked)
      if (!chosen.length) {
        say('file-note', 'Tick at least one recipe, or click Choose all.')
        return
      }
      checkRecipes(chosen, fileName)
    })
    list.appendChild(el('div', { class: 'buttons' }, [all, check]))
  }

  async function checkRecipes(recipes, fileName) {
    say('file-note', '')
    const ids = []
    try {
      for (let i = 0; i < recipes.length; i += 1) {
        const id = newId('f')
        await wizard('recipe', id, [recipes[i].text, fileName],
          recipes.length > 1 ? `Checking recipe ${i + 1} of ${recipes.length}` : STEP_WORDS.checking)
        ids.push(id)
      }
    } catch (problem) {
      say('file-note', problemText(problem))
      if (!ids.length) return
    }
    state.shown = ids
    go('review')
  }

  async function loadExample() {
    say('file-note', '')
    const asked = cancels
    showWorking(true, 'Opening the practice example')   // at once: the file may wait behind Python's own downloads
    let text
    try {
      const response = await fetch(EXAMPLE, { cache: 'no-cache' })
      if (!response.ok) throw new Error(`HTTP ${response.status}`)
      text = await response.text()
    } catch (error) {
      if (!job) showWorking(false)
      say('file-note', `I couldn't load the practice example (${error.message}). Start KRONOS again from its folder.`)
      return
    }
    if (asked !== cancels) return   // Cancel was clicked while the file was on its way
    await checkRecipes([{ title: 'The practice example', text }], 'the practice example')
  }

  // ---------- 2 Questions ----------

  function renderQuestion() {
    const ai = state.ai
    const view = ai && state.views[ai.id]
    if (!view) return
    $('question-text').textContent = view.question || ''
    const panel = $('ai-panel-question')
    if (view.status === 'waiting' && state.messages[ai.id]) {
      aiPanel(panel, ai.id, 'the answer')
      return
    }
    clear(panel)
    const write = el('button', { type: 'button', id: 'write-answer' }, 'Write the answer message for the AI')
    write.addEventListener('click', async () => {
      say('question-note', '')
      try {
        await wizard('answer', ai.id, [$('answer').value], 'Writing the message')
        await fetchMessages(ai.id)
        renderQuestion()
      } catch (problem) {
        say('question-note', problemText(problem))
      }
    })
    panel.appendChild(el('div', { class: 'buttons' }, write))
  }

  // ---------- 3 Review ----------

  function renderReview() {
    const list = clear($('review-list'))
    state.shown.forEach((id) => list.appendChild(reviewSection(id)))
    updateGoSave()
  }

  function proven(id) {
    const view = state.views[id]
    return !!(view && view.best && view.best.proven && view.status !== 'waiting')
  }

  function updateGoSave() {
    const ready = state.shown.filter(proven)
    const unticked = ready.filter((id) => !state.ticks[id])
    $('go-save').disabled = !ready.length || unticked.length > 0
    say('review-note', !ready.length ? 'Save opens when a recipe is PROVEN.'
      : unticked.length ? 'Tick "These assumptions are right" for each PROVEN recipe to go on to Save.' : '')
  }

  function reviewSection(id) {
    const view = state.views[id]
    const best = view.best
    const card = el('section', { class: 'card review', id: `review-${id}` })
    const title = best ? best.title : view.request
    card.appendChild(el('h1', {}, title))
    if (view.status === 'waiting') {
      card.appendChild(banner('not-proven', 'WAITING FOR THE AI', view.newest === 'fix' || view.newest === 'repair'
        ? 'Send the problems to the AI and paste its new reply.' : 'Paste the AI\'s reply.', `banner-${id}`))
    } else if (best && best.proven) {
      card.appendChild(banner('proven', 'PROVEN', `It passed every check and ${best.tests.length} of ${best.tests.length} practice tests.`, `banner-${id}`))
    } else {
      card.appendChild(banner('not-proven', 'NOT PROVEN', view.note || 'The recipe did not pass its checks.', `banner-${id}`))
    }
    if (best) {
      const summary = best.summary
      card.appendChild(el('h2', {}, 'What it does'))
      card.appendChild(el('p', { class: 'details' }, summary.does || '(the recipe has no GOAL line)'))
      if (summary.steps.length) card.appendChild(el('ul', { class: 'plain-list' }, summary.steps.map((s) => el('li', {}, s))))
      if (summary.assumes.length) {
        card.appendChild(el('h2', {}, 'What it assumes'))
        card.appendChild(el('ul', { class: 'plain-list' }, summary.assumes.map((s) => el('li', {}, s))))
      }
      if (best.proven && view.status !== 'waiting') {
        const tick = el('input', { type: 'checkbox', id: `tick-${id}` })
        tick.checked = !!state.ticks[id]
        tick.addEventListener('change', () => { state.ticks[id] = tick.checked; updateGoSave() })
        card.appendChild(el('label', { class: 'checkline', for: `tick-${id}` }, [tick, 'These assumptions are right for my workbook']))
      }
      card.appendChild(el('h2', {}, 'What it changes'))
      card.appendChild(el('ul', { class: 'plain-list' }, (summary.changes.length ? summary.changes : ['nothing in the workbook'])
        .map((s) => el('li', {}, s))))
      if (summary.fills && summary.fills.length) {
        card.appendChild(el('h2', {}, 'What it fills'))
        card.appendChild(el('ul', { class: 'plain-list fills' }, summary.fills.map((s) => el('li', {}, s))))
      }
      if (best.ai_texts && best.ai_texts.length) {
        card.appendChild(el('h2', {}, 'Written by the AI, not proven: check it'))
        card.appendChild(el('p', { class: 'details' }, 'The tests cannot tell whether these words are right. Read them, and ' +
          'correct them in your file if they are not.'))
        card.appendChild(el('ul', { class: 'plain-list ai-texts' }, best.ai_texts.map((a) => el('li', {}, a.text))))
      }
      card.appendChild(el('h2', {}, 'What it never does'))
      card.appendChild(el('ul', { class: 'plain-list' }, (state.words ? state.words.never : []).map((s) => el('li', {}, s))))
      card.appendChild(el('h2', {}, 'How it was tested'))
      card.appendChild(table(['TEST', 'What it checks', 'Passed'], ['col-short', '', 'col-short'],
        best.tests.map((t) => [`Line ${t.line}`, t.checks, el('span', { class: t.passed ? 'yes' : 'no' }, t.passed ? 'Yes' : 'No')])))
      if (view.attempts.length) {
        card.appendChild(el('ul', { class: 'plain-list muted' }, view.attempts.map((a) => el('li', {},
          `Try ${a.number} (${a.source === 'ai' ? 'the AI' : 'the recipe'}): ${a.words}`))))
      }
      if (!best.proven) {
        const rows = best.findings.filter((f) => f.severity === 'error').map((f) => [f.stage === 'check' && f.line ? `Line ${f.line}` : f.stage, f.message, f.code])
        best.tests.filter((t) => !t.passed).forEach((t) => (t.why.length ? t.why : ['it failed']).forEach((w) => rows.push([`TEST on line ${t.line}`, w, ''])))
        if (rows.length) {
          card.appendChild(el('h2', {}, rows.length === 1 ? 'The problem' : `The problems (${rows.length})`))
          card.appendChild(table(['Where', 'Problem', 'Code'], ['col-where', '', 'col-code'], rows))
        }
      }
      card.appendChild(el('details', { class: 'more' }, [el('summary', {}, 'Show the recipe (advanced)'),
        el('pre', { class: 'recipe-text' }, best.spec)]))
    } else if (view.note) {
      card.appendChild(el('p', { class: 'details' }, view.note))
    }
    const panel = el('div', { id: `panel-${id}` })
    card.appendChild(panel)
    const note = el('p', { class: 'note', id: `review-note-${id}`, hidden: 'hidden' })
    if (view.status === 'waiting' && state.messages[id]) {
      aiPanel(panel, id, 'the problems')
    } else if (best && !best.proven) {
      const send = el('button', { type: 'button', class: 'send-problems' }, 'Send the problems to the AI')
      send.addEventListener('click', async () => {
        try {
          await wizard('problems', id, [], 'Writing the message')
          await fetchMessages(id)
          if (state.ai === null || state.ai.id !== id) state.ai = { id, request: view.request, book: null }
          renderReview()
        } catch (problem) {
          say(note.id, problemText(problem))
        }
      })
      panel.appendChild(el('div', { class: 'buttons' }, send))
    }
    card.appendChild(note)
    return card
  }

  // ---------- 4 Save ----------

  function renderSave() {
    const list = clear($('save-list'))
    const ids = state.shown.filter(proven)
    ids.forEach((id) => list.appendChild(saveSection(id)))
    const skipped = state.shown.length - ids.length
    if (skipped) list.appendChild(el('p', { class: 'details muted' }, `${skipped} recipe${skipped === 1 ? ' is' : 's are'} not ` +
      'PROVEN and not offered here.'))
    renderApply(ids.filter((id) => state.views[id].best.kind !== 'DOCUMENT'))   // D120: a DOCUMENT has its own card
    renderDocs(ids.filter((id) => state.views[id].best.kind === 'DOCUMENT'))
  }

  function saveSection(id) {
    const view = state.views[id]
    const card = el('section', { class: 'card save', id: `save-${id}` })
    card.appendChild(el('h1', {}, view.best.title))
    if (view.best.offline_only) {   // D118: a DOCUMENT or a recipe that reads a second workbook makes no macro
      card.appendChild(el('p', { class: 'details offline-only' }, 'This recipe runs in KRONOS Offline only, so there ' +
        'is no macro to download: use it on your files on this page.'))
      return card
    }
    const fields = el('div', { class: 'fields' })
    const inputs = {}
    if (view.starter && view.starter.length) {
      card.appendChild(el('p', { class: 'details' }, 'The macro comes with a starter macro that runs it on your real ' +
        'sheets. Type their real names:'))
      view.starter.forEach((arg, i) => {
        const fid = `start-${id}-${i}`
        const input = el('input', { class: 'box', type: 'text', id: fid, value: arg.default })
        inputs[arg.name] = input
        fields.appendChild(el('div', { class: 'field' }, [el('label', { for: fid },
          arg.type === 'Worksheet' ? `The sheet for ${arg.name}` : `${arg.name} (${arg.type})`), input]))
      })
    }
    card.appendChild(fields)
    const result = el('div', { id: `saved-${id}` })
    const note = el('p', { class: 'note', id: `save-note-${id}`, hidden: 'hidden' })
    const save = el('button', { type: 'button', class: 'download-bas' }, 'Download the macro (.bas)')
    save.addEventListener('click', async () => {
      say(note.id, '')
      const values = {}
      Object.keys(inputs).forEach((name) => { values[name] = inputs[name].value })
      try {
        const answer = await wizard('save', id, [values], 'Building the macro')
        const saved = answer.saved
        window.kronosLastSaved = saved
        download(saved.bas, saved.bas_name, 'text/plain')
        clear(result)
        result.appendChild(el('p', { class: 'details' }, `Saved ${saved.bas_name} to your Downloads folder. Then, in Excel:`))
        result.appendChild(el('ul', { class: 'steps-list' }, saved.steps.map((s) => el('li', {}, s))))
        if (saved.practice) {
          const practice = el('button', { type: 'button', class: 'download-practice' }, 'Download the practice run (.bas)')
          practice.addEventListener('click', () => download(saved.practice, saved.practice_name, 'text/plain'))
          result.appendChild(el('div', { class: 'buttons' }, practice))
          result.appendChild(el('ul', { class: 'steps-list' }, saved.practice_steps.map((s) => el('li', {}, s))))
        } else if (saved.practice_note) {
          result.appendChild(el('p', { class: 'details muted' }, `No practice run: ${saved.practice_note}.`))
        }
      } catch (problem) {
        say(note.id, problemText(problem))
      }
    })
    card.appendChild(el('div', { class: 'buttons' }, save))
    card.appendChild(note)
    card.appendChild(result)
    return card
  }

  // ---------- making the changes on a copy of the workbook ----------

  let applyArgs = {}   // id -> the recipe's arguments (kronos_web.arguments)

  async function renderApply(ids) {
    const card = $('apply-card')
    const book = state.book
    if (!book || !book.info || !ids.length) {
      card.hidden = true
      return
    }
    card.hidden = false
    say('apply-note', '')
    $('apply-lead').textContent = (ids.length === 1 ? `KRONOS will run this recipe on a copy of ${book.name}`
      : `KRONOS will run these ${ids.length} recipes one after another, in this order, on one copy of ${book.name}`) +
      ' and give you the changed copy. Your original file is not touched.'
    const fields = clear($('apply-fields'))
    for (const id of ids) {
      if (!applyArgs[id]) {
        try {
          applyArgs[id] = (await call({ type: 'arguments', text: state.views[id].best.spec })).json
        } catch (problem) {
          say('apply-note', problemText(problem))
          return
        }
      }
    }
    let files = false
    ids.forEach((id, n) => {
      const args = applyArgs[id]
      if (ids.length > 1) fields.appendChild(el('h2', {}, `${n + 1}. ${state.views[id].best.title}`))
      if (!args.ok || args.files) {
        files = true
        fields.appendChild(el('p', { class: 'details' }, 'This recipe works with files or folders, which KRONOS cannot do on an ' +
          'uploaded workbook. Use its macro instead.'))
        return
      }
      args.args.forEach((arg, i) => {
        const fid = `arg-${id}-${i}`
        let input
        if (arg.type === 'Worksheet') {
          input = el('select', { id: fid, class: 'box', 'data-recipe': id, 'data-arg': arg.name })
          const names = book.info.sheets.map((s) => s.name)
          const chosen = names.filter((name) => name.toLowerCase() === String(arg.default).toLowerCase())[0]
          if (!chosen) input.appendChild(el('option', { value: '' }, 'Choose a sheet…'))
          names.forEach((name) => input.appendChild(el('option', { value: name, selected: name === chosen ? 'selected' : null }, name)))
        } else if (arg.type === 'Workbook' && arg.reads) {   // D120: a second workbook, only read
          const chosen = state.readBook ? `${state.readBook.name} (${size(state.readBook.bytes.byteLength)})` : 'none chosen yet'
          const choose = el('button', { type: 'button', class: 'choose-read' },
            state.readBook ? 'Choose another…' : 'Choose the workbook to read from…')
          choose.addEventListener('click', () => $('read-file').click())
          input = el('div', { class: 'static', 'data-recipe': id, 'data-arg': arg.name, 'data-reads': 'yes' },
            [el('span', { class: 'read-name' }, chosen), ' ', choose])
        } else if (arg.type === 'Workbook') {
          input = el('div', { class: 'static', 'data-recipe': id, 'data-arg': arg.name }, 'the whole workbook')
        } else {
          input = el('input', { id: fid, class: 'box', type: 'text', 'data-recipe': id, 'data-arg': arg.name, value: arg.default })
        }
        const label = arg.type === 'Worksheet' ? `Sheet for ${arg.name}`
          : (arg.type === 'Workbook' ? (arg.reads ? `Workbook to read from (${arg.name})` : arg.name) : `${arg.name} (${arg.type})`)
        fields.appendChild(el('div', { class: 'field' }, [el('label', { for: fid }, label), input]))
      })
    })
    if (ids.some((id) => applyArgs[id].ok && applyArgs[id].args.some((arg) => arg.reads))) {
      fields.appendChild(el('p', { class: 'details muted' }, 'KRONOS only reads the workbook to read from, as Excel last ' +
        'saved it. It is never changed, and it stays on this computer.'))
    }
    $('apply').hidden = files
  }

  async function chooseReadBook(file) {   // D120
    say('apply-note', '')
    if (!file) return
    if (bookProblem(file)) {
      say('apply-note', bookProblem(file))
      return
    }
    try {
      state.readBook = { name: file.name, bytes: await file.arrayBuffer() }
    } catch (error) {
      say('apply-note', `KRONOS could not read the file (${error.message}).`)
      return
    }
    state.copy = null
    $('read-file').value = ''
    renderSave()
  }

  async function apply() {
    say('apply-note', '')
    const ids = state.shown.filter(proven)
    const args = ids.map(() => ({}))
    let missing = null
    ;[].forEach.call($('apply-fields').querySelectorAll('[data-arg]'), (input) => {
      if (input.tagName === 'DIV') {
        if (input.getAttribute('data-reads') && !state.readBook) missing = 'the workbook to read from'
        return
      }
      const at = ids.indexOf(input.getAttribute('data-recipe'))
      args[at][input.getAttribute('data-arg')] = input.value
      if (input.tagName === 'SELECT' && !input.value) missing = input.getAttribute('data-arg')
    })
    if (missing) {
      say('apply-note', missing === 'the workbook to read from' ? 'Choose the workbook to read from first.'
        : `Choose a sheet for ${missing} first.`)
      return
    }
    const reads = ids.some((id) => applyArgs[id].ok && applyArgs[id].args.some((arg) => arg.reads))
    $('applied').hidden = true
    state.copy = null
    try {
      const answer = await call({ type: 'apply', texts: ids.map((id) => state.views[id].best.spec), data: state.book.bytes.slice(0),
        name: state.book.name, args, today: TODAY,
        read: reads ? { data: state.readBook.bytes.slice(0), name: state.readBook.name } : null }, STEP_WORDS.checking)
      window.kronosLastApplied = JSON.stringify(answer.json)
      showApplied(answer.json, answer.data)
    } catch (problem) {
      say('apply-note', problemText(problem))
    }
  }

  // ---------- filling a Word template (D121) ----------

  let docFor = null   // the recipe whose template the person is choosing

  async function renderDocs(ids) {
    const list = clear($('doc-list'))
    for (const id of ids) {
      if (!applyArgs[id]) {
        try {
          applyArgs[id] = (await call({ type: 'arguments', text: state.views[id].best.spec })).json
        } catch (problem) {
          list.appendChild(el('p', { class: 'note' }, problemText(problem)))
          return
        }
      }
      list.appendChild(docCard(id))
    }
  }

  function docCard(id) {
    const view = state.views[id]
    const args = applyArgs[id].ok ? applyArgs[id].args : []
    const card = el('section', { class: 'card doc', id: `doc-${id}` })
    card.appendChild(el('h1', {}, `Fill a Word file: ${view.best.title}`))
    card.appendChild(el('p', { class: 'lead' }, 'KRONOS fills your Word template with this recipe and gives you a new Word ' +
      'file. Your template and your workbook are not touched.'))
    const fields = el('div', { class: 'fields' })
    const template = state.templates[id]
    const choose = el('button', { type: 'button', class: 'choose-template' },
      template ? 'Choose another…' : 'Choose the Word template…')
    choose.addEventListener('click', () => { docFor = id; $('doc-file').click() })
    fields.appendChild(el('div', { class: 'field' }, [el('label', {}, 'Word template (.docx)'),
      el('div', { class: 'static' }, [el('span', { class: 'template-name' },
        template ? `${template.name} (${size(template.bytes.byteLength)})` : 'none chosen yet'), ' ', choose])]))
    args.forEach((arg, i) => {
      const fid = `doc-arg-${id}-${i}`
      let input
      if (arg.type === 'Workbook') {
        input = el('div', { class: 'static', 'data-arg': arg.name, 'data-book': 'yes' },
          state.book ? `${state.book.name}, as Excel last saved it` : 'choose it under "Your workbook" in step 1, Ask')
      } else if (arg.type === 'Boolean') {
        input = el('select', { id: fid, class: 'box', 'data-arg': arg.name })
        ;['TRUE', 'FALSE'].forEach((v) => input.appendChild(el('option', {
          value: v, selected: String(arg.default).toUpperCase() === v ? 'selected' : null }, v === 'TRUE' ? 'Yes (TRUE)' : 'No (FALSE)')))
      } else {
        input = el('input', { id: fid, class: 'box', type: 'text', 'data-arg': arg.name, value: arg.default })
      }
      const label = arg.type === 'Workbook' ? `Workbook (${arg.name})` : `${arg.name} (${arg.type})`
      fields.appendChild(el('div', { class: 'field' }, [el('label', { for: fid }, label), input]))
    })
    card.appendChild(fields)
    const texts = view.best.ai_texts || []
    if (texts.length) {
      card.appendChild(el('p', { class: 'details' }, el('strong', {}, 'Written by the AI, not proven: read these in the ' +
        'Word file before you use it.')))
      card.appendChild(el('ul', { class: 'plain-list ai-texts' }, texts.map((t) => el('li', {}, t.text))))
    }
    const fill = el('button', { type: 'button', class: 'fill' }, 'Fill the Word file')
    fill.addEventListener('click', () => fillDocument(id))
    card.appendChild(el('div', { class: 'buttons' }, fill))
    card.appendChild(el('p', { class: 'note', id: `doc-note-${id}`, hidden: 'hidden' }))
    card.appendChild(el('section', { class: 'result', id: `filled-${id}`, hidden: 'hidden', 'aria-live': 'polite' }))
    return card
  }

  async function chooseTemplate(file) {
    const id = docFor
    $('doc-file').value = ''
    if (!file || !id) return
    say(`doc-note-${id}`, '')
    const lower = file.name.toLowerCase()
    if (/\.doc$/.test(lower)) {
      say(`doc-note-${id}`, 'This is an old .doc file. In Word choose File, Save As, and save it as .docx; then choose that.')
      return
    }
    if (!/\.docx$/.test(lower)) {
      say(`doc-note-${id}`, 'KRONOS fills Word templates saved as .docx.')
      return
    }
    if (file.size > MAX_TEMPLATE) {
      say(`doc-note-${id}`, `This template is ${size(file.size)}. A template may be up to ${size(MAX_TEMPLATE)}.`)
      return
    }
    try {
      state.templates[id] = { name: file.name, bytes: await file.arrayBuffer() }
    } catch (error) {
      say(`doc-note-${id}`, `KRONOS could not read the file (${error.message}).`)
      return
    }
    delete state.filled[id]
    renderSave()
  }

  async function fillDocument(id) {
    const card = $(`doc-${id}`)
    say(`doc-note-${id}`, '')
    const template = state.templates[id]
    if (!template) {
      say(`doc-note-${id}`, 'Choose the Word template first.')
      return
    }
    const needsBook = !!card.querySelector('[data-book]')
    if (needsBook && !state.book) {
      say(`doc-note-${id}`, 'This recipe reads a workbook: choose it under "Your workbook" in step 1, Ask, first.')
      return
    }
    const args = {}
    ;[].forEach.call(card.querySelectorAll('input[data-arg], select[data-arg]'), (input) => {
      args[input.getAttribute('data-arg')] = input.value
    })
    $(`filled-${id}`).hidden = true
    delete state.filled[id]
    try {
      const answer = await call({ type: 'fill', text: state.views[id].best.spec, template: template.bytes.slice(0),
        templateName: template.name, args, today: TODAY,
        book: needsBook ? { data: state.book.bytes.slice(0), name: state.book.name } : null }, STEP_WORDS.checking)
      window.kronosLastFilled = JSON.stringify(answer.json)
      showFilled(id, answer.json, answer.data)
    } catch (problem) {
      say(`doc-note-${id}`, problemText(problem))
    }
  }

  function showFilled(id, data, bytes) {
    const into = clear($(`filled-${id}`))
    into.hidden = false
    if (!data.ok) {
      into.appendChild(banner('not-proven', 'NOT FILLED', 'KRONOS made no Word file.', `filled-banner-${id}`))
      into.appendChild(el('p', { class: 'details' }, data.message))
      return
    }
    state.filled[id] = { data: bytes, name: `${data.stem} (KRONOS ${stamp()}).docx` }
    into.appendChild(banner('proven', 'WORD FILE READY', data.message, `filled-banner-${id}`))
    if (data.read_name) into.appendChild(el('p', { class: 'details' }, `It read ${data.read_name}, which is not changed.`))
    const sections = Object.keys(data.sections).map((name) => `${name}: ${data.sections[name]}`)
    into.appendChild(el('p', { class: 'details' }, `${data.fields.length} field${data.fields.length === 1 ? '' : 's'} filled` +
      (sections.length ? `; repeated sections (${sections.join(', ')})` : '') + `; ${data.tables} tables.`))
    into.appendChild(el('details', {}, [el('summary', {}, 'Show what each field says'),
      table(['Field', 'Text'], ['col-where', ''], data.fields.map((f) => [f.name, f.text]))]))
    let understand = null
    if (data.warnings.length) {
      const box = el('div', { class: 'warn' }, [el('p', {}, el('strong', {}, 'Before you download, please read:'))])
      data.warnings.forEach((w) => box.appendChild(el('p', {}, [el('strong', {}, 'Check: '), w.text])))
      into.appendChild(box)
      understand = el('input', { type: 'checkbox', class: 'understand' })
      into.appendChild(el('label', { class: 'checkline' }, [understand, 'I understand, and I want the Word file']))
    }
    const button = el('button', { type: 'button', class: 'download-doc' }, 'Download the Word file')
    button.disabled = !!understand
    if (understand) understand.addEventListener('change', () => { button.disabled = !understand.checked })
    button.addEventListener('click', () => {
      const filled = state.filled[id]
      if (filled) download(filled.data, filled.name, DOCX)
    })
    into.appendChild(el('div', { class: 'buttons' }, button))
    into.appendChild(el('p', { class: 'details muted' }, 'Open it in Word and read it through before you send it.'))
  }

  function stamp() {
    const d = new Date()
    const two = (n) => (n < 10 ? '0' : '') + n
    return `${d.getFullYear()}-${two(d.getMonth() + 1)}-${two(d.getDate())} ${two(d.getHours())}-${two(d.getMinutes())}`
  }

  const KINDS = { lost: 'Lost: ', 'saved again': 'Saved again: ', moved: 'Check: ', load: 'Note: ' }
  const STATUS = { changed: 'Changed', unchanged: 'Not changed', new: 'New sheet' }

  function sheetTable(sheets) {
    return table(['Sheet', 'What happened', 'Cells changed', 'Rows deleted', 'Rows added', 'Rows moved', 'Fills changed', 'Rows hidden'],
      ['col-where', 'col-short', '', '', '', '', '', ''],
      sheets.map((s) => [s.name, STATUS[s.status] || s.status, s.cells, s.rows_deleted, s.rows_added, s.rows_moved, s.fills, s.hidden]),
      [2, 3, 4, 5, 6, 7])
  }

  function showApplied(data, bytes) {
    const into = clear($('applied'))
    into.hidden = false
    if (!data.ok) {
      into.appendChild(banner('not-proven', 'NOT CHANGED', 'KRONOS made no copy.', 'applied-banner'))
      into.appendChild(el('p', { class: 'details' }, data.message))
      return
    }
    state.copy = { data: bytes, name: `${data.stem} (KRONOS ${stamp()}).${data.ext}`, type: data.ext === 'xlsm' ? XLSM : XLSX }
    into.appendChild(banner('proven', 'CHANGED COPY READY', data.message, 'applied-banner'))
    if (data.read_name) into.appendChild(el('p', { class: 'details' }, `It read ${data.read_name}, which is not changed.`))
    if (data.steps.length > 1) {
      data.steps.forEach((step, n) => {
        into.appendChild(el('h2', {}, `${n + 1}. ${step.recipe}`))
        if (step.result !== null) into.appendChild(el('p', { class: 'details' }, `It returned: ${step.result}`))
        step.messages.forEach((m) => into.appendChild(el('p', { class: 'details muted' }, m)))
        into.appendChild(sheetTable(step.sheets))
      })
      into.appendChild(el('h2', {}, 'All together'))
    } else if (data.result !== null) {
      into.appendChild(el('p', { class: 'details' }, `The recipe returned: ${data.result}`))
      data.messages.forEach((m) => into.appendChild(el('p', { class: 'details muted' }, m)))
      into.appendChild(el('h2', {}, 'What changed'))
    }
    into.appendChild(sheetTable(data.sheets))
    if (data.changes.length) {
      const more = data.changes_total > data.changes.length ? ` (${data.changes_total} cells changed in all)` : ''
      into.appendChild(el('details', {}, [el('summary', {}, `Show the first ${data.changes.length} changes${more}`),
        table(['Sheet', 'Cell', 'Before', 'After'], ['col-where', 'col-short', '', ''], data.changes)]))
    }
    let understand = null
    if (data.warnings.length) {
      const box = el('div', { class: 'warn', id: 'warnings' }, [el('p', {}, el('strong', {}, 'Before you download, please read:'))])
      data.warnings.forEach((w) => box.appendChild(el('p', {}, [el('strong', {}, KINDS[w.kind] || ''), w.text])))
      into.appendChild(box)
      understand = el('input', { type: 'checkbox', id: 'understand' })
      into.appendChild(el('label', { class: 'checkline' }, [understand, 'I understand, and I want the changed copy']))
    }
    const button = el('button', { type: 'button', id: 'download' }, 'Download the changed copy')
    button.disabled = !!understand
    if (understand) understand.addEventListener('change', () => { button.disabled = !understand.checked })
    button.addEventListener('click', () => { if (state.copy) download(state.copy.data, state.copy.name, state.copy.type) })
    into.appendChild(el('div', { class: 'buttons' }, button))
    into.appendChild(el('p', { class: 'details muted' }, 'Open the copy in Excel and look it over before you use it. Your original file is untouched.'))
  }

  function loadWords() {
    call({ type: 'words' }).then((answer) => {
      state.words = answer.json
      render()
    }, (problem) => { if (!problem.cancelled) say('ask-note', problemText(problem)) })
  }

  // ---------- starting over ----------

  function startOver() {
    state.ai = null
    state.shown = []
    state.ticks = {}
    state.messages = {}
    state.copy = null
    state.readBook = null
    state.templates = {}
    state.filled = {}
    applyArgs = {}
    $('request').value = ''
    $('applied').hidden = true
    clear($('recipe-list')).hidden = true
    say('ask-note', '')
    say('file-note', '')
    go('ask')
  }

  // ---------- the browser line, and off we go ----------

  function browserName() {
    const ua = navigator.userAgent
    const found = /Edg\/(\d+)/.exec(ua) || /Firefox\/(\d+)/.exec(ua) || /Chrome\/(\d+)/.exec(ua)
    if (!found) return ua
    const name = found[0].indexOf('Edg') === 0 ? 'Edge' : found[0].indexOf('Firefox') === 0 ? 'Firefox' : 'Chrome'
    return `${name} ${found[1]}`
  }

  $('browser').textContent = browserName()
  $('cancel').addEventListener('click', cancel)
  $('choose-book').addEventListener('click', () => $('book-file').click())
  $('read-file').addEventListener('change', (event) => { chooseReadBook(event.target.files[0]) })
  $('doc-file').addEventListener('change', (event) => { chooseTemplate(event.target.files[0]) })
  $('book-file').addEventListener('change', (event) => {
    const file = event.target.files && event.target.files[0]
    event.target.value = ''   // so that choosing the same file again still fires
    chooseBook(file)
  })
  $('remove-book').addEventListener('click', removeBook)
  $('choose-recipes').addEventListener('click', () => $('recipe-file').click())
  $('recipe-file').addEventListener('change', (event) => {
    const file = event.target.files && event.target.files[0]
    event.target.value = ''
    openRecipes(file)
  })
  $('load-example').addEventListener('click', loadExample)
  $('go-save').addEventListener('click', () => go('save'))
  $('apply').addEventListener('click', apply)
  $('start-over').addEventListener('click', startOver)
  ;[].forEach.call(document.querySelectorAll('.back'), (button) => {
    button.addEventListener('click', () => go(button.getAttribute('data-to')))
  })
  startWorker()
  call({ type: 'reset', today: TODAY }).catch(() => null)
  loadWords()
  go('ask')
})()
