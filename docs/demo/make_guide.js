// The KRONOS demo's Word guide (D88): KRONOS_Demo_NCR_Register.docx, for colleagues and managers.
//
//   python docs/demo/make_demo.py          (the workbook, the copies and facts.json)
//   python docs/demo/take_screenshots.py   (shots/, from the built KRONOS Offline)
//   node docs/demo/make_guide.js           (needs the npm package docx)
//
// Every number in the guide comes from facts.json, which make_demo.py writes from KRONOS's own runs.
'use strict'
const fs = require('fs')
const path = require('path')
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType,
  ShadingType, BorderStyle, LevelFormat, Footer, PageNumber, ImageRun, ExternalHyperlink,
} = require('docx')

const HERE = __dirname
const facts = JSON.parse(fs.readFileSync(path.join(HERE, 'facts.json'), 'utf-8'))
const RAW = 'https://github.com/4ptech1/KGC-KRONOS/raw/main/docs/demo/'
const NAVY = '0B2A6F', GREEN = '00870E', SAND = 'EFE6D4', GREY = 'F5F5F5', BAND = 'F7F2E8', RED = 'FFC7CE'
const FONT = 'Calibri'
const W = 9026   // A4 text width at 1" margins

const t = (s, o) => new TextRun(Object.assign({ text: s }, o || {}))
const b = (s) => t(s, { bold: true })
const p = (runs, opts) => new Paragraph(Object.assign({ spacing: { after: 120 },
  children: (Array.isArray(runs) ? runs : [t(runs)]) }, opts || {}))
const h1 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(s)] })
const h2 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(s)] })
const bullet = (runs) => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 60 },
  children: Array.isArray(runs) ? runs : [t(runs)] })
const step = (runs, ref) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60 },
  children: Array.isArray(runs) ? runs : [t(runs)] })
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }
const noBorders = { top: none, bottom: none, left: none, right: none }
const spacer = () => new Paragraph({ spacing: { after: 60 }, children: [] })

function box(children, fill) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [
    new TableCell({ width: { size: W, type: WidthType.DXA }, borders: noBorders,
      shading: { type: ShadingType.CLEAR, fill: fill || GREY, color: 'auto' },
      margins: { top: 160, bottom: 160, left: 240, right: 240 }, children })] })] })
}

// a table in BlueForce's look: a light header band, a thin line under each row, every other row sand
function table(headers, rows, widths, opts) {
  opts = opts || {}
  const line = { style: BorderStyle.SINGLE, size: 4, color: '000000' }
  const borders = { top: none, left: none, right: none, bottom: line }
  const size = opts.size || 18
  const cell = (text, i, fill, bold) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA }, borders,
    shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({ children: [t(String(text), { bold, size })] })] })
  return new Table({ width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: widths, rows: [
    new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, BAND, true)) }),
    ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(c, i,
      opts.fill ? opts.fill(r, ri) : (ri % 2 ? SAND : undefined), false)) })),
  ] })
}

function image(file, widthPx) {
  const data = fs.readFileSync(path.join(HERE, 'shots', file))
  const w = data.readUInt32BE(16), h = data.readUInt32BE(20)   // a PNG's size, from its header
  const width = widthPx || 600
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 160 }, children: [
    new ImageRun({ type: 'png', data, transformation: { width, height: Math.round(width * h / w) },
      altText: { title: file, description: file, name: file } })] })
}

const caption = (s) => p([t(s, { italics: true, size: 18, color: '595959' })], { alignment: AlignmentType.CENTER })
const link = (file) => new ExternalHyperlink({ link: RAW + file, children: [t(file, { style: 'Hyperlink' })] })

const runs = facts.runs
const [r1, r2, r3] = runs
const sheetOf = (run, name) => run.sheets.find((s) => s.name === name)
const reg1 = sheetOf(r1, facts.sheet), reg2 = sheetOf(r2, facts.sheet)
const pick = (row, cols) => cols.map((i) => row[i])
const twice = r1.highlighted.map((r) => r[0]).filter((n, i, all) => all.indexOf(n) !== i)
const chained = facts.chained_summary.summary
const chainedStatus = chained.slice(1, chained.findIndex((r) => r[0] === '')).map((r) => `${r[0]} ${r[1]}`).join(', ')
const COLS = [0, 1, 5, 6, 7]   // NCR No, Date Raised, Severity, Department, Status
const summaryRows = (rows) => rows.slice(1).filter((r) => r[0] !== '')
const recipeText = (name) => fs.readFileSync(path.join(HERE, name), 'utf-8').replace(/\r?\n$/, '').split(/\r?\n/)

const children = [
  new Paragraph({ heading: HeadingLevel.TITLE, children: [t('KRONOS demo: an NCR register')] }),
  p([t('Three everyday jobs on a register, done by KRONOS, in plain words', { size: 26, color: '595959' })],
    { spacing: { after: 240 } }),
  box([
    p([t('In short', { bold: true, color: NAVY, size: 24 })]),
    bullet('This demo gives KRONOS a made-up register of non-conformance reports (NCRs), and asks for three everyday jobs.'),
    bullet([b('Job 1: '), t(`colour the ${r1.result} NCRs that are Major and still Open, so they stand out.`)]),
    bullet([b('Job 2: '), t(`tidy the register: clean up the Status words, remove the ${r2.result} NCRs that were typed twice, and sort by date.`)]),
    bullet([b('Job 3: '), t('add a Summary sheet that counts the NCRs by status and by department.')]),
    bullet('For each job KRONOS proved the instructions on practice data first, then made the changes on a copy. The original file was never changed.'),
  ]),

  h1('1. What KRONOS is'),
  p('KRONOS helps people automate repetitive Excel jobs without writing code. You say what you want done; an AI chat ' +
    'turns that into a short, structured "recipe". KRONOS then does three things, all on your own PC, with nothing sent anywhere:'),
  bullet([b('Checks '), t('the recipe against its rules.')]),
  bullet([b('Tests '), t('it on made-up practice data, and shows PROVEN only when every test passes.')]),
  bullet([b('Makes the changes on a copy '), t('of your workbook, tells you exactly what changed, and lets you download the copy.')]),
  p([t('The AI only writes the recipe. KRONOS does the checking and the work, and it never changes your original file.')]),

  h1('2. The demo workbook'),
  p([t('The file is '), b(facts.workbook), t('. It holds a register of '), b(String(facts.rows)),
    t(' NCRs (non-conformance reports: records of things found not to meet the drawing, the standard or the order) on a sheet called '),
    b(facts.sheet), t('. Everything in it is invented: no job, item, defect or date is real.')]),
  table(['Column', 'What it holds'], [
    ['NCR No', 'The report number, such as NCR-2026-001'],
    ['Date Raised', 'The day the problem was reported'],
    ['Job No, Item', 'Which (invented) job and part it is about'],
    ['Description', 'What was found, for example "Porosity found by RT"'],
    ['Severity', 'Major or Minor'],
    ['Department', 'Who raised it: Welding, QA, Machining, Fabrication or Stores'],
    ['Status', 'Open, Closed or On Hold'],
    ['Date Closed', 'When it was closed, if it is'],
  ], [2400, 6626]),
  spacer(),
  p('Like many real registers, it is a little untidy, and that is on purpose:'),
  bullet('Some Status cells were typed in a hurry: " open", "OPEN ", "closed", "On hold ".'),
  bullet('Two NCRs were entered twice (NCR-2026-007 and NCR-2026-020).'),
  bullet('The rows are in the order they were typed, not in date order.'),
  p('The first rows look like this:'),
  table(['NCR No', 'Date Raised', 'Severity', 'Department', 'Status'],
    facts.first_rows.slice(1, 7).map((r) => pick(r, COLS).map((v) => (v === '' ? '' : `"${v}"`.replace(/^"(.*)"$/, (m, x) => (x !== x.trim() ? m : x))))),
    [1900, 1700, 1500, 1900, 2026]),
  spacer(),
  p('When you choose the workbook, KRONOS reads it and shows its sheets and columns. It also writes a short description you can give the AI chat, so the AI uses your real sheet and column names:'),
  image('workbook.png', 560),
  caption('KRONOS after choosing the demo workbook'),

  h1('3. How each job runs'),
  p('Every job follows the same steps in KRONOS:'),
  step('In step 1 (Ask), choose the workbook (once).', 'how'),
  step('Click Open a recipe file… and choose the recipe. KRONOS checks it straight away, runs its practice tests and shows step 3, Review.', 'how'),
  step('When it says PROVEN, tick "These assumptions are right" and click Go to Save.', 'how'),
  step('In Make the changes on a copy, pick the sheet (NCR Register), click the button, read what changed, and click Download the changed copy.', 'how'),
  image('recipe1-proven.png', 520),
  caption('Review: PROVEN, with what the recipe does, assumes and changes, and how it was tested'),
  p([t('The practice tests are small made-up sheets with the awkward cases built in (extra spaces, capital letters, a wrong sheet). ' +
    'As an extra check, we planted deliberate mistakes in each recipe; its tests caught every one that would make a difference.')]),

  h1('4. Job 1: highlight the open Major NCRs'),
  box([p([b('The request: '), t('"Colour every NCR that is Major and still Open, so I can see what needs attention first."')], { spacing: { after: 0 } })], SAND),
  h2('What the recipe tells KRONOS'),
  bullet('Find the columns called NCR No, Severity and Status (wherever they are).'),
  bullet('Go down the register row by row. Clear any old colour on the row first.'),
  bullet('If Severity is Major and Status is Open, ignoring capital letters and extra spaces, colour the whole row light red.'),
  bullet('Report how many rows were coloured. Change nothing else.'),
  h2('What KRONOS did'),
  p([t('It coloured '), b(`${r1.result} rows`), t(` (${reg1.fills} cells: ${r1.result} rows of 9 columns). No value changed. ` +
    'Note that rows typed as " open" or "OPEN " were recognised as Open too.')]),
  table(['NCR No', 'Severity', 'Department', 'Status (as typed)'], r1.highlighted.map((r) => r.map((v) => (v !== v.trim() ? `"${v}"` : v))),
    [2400, 1800, 2400, 2426], { fill: () => RED }),
  spacer(),
  p([t(`${twice.join(' and ')} appear${twice.length === 1 ? 's' : ''} twice because ${twice.length === 1 ? 'it was' : 'they were'} typed twice; Job 2 removes the copies.`, { italics: true })]),
  image('recipe1-copy.png', 560),
  caption('KRONOS\'s report for Job 1'),

  h1('5. Job 2: tidy the register'),
  box([p([b('The request: '), t('"Write every status the same way, remove NCRs that were entered twice, and sort the register with the newest first."')], { spacing: { after: 0 } })], SAND),
  h2('What the recipe tells KRONOS'),
  bullet('Rewrite each Status that is Open, Closed or On Hold, apart from capitals and spaces, exactly as Open, Closed or On Hold.'),
  bullet('Where an NCR No appears more than once, keep the first row and remove the later copies. A row with no NCR No is never removed.'),
  bullet('Sort all rows by Date Raised, newest first. Report how many copies were removed.'),
  h2('What KRONOS did'),
  p([t('It removed '), b(`${r2.result} repeated rows`), t(`, cleaned the Status words, and put ${reg2.rows_moved} rows into date order. ` +
    `The register now holds ${facts.rows - Number(r2.result)} NCRs, each once. Every other row and every other column is still there.`)]),
  p('The top of the register, after:'),
  table(['NCR No', 'Date Raised', 'Severity', 'Department', 'Status'], r2.first_rows_after.slice(1, 7).map((r) => pick(r, COLS)),
    [1900, 1700, 1500, 1900, 2026]),
  spacer(),
  image('recipe2-copy.png', 560),
  caption('KRONOS\'s report for Job 2'),

  h1('6. Job 3: add a summary sheet'),
  box([p([b('The request: '), t('"Add a sheet that counts the NCRs by status and by department, with a total."')], { spacing: { after: 0 } })], SAND),
  h2('What the recipe tells KRONOS'),
  bullet('Count every row that has an NCR No, by Status (treating " open" and "Open" as the same) and by Department.'),
  bullet('Add a new sheet called NCR Summary with the two small tables and a total. If the sheet is already there, change nothing.'),
  bullet('Leave the register itself exactly as it is.'),
  h2('What KRONOS did'),
  p([t('It added the sheet '), b('NCR Summary'), t(` and counted ${r3.result} NCRs. The register was not touched.`)]),
  table(['On the NCR Summary sheet', 'NCRs'], summaryRows(r3.summary), [5000, 2000]),
  spacer(),
  p([b('Tip: run Job 2 first. '), t(`The ${r3.result} above includes the 2 NCRs that were typed twice. Tidy first, then summarise the tidy copy, and the total is `),
    b(facts.chained_summary.result), t(` (${chainedStatus}). To do that, upload the copy Job 2 gave you and run Job 3 on it.`)]),
  image('recipe3-copy.png', 560),
  caption('KRONOS\'s report for Job 3: a new sheet, nothing else changed'),

  h1('7. Try it yourself'),
  p('All the demo files are in the project on GitHub. Click a name to download it:'),
  bullet([t('The workbook: '), link('NCR_Register_Demo.xlsx')]),
  bullet([t('The recipes: '), link('recipe_1_highlight.vcl'), t(', '), link('recipe_2_tidy.vcl'), t(', '), link('recipe_3_summary.vcl')]),
  bullet([t('What KRONOS gave back, to compare: '), link('NCR_Register_Demo_after_1.xlsx'), t(', '),
    link('NCR_Register_Demo_after_2.xlsx'), t(', '), link('NCR_Register_Demo_after_3.xlsx')]),
  p('Then, in KRONOS Offline:', { spacing: { before: 120, after: 60 } }),
  step('Start KRONOS (double-click Start KRONOS (Windows) in its folder).', 'try'),
  step('Click Choose a workbook… and pick NCR_Register_Demo.xlsx. You should see the sheets NCR Register (41 rows) and About.', 'try'),
  step('Click Open a recipe file… and choose a recipe file (or this Word guide itself: KRONOS finds all three recipes in it; click Choose all, then Check the chosen recipes). You should see step 3, Review, with the green PROVEN banner.', 'try'),
  step('Tick "These assumptions are right" and click Go to Save. There you can download the macro, and, in Make the changes on a copy, choose NCR Register as the sheet and click Make the changes on a copy.', 'try'),
  step('You should see CHANGED COPY READY with the same numbers as in this guide (all three at once give one copy, with what each did). Click Download the changed copy and open it in Excel.', 'try'),

  h1('8. Good to know'),
  bullet([b('Your original is never changed. '), t('KRONOS always works on a copy, inside your browser, and sends nothing anywhere.')]),
  bullet([b('Look the copy over. '), t('Open it in Excel before you rely on it. If a workbook holds something a copy cannot keep ' +
    '(such as pictures, or formulas after rows have moved), KRONOS says so and asks you to tick "I understand" first.')]),
  bullet([b('In real use, '), t('you describe the job in your own words to the AI chat, together with the description of your ' +
    'workbook that KRONOS writes, and paste the AI\'s recipe into KRONOS. The demo recipes were written that way, ahead of time.')]),
  bullet([b('Nothing here is real. '), t('The workbook, its jobs, items, defects and dates are invented for this demo.')]),

  h1('Appendix: the three recipes'),
  p([t('For the curious: this is what a recipe looks like. Each line starts with a keyword (GOAL, OP, TEST …); ' +
    'the ASSUME lines say in plain words what the recipe takes for granted, and the TEST lines are the practice tests.', { size: 20 })]),
  ...['recipe_1_highlight.vcl', 'recipe_2_tidy.vcl', 'recipe_3_summary.vcl'].flatMap((name) => [
    h2(name),
    ...recipeText(name).map((line) => new Paragraph({ spacing: { after: 0 }, children: [t(line, { font: 'Consolas', size: 13 })] })),
  ]),
]

const doc = new Document({
  creator: 'KGC KRONOS',
  title: 'KRONOS demo: an NCR register',
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { font: FONT, size: 44, bold: true, color: NAVY },
        paragraph: { spacing: { before: 0, after: 80 }, keepNext: true } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: NAVY },
        paragraph: { spacing: { before: 320, after: 120 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 24, bold: true, color: '000000' },
        paragraph: { spacing: { before: 200, after: 80 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    ...['how', 'try'].map((ref) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] })),
  ] },
  sections: [{
    properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      t('KRONOS demo: an NCR register (invented data)   ·   page ', { size: 18, color: '595959' }),
      new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '595959' })] })] }) },
    children,
  }],
})

Packer.toBuffer(doc).then((buffer) => {
  const out = path.join(HERE, 'KRONOS_Demo_NCR_Register.docx')
  fs.writeFileSync(out, buffer)
  console.log('wrote', out)
})
