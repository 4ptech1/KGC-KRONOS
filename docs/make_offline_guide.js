// The plain-language guide to KRONOS Offline (docs/KRONOS_Offline_Guide.docx), for people who aren't technical.
//
//   node docs/make_offline_guide.js docs/KRONOS_Offline_Guide.docx      (needs the npm package docx)
//
// Facts only from docs/kronos-offline.md; no names or real data. Update it with each version handed over.
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, LevelFormat, Footer, PageNumber,
} = require('docx');

const NAVY = '0B2A6F', GREEN = '00870E', AMBER = 'B26A00', SAND = 'EFE6D4', GREY = 'F5F5F5';
const FONT = 'Calibri';
const W = 9026; // A4 text width at 1" margins (11906 - 2*1440)

const p = (text, opts = {}) => new Paragraph({
  spacing: { after: 120 }, ...opts,
  children: Array.isArray(text) ? text : [new TextRun(text)],
});
const b = (t) => new TextRun({ text: t, bold: true });
const t = (s) => new TextRun(s);
const h1 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(s)] });
const bullet = (runs) => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 60 },
  children: Array.isArray(runs) ? runs : [new TextRun(runs)] });
const step = (runs, ref = 'steps') => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60 },
  children: Array.isArray(runs) ? runs : [new TextRun(runs)] });

const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: none, bottom: none, left: none, right: none };

function box(children, fill = GREY) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, borders: noBorders,
      shading: { type: ShadingType.CLEAR, fill, color: 'auto' },
      margins: { top: 160, bottom: 160, left: 240, right: 240 },
      children,
    })] })],
  });
}

function table(headers, rows, widths) {
  const line = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
  const cellBorders = { top: none, left: none, right: none, bottom: line };
  const cell = (content, i, fill, bold) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA }, borders: cellBorders,
    shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: content, bold })] })],
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, 'F7F2E8', true)) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(c, i, ri % 2 ? SAND : undefined, i === 0)) })),
    ],
  });
}

const spacer = () => new Paragraph({ spacing: { after: 120 }, children: [] });

const doc = new Document({
  creator: 'KGC KRONOS',
  title: 'KRONOS Offline: what it is and what it does',
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { font: FONT, size: 44, bold: true, color: NAVY },
        paragraph: { spacing: { before: 0, after: 80 }, keepNext: true } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: NAVY },
        paragraph: { spacing: { before: 320, after: 120 }, keepNext: true, outlineLevel: 0 } },
    ],
  },
  numbering: { config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    ...['steps', 'checks', 'try'].map((ref) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] })),
  ] },
  sections: [{
    properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ text: 'KRONOS Offline, version 5   ·   page ', size: 18, color: '595959' }),
      new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '595959' }),
    ] })] }) },
    children: [
      new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun('KRONOS Offline')] }),
      p([new TextRun({ text: 'What it is and what it does, in plain words', size: 26, color: '595959' })], { spacing: { after: 240 } }),

      box([
        p([new TextRun({ text: 'In short', bold: true, color: NAVY, size: 24 })]),
        bullet('KRONOS helps you make Excel macros: small programs that do a repetitive Excel job for you.'),
        bullet('You describe the job. An AI chat writes a short "recipe" for it. KRONOS checks the recipe, tests it on made-up data and, only then, turns it into a macro.'),
        bullet('It runs in your web browser. Nothing to install, no internet needed. It never changes your original file and never opens Excel.'),
        bullet('It walks you through four steps: Ask, Questions, Review and Save. You can also open a file that already holds recipes.'),
        bullet('If you choose one of your workbooks, KRONOS can make the recipe\'s changes on a copy of it, which you download.'),
      ]),

      h1('1. KRONOS in one paragraph'),
      p('Many office jobs in Excel are the same few steps, done again and again: remove the cancelled rows, sort by date, colour the rejected items, count what is still open. A macro can do these jobs in a second. Writing a macro, though, normally needs a programmer. KRONOS closes that gap. You (or an AI chat helping you) describe the job as a simple recipe. KRONOS reads the recipe, proves it does what it says, and writes the macro for you.'),
      p([t('The important idea: '), b('the AI only writes the recipe, and KRONOS writes the macro.'), t(' KRONOS will not hand over a macro until its recipe has passed its tests. That is what makes it trustworthy.')]),

      h1('2. Why it is safe'),
      bullet([b('No internet. '), t('Everything KRONOS needs is inside its folder. It works with the network cable unplugged.')]),
      bullet([b('Nothing installed. '), t('You unzip a folder and double-click. Nothing is added to your PC.')]),
      bullet([b('It never opens Excel or runs a macro. '), t('It only prepares the macro. Using it in Excel is always your own choice, later.')]),
      bullet([b('It never changes your original file. '), t('If you choose a workbook, KRONOS reads it inside your browser, makes the changes on a copy and gives you the copy to download. Your original is untouched, and the workbook never leaves your PC.')]),
      bullet([b('Practice tests use invented data. '), t('A recipe is tested on made-up example sheets first, never on your workbook. Your workbook is only used when you ask for the changes to be made on a copy.')]),

      h1('3. How it works: Check, Test, Build'),
      p('Every recipe goes through the same three steps, in order:'),
      table(['Step', 'What KRONOS does', 'In everyday words'], [
        ['1. Check', 'Reads the recipe and makes sure it follows the rules and makes sense.', 'Like proof-reading a form before it is sent.'],
        ['2. Test', 'Runs the recipe on an invented practice sheet and checks the answers.', 'Like a trial run on a copy, never the real thing.'],
        ['3. Build', 'Writes the Excel macro (a .bas file that can be imported into Excel).', 'The finished tool, ready when you are.'],
      ], [1600, 4026, 3400]),
      spacer(),
      p([t('At the end you see a clear answer: '), new TextRun({ text: 'PROVEN', bold: true, color: GREEN }),
        t(' when every test passed, or '), new TextRun({ text: 'NOT PROVEN', bold: true, color: AMBER }),
        t(', with the problems listed in plain words.')]),

      h1('4. The four steps'),
      p('KRONOS walks you through the same four steps every time. The step you are on is underlined at the top of the page.'),
      table(['Step', 'What you do', 'What KRONOS does'], [
        ['1. Ask', 'Write what the macro should do, in plain words. Click "Write the message for the AI", copy it, and paste it into your AI chat. Then paste the AI\'s whole reply back.', 'Writes the message for the AI (with your workbook\'s description, if you chose one), then checks the reply.'],
        ['2. Questions', 'Only if the AI asks something: write your answer, copy it to the AI, and paste its new reply.', 'Keeps the whole conversation, so the AI knows what was said.'],
        ['3. Review', 'Read what the recipe does, what it assumes and what it changes, and tick "These assumptions are right".', 'Shows PROVEN or NOT PROVEN. If something is wrong, gives you the problems to send back to the AI (up to 3 tries).'],
        ['4. Save', 'Type your real sheet names and download the macro (.bas), or make the changes on a copy of your workbook.', 'Builds the macro with import steps, a practice run, and the changed copy.'],
      ], [1500, 3900, 3626]),
      spacer(),
      p([b('Two ways to copy for the AI. '), t('"for the Orchestrator" is for your claude.ai Orchestrator project, which already has KRONOS\'s instructions. "with instructions" is for ChatGPT or any new chat: it includes everything the AI needs.')]),

      h1('5. Recipes you already have'),
      p('If a recipe came to you in a file, there is no need to copy and paste it. Click "Open a recipe file…" and choose the file: a .vcl, .txt or .md file, or a Word document (.docx).'),
      bullet([b('One recipe in the file: '), t('KRONOS checks it straight away and goes to Review.')]),
      bullet([b('Several recipes: '), t('KRONOS lists them by what they do. Tick the ones you want, or click "Choose all", then "Check the chosen recipes".')]),
      bullet([b('A recipe with problems: '), t('"Send the problems to the AI" gives you a message for your AI chat with the recipe and what is wrong with it.')]),
      p('When you make the changes on a copy of your workbook with several recipes, KRONOS runs them one after another, in order, on one copy, and shows what each one changed.', { spacing: { before: 120, after: 120 } }),

      h1('6. Using your own workbook'),
      p('You can give KRONOS one of your Excel workbooks (a .xlsx or .xlsm file, up to 25 MB). KRONOS then helps in two ways:'),
      bullet([b('It describes the workbook for the AI. '), t('KRONOS lists your sheets, their column headings and the first few rows, and adds this to the message for the AI, so the AI uses your real names. You can untick "Include my workbook\'s description" to leave it out.')]),
      bullet([b('It makes the changes on a copy. '), t('In Save, KRONOS runs the proven recipes on a copy of your workbook and shows what changed: cells, rows deleted, rows moved. You then download the changed copy.')]),
      p('Fonts, borders, column widths, notes and links stay with their cells. Some things a copy cannot keep, such as pictures, or formulas and merged cells after rows have moved. KRONOS lists exactly what it found, and asks you to tick "I understand" before you can download. Always open the copy in Excel and look it over before you use it.', { spacing: { before: 120, after: 120 } }),

      h1('7. Getting started'),
      box([
        step('Download the KRONOS Offline zip file.'),
        step('Right-click the zip, choose "Extract All...", then "Extract".'),
        step('Open the KRONOS-Offline folder and double-click "Start KRONOS (Windows)".'),
        step('A small black window opens, and KRONOS opens in your browser. If the browser does not open by itself, open Chrome, Edge or Firefox and type localhost:5181 in the address bar.'),
        step('To stop KRONOS, close the black window.'),
        p([new TextRun({ text: 'Please use Chrome, Edge or Firefox. KRONOS does not work in Internet Explorer.', italics: true })], { spacing: { before: 120, after: 0 } }),
      ], SAND),

      h1('8. Try the practice example'),
      p('KRONOS comes with one ready-made example, so you can see it work straight away. The example tidies an order list: it deletes every order marked "Cancelled", then sorts the rest by order date, newest first.'),
      step('In step 1, click "Try the practice example".', 'try'),
      step([t('Wait a moment. KRONOS goes to Review and shows '), new TextRun({ text: 'PROVEN', bold: true, color: GREEN }), t('.')], 'try'),
      step('Tick "These assumptions are right", click "Go to Save" and download the macro if you like.', 'try'),
      p('Behind the scenes, KRONOS tested the recipe on an invented sheet of 40 orders. It checked that no cancelled orders are left, that the right number of rows was removed, and that the dates really are newest first. Seeing PROVEN means KRONOS works on your PC.', { spacing: { before: 120, after: 120 } }),

      h1('9. What it can do today, and what is coming next'),
      p('This is version 5, dated 16 November 2026. It has the four steps, the copy-and-paste link to your AI chat, recipe files and changes on a copy of your workbook.'),
      table(['Today (version 5)', 'Coming in later versions'], [
        ['The four steps: Ask, Questions, Review, Save', 'A history of past runs, kept in your browser'],
        ['Messages for your AI chat, and its replies checked', 'A Help page'],
        ['Recipe files (.vcl, .txt, .md, .docx), one or many at a time', 'Checking a practice run from Excel inside the page'],
        ['Download the macro (.bas) and its practice run', 'Suggesting a good recipe for KRONOS\'s own collection'],
        ['Make proven recipes\' changes on a copy of your workbook', ''],
      ], [4513, 4513]),
      spacer(),
      p([b('Please tell us how it goes: '), t('KRONOS has been tested in several browsers, but your own work PC is the real test. Anything it shows will be fixed in the next version. Large sheets take longer: deleting many rows from a sheet of 5,000 rows takes about a minute, and Cancel stops it.')]),

      h1('10. Quick questions and answers'),
      ...[
        ['Does it need the internet?', 'No. Everything is inside the folder.'],
        ['Do I need to install anything?', 'No. Unzip and double-click.'],
        ['Will it change my Excel files?', 'No. It never opens Excel and never changes your original. If you choose a workbook, it makes a changed copy for you to download.'],
        ['Does my workbook leave my PC?', 'No. KRONOS reads it inside your browser, and nothing is sent anywhere.'],
        ['Is the changed copy always perfect?', 'Look it over in Excel before you use it. KRONOS tells you what it could not keep, and you must tick "I understand" before you can download.'],
        ['Which browser should I use?', 'Chrome, Edge or Firefox. Not Internet Explorer.'],
        ['What is the black window?', 'A tiny helper that shows KRONOS to your browser. It only talks to your own PC. Close it to stop KRONOS.'],
        ['Does KRONOS use AI?', 'Not by itself. You copy KRONOS\'s message into an AI chat and paste the reply back. KRONOS then does all the checking, testing and building on your PC.'],
        ['What if I see NOT PROVEN?', 'Nothing is harmed. The recipe has a problem, and KRONOS lists it. Copy the problems to the AI chat and paste its new reply; KRONOS gives it up to 3 tries.'],
        ['What happens if I click Cancel?', 'KRONOS stops what it is doing and keeps everything you pasted, so you can carry on.'],
      ].flatMap(([q, a]) => [
        p([b(q)], { spacing: { before: 120, after: 40 }, keepNext: true }),
        p(a),
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2], buf);
  console.log('wrote', process.argv[2]);
});
