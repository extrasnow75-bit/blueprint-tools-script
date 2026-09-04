/**
 * ================================================================
 * BLUEPRINT TOOLS  |  'Add Activity Titles, Tools, Due Date Headers, & Times'
 * ================================================================
 * Last updated on 2026-09-04 at 09:44 MDT
 * ================================================================
 */
const RED       = '#ff0000';
const DEEP_BLUE = '#0033a0';
const BLACK     = '#000000';
const GREY_CHIP = '#e8eaed'; // the OLD Canvas-tool highlight. No longer applied —
                             // kept only so _clearGreyChip can recognise and
                             // remove it from Blueprints built before the marker.
const FONT      = 'Arial';
// A due-by marker line, in either of the two shapes a Blueprint can hold: the
// template's placeholder, whose day sits in a dropdown chip and so reads as
// empty ("Due by  at 11:59 p.m. Mountain Time"), and a header this tool inserted,
// which carries a real day. Keying on the fixed wording catches both.
const DUE_HEADER_RE = /due by\b.*mountain time/i;
// The pattern that authorises DELETION, and it is deliberately much stricter
// than the one above. DUE_HEADER_RE only decides what gets a blue border, so
// being permissive there costs nothing. Here a false positive destroys the
// user's writing, so the line must START with "Due by" and carry no more than a
// short tail — enough for "Unpublished text header in Canvas" or a trailing
// "Display header as:" chip, nowhere near enough for a sentence.
//
// The case this exists to refuse: a Directions paragraph reading "Submit your
// draft to the dropbox. This is due by Sunday at 11:59 p.m. Mountain Time."
// It matches DUE_HEADER_RE and must never be deleted.
const DUE_MARKER_RE = /^\s*due by\b[^.]{0,20}?\bat 11:59 p\.m\. mountain time\b.{0,60}$/i;
// The placeholder's companion chip, for the case where it is its own paragraph
// rather than a right-tabbed run on the due-by line.
const DISPLAY_AS_RE = /display header as/i;
// The Canvas-tool line reads "<Tool> ⏺ Link to settings tab".
//
// The circle replaced a grey background highlight on the tool name. QA marks
// text light green once it is ready to go into Canvas, and a background colour
// is a single attribute per character — the green overwrote the grey outright,
// so the one cue that identified a Canvas tool vanished at exactly the moment
// the line mattered most. A real character cannot be overwritten by a highlight.
const TOOL_MARKER = '⏺';   // ⏺ BLACK CIRCLE FOR RECORD
// The marker is black, not red: it is a structural mark that says "this line
// tags a Canvas tool", not part of the tool's name. Keeping it out of the red
// also stops it reading as punctuation belonging to the words on either side.
const TOOL_MARKER_COLOR = BLACK;
const TOOL_SUFFIX = ' ' + TOOL_MARKER + ' Link to settings tab';
// Blueprints built before the marker landed still read "<Tool>; Link to settings
// tab". Both forms must be recognised: on a re-run over an old document here,
// and by getToolTypeForSlot in Code2.gs, which reads this line to decide which
// directions to deploy.
const TOOL_SUFFIX_LEGACY = '; Link to settings tab';
// The complete set of Canvas tools a slot may be tagged with. These strings are
// load-bearing downstream — they are the DIRECTION_OPTIONS keys in Code2.gs and
// the normaliseToolKey values in Sidebar5 — so they must match exactly.
const CANVAS_TOOL_OPTIONS = [
  'Assignment',
  'Assignment (Not Graded)',
  'Discussion',
  'Page',
  'Quiz (Classic)',
  'Quiz (New)'
];
// Sentinel the pre-flight panel sends when a tool cell should stay untagged.
// The slot is not left blank — it keeps the TOOL_PLACEHOLDER line, so the
// designer still sees a prompt to pick a tool by hand.
const TOOL_LEAVE_UNCHANGED = '[Leave unchanged]';
// The un-set tool name. setNearbyTool swaps this for a real tool when one is
// known; until then it stays put, marked like any tagged tool.
const TOOL_PLACEHOLDER = 'Select Tool';
// Upper bound on modules, enforced server-side in processBlueprint. Matches the
// max="52" on the sidebar's numModules input (one module per week, a full year).
const MAX_MODULES = 52;
// ── MENU ──────────────────────────────────────────────────────────
function onOpen() {
  DocumentApp.getUi()
    .createMenu('🎓 Blueprint Tools')
    .addItem('Add Activity Titles, Tools, Due Date Headers, & Times', 'showSidebar')
    // Marked Beta until it has been exercised against real Blueprints — it is
    // the only tool that writes to module headings and the only one that
    // depends on an external site's markup.
    .addItem('Add Module Titles & Module Dates (Beta)', 'showModuleTitlesSidebar7')
    .addSeparator()
    .addItem('Create Model Module',                    'showModelModuleNoAiSidebar')
    .addSeparator()
    .addItem('Deploy Activity Directions',             'showDirectionsSidebar')
    .addSeparator()
    .addItem('Time Estimator',                         'showTimeEstimatorSidebar')
    .addSeparator()
    .addItem('KB Article',                             'showKbArticle')
    .addToUi();
}
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Add Activity Titles, Tools, Due Date Headers, & Times')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}
// ── KB ARTICLE ────────────────────────────────────────────────────
// A menu can't open a URL directly, so show a small dialog that opens the
// Knowledge Base article in a new tab (with a click-through link as a fallback
// in case the browser blocks the automatic pop-up).
const KB_ARTICLE_URL =
  'https://docs.google.com/document/d/1fSKIamcfOZFthkEtScX_OGEDjx6g8uMLX5j7N5wmUTQ/edit?tab=t.dt1zwyxe9p6f#heading=h.wac296jbizmu';
function showKbArticle() {
  const url  = KB_ARTICLE_URL;
  const html = HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;font-size:13px;color:#333;padding:4px 2px;">' +
      '<p style="margin:0 0 12px;">Opening the Blueprint Tools Knowledge Base article in a new tab&hellip;</p>' +
      '<p style="margin:0;">If it doesn’t open, ' +
      '<a href="' + url + '" target="_blank" rel="noopener" style="color:#0033a0;font-weight:bold;">' +
      'click here to open the KB Article &rarr;</a></p>' +
      '<script>try{window.open(' + JSON.stringify(url) + ',"_blank","noopener");}catch(e){}</script>' +
      '</div>')
    .setWidth(340)
    .setHeight(120);
  DocumentApp.getUi().showModalDialog(html, 'KB Article');
}
// ── MAIN ──────────────────────────────────────────────────────────
function processBlueprint(params) {
  // The sidebar's max="52" and parseInt are client-side only — google.script.run
  // is callable directly, so the bound has to hold here too. An unbounded value
  // drives the createModule/deleteModule loops below past the 6-minute limit and
  // leaves a half-built Development tab behind.
  const requested = Number(params.numModules);
  if (!Number.isInteger(requested) || requested < 1 || requested > MAX_MODULES)
    throw new Error(`Number of modules must be a whole number from 1 to ${MAX_MODULES}.`);
  const doc  = DocumentApp.getActiveDocument();
  const tabs = collectTabs(doc);
  const designTab = tabs.find(t => /\bdesign\b/i.test(t.title));
  const devTab    = tabs.find(t => /\bdevelopment\b/i.test(t.title));
  if (!designTab || !devTab)
    throw new Error('Could not find "Design" and/or "Development" tabs.');
  const activities = parseCoursePattern(designTab.body, params.toolOverrides);
  if (activities.length === 0)
    throw new Error('No activities found in the course pattern table.');
  const stats      = { created: 0, deleted: 0, filled: 0, tools: 0, slotsDeleted: 0,
                       headers: 0, headersRemoved: 0 };
  const numModules = requested;
  const existing   = countExistingModules(devTab.body);
  const indent     = getTemplateIndent(devTab.body);
  for (let m = existing; m > numModules; m--) {
    deleteModule(devTab.body, m);
    stats.deleted++;
  }
  const afterMod  = Math.min(existing, numModules);
  let   insertIdx = findModuleInsertionPoint(devTab.body, afterMod);
  for (let m = afterMod + 1; m <= numModules; m++) {
    insertIdx = createModule(devTab.body, m, params, indent, insertIdx, activities);
    stats.created++;
  }
  for (let m = 1; m <= numModules; m++) {
    processModule(devTab.body, m, activities, params, indent, stats);
  }
  // Add blue top/bottom borders to every due-by header in the Development tab.
  // Paragraph borders are single-valued style attributes, so re-applying the
  // same spec to a header that already has it is a harmless no-op — which lets
  // us border all headers uniformly without tracking which are new.
  // Must be last: it persists the document (saveAndClose) so the Docs API can
  // read the freshly inserted headers.
  applyDueHeaderBorders(doc, devTab.title);
  return buildSummary(stats, params, activities, numModules);
}
// ── COLLECT TABS ──────────────────────────────────────────────────
function collectTabs(doc) {
  const result = [];
  function walk(tab) {
    result.push({ title: tab.getTitle(), body: tab.asDocumentTab().getBody() });
    tab.getChildTabs().forEach(walk);
  }
  doc.getTabs().forEach(walk);
  return result;
}
// ── PARSE COURSE PATTERN TABLE ────────────────────────────────────
/**
 * Reads the Course Pattern Table into activity records.
 *
 * @param {GoogleAppsScript.Document.Body} body  the Design tab body
 * @param {Object} [toolOverrides]  user picks from the sidebar's pre-flight
 *   panel, keyed by LOWERCASED raw cell text → canonical tool name (or
 *   TOOL_LEAVE_UNCHANGED). Keying by cell text rather than activity name means one
 *   pick resolves every row that shares the same spelling.
 * @returns {Array<{name, tool, rawTool, dueDay, time}>}  `tool` is null when the
 *   cell was blank OR unrecognized and unresolved; `rawTool` keeps the original
 *   cell text so callers can tell those two cases apart.
 */
function parseCoursePattern(body, toolOverrides) {
  const overrides  = toolOverrides || {};
  const activities = [];
  for (const table of body.getTables()) {
    if (table.getNumRows() < 2) continue;
    const headerRow = table.getRow(0);
    const numCells  = headerRow.getNumCells();
    if (numCells < 2) continue;
    let actCol = -1, toolCol = -1, dayCol = -1, timeCol = -1;
    for (let c = 0; c < numCells; c++) {
      const h = headerRow.getCell(c).getText().toLowerCase();
      if      (actCol  < 0 && (h.includes('activity') || h.includes('assessment'))) actCol  = c;
      else if (toolCol < 0 && (h.includes('tool') || h.includes('canvas')))         toolCol = c;
      else if (dayCol  < 0 && (h.includes('due') || h.includes('day')))             dayCol  = c;
      else if (timeCol < 0 && (h.includes('time') || h.includes('estimate')))       timeCol = c;
    }
    if (actCol < 0) continue;
    for (let r = 1; r < table.getNumRows(); r++) {
      const row  = table.getRow(r);
      const n    = row.getNumCells();
      const name = actCol < n ? row.getCell(actCol).getText().trim() : '';
      if (!name) continue;
      const rawTool = toolCol < n ? row.getCell(toolCol).getText().trim() : '';
      let   tool    = normalizeToolName(rawTool);
      // A non-empty cell we couldn't interpret may have been resolved by the
      // user in the pre-flight panel. Validate the pick against the canonical
      // list so a stale or hand-edited value can't inject an unknown tool name.
      if (!tool && rawTool) {
        const picked = overrides[rawTool.toLowerCase()];
        if (picked && CANVAS_TOOL_OPTIONS.indexOf(picked) !== -1) tool = picked;
      }
      activities.push({
        name,
        tool,
        rawTool,
        dueDay: parseDueDay(dayCol  < n ? row.getCell(dayCol).getText().trim()  : ''),
        time:   timeCol < n ? row.getCell(timeCol).getText().trim() : ''
      });
    }
  }
  return activities;
}
// ── PRE-FLIGHT: UNRECOGNIZED TOOL CELLS ───────────────────────────
/**
 * Sidebar-callable. Scans the Course Pattern Table for Canvas Tool cells that
 * normalizeToolName cannot interpret, so the user can map them to a real tool
 * BEFORE the run rather than discovering a missing tool line afterwards.
 *
 * An unrecognized cell is otherwise silent: the activity simply gets no tool
 * line, and the summary still reports success. That is how "Assignment
 * (Ungraded)" went unnoticed.
 *
 * Blank cells are NOT reported — a row with no Canvas tool is legitimate, and
 * flagging it would nag on every run.
 *
 * Results are grouped by cell text, so a spelling used by six activities is one
 * dropdown, not six.
 *
 * @returns {{unrecognized: Array<{rawTool: string, activities: string[]}>,
 *            toolOptions: string[], leaveUnchanged: string, error?: string}}
 */
function scanCoursePatternTools() {
  const base = { unrecognized: [], toolOptions: CANVAS_TOOL_OPTIONS, leaveUnchanged: TOOL_LEAVE_UNCHANGED };
  const doc  = DocumentApp.getActiveDocument();
  const designTab = collectTabs(doc).find(t => /\bdesign\b/i.test(t.title));
  if (!designTab) {
    // Non-fatal: the sidebar still opens and Run reports the missing tab.
    base.error = 'Could not find a "Design" tab in this document.';
    return base;
  }

  const byCell = {};
  const order  = [];
  for (const act of parseCoursePattern(designTab.body)) {
    if (act.tool || !act.rawTool) continue;   // recognized, or legitimately blank
    const key = act.rawTool.toLowerCase();
    if (!byCell[key]) {
      byCell[key] = { rawTool: act.rawTool, activities: [] };
      order.push(key);
    }
    byCell[key].activities.push(act.name);
  }

  base.unrecognized = order.map(k => byCell[k]);
  Logger.log('scanCoursePatternTools: %s unrecognized tool cell(s).', base.unrecognized.length);
  return base;
}
// ── NORMALIZE TOOL NAME ───────────────────────────────────────────
function normalizeToolName(raw) {
  const t = raw.toLowerCase().trim();
  if (!t) return null;
  // Designers write this cell several ways — "Assignment (Ungraded)",
  // "Assignment (Not Graded)", "non-graded". All normalize to the official
  // Canvas marker "Assignment (Not Graded)", which is a DISTINCT tool type
  // downstream: it has its own direction menu in Code2.gs (DIRECTION_OPTIONS)
  // and its own key in Sidebar5's normaliseToolKey. Falling through to plain
  // 'Assignment' silently offers the graded-assignment directions instead.
  //
  // "Ungraded" is only ever a qualifier on an ASSIGNMENT. A cell that names a
  // different tool — "Discussion (Ungraded)", "Page (Ungraded)" — keeps that
  // tool, so the check is gated on no other tool being named. Without the gate
  // it swallows every ungraded cell and rewrites discussions as assignments.
  // Anchored on \b so the Feedback column's "Auto graded" can never match.
  const ungraded = /\b(un|non[- ]?|not )graded\b/.test(t);
  if (ungraded && !/discussion|page|quiz/.test(t))       return 'Assignment (Not Graded)';
  if (t.includes('assignment'))                          return ungraded ? 'Assignment (Not Graded)' : 'Assignment';
  if (t.includes('discussion'))                          return 'Discussion';
  if (t.includes('page'))                                return 'Page';
  if (t.includes('quiz') && t.includes('new'))           return 'Quiz (New)';
  if (t.includes('quiz'))                                return 'Quiz (Classic)';
  return null;
}
// ── PARSE DUE DAY ─────────────────────────────────────────────────
const DAY_ALIASES = [
  [/\bthursday\b/i,  'Thursday'],  [/\btuesday\b/i,   'Tuesday'],
  [/\bsaturday\b/i,  'Saturday'],  [/\bwednesday\b/i, 'Wednesday'],
  [/\bmonday\b/i,    'Monday'],    [/\bsunday\b/i,    'Sunday'],
  [/\bfriday\b/i,    'Friday'],    [/\bthurs\b/i,     'Thursday'],
  [/\btues\b/i,      'Tuesday'],   [/\bthur\b/i,      'Thursday'],
  [/\bwed\b/i,       'Wednesday'], [/\bmon\b/i,       'Monday'],
  [/\bfri\b/i,       'Friday'],    [/\bsat\b/i,       'Saturday'],
  [/\bsun\b/i,       'Sunday'],    [/\bthu\b/i,       'Thursday'],
  [/\btue\b/i,       'Tuesday']
];
function parseDueDay(text) {
  if (!text) return null;
  for (const [re, full] of DAY_ALIASES) {
    if (re.test(text)) return full;
  }
  return null;
}
// ── DUE DAY GROUPS ────────────────────────────────────────────────
function getDueDayGroups(activities) {
  const groups = [], seen = new Set();
  for (let i = 0; i < activities.length; i++) {
    const day = activities[i].dueDay;
    if (!day || seen.has(day)) continue;
    seen.add(day);
    groups.push({ day, startIndex: i });
  }
  return groups;
}
// ── COUNT EXISTING MODULES ────────────────────────────────────────
function countExistingModules(body) {
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  let count = 0;
  for (const p of body.getParagraphs()) {
    if (p.getHeading() === H2 && /^Module\s+\d+[:\s]/i.test(p.getText().trim())) count++;
  }
  return count;
}
// ── FIND MODULE INSERTION POINT ───────────────────────────────────
function findModuleInsertionPoint(body, afterModNum) {
  if (afterModNum <= 0) return -1;
  const paras = body.getParagraphs();
  const H2    = DocumentApp.ParagraphHeading.HEADING2;
  const modRe = new RegExp(`^Module\\s+${afterModNum}[:\\s]`, 'i');
  let inMod   = false;
  for (let i = 0; i < paras.length; i++) {
    if (paras[i].getHeading() !== H2) continue;
    const t = paras[i].getText().trim();
    if (modRe.test(t)) { inMod = true; continue; }
    if (inMod) return i;
  }
  return -1;
}
// ── GET TEMPLATE INDENT ───────────────────────────────────────────
function getTemplateIndent(body) {
  const slots = getSlotsInModule(body, 1);
  if (slots.length === 0) return 36;
  return slots[0].para.getIndentStart() || 36;
}
// ── CREATE MODULE ─────────────────────────────────────────────────
function createModule(body, modNum, params, indent, insertIdx, activities) {
  const H2     = DocumentApp.ParagraphHeading.HEADING2;
  const H3     = DocumentApp.ParagraphHeading.HEADING3;
  const H4     = DocumentApp.ParagraphHeading.HEADING4;
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  let   idx    = insertIdx;
  function add(text) {
    return (idx < 0) ? body.appendParagraph(text) : body.insertParagraph(idx++, text);
  }
  // Module heading — H2, bold
  const hPara = add(`Module ${modNum}: Title (start date - end date)`);
  hPara.setHeading(H2);
  _fmt(hPara.editAsText(), { font: FONT, bold: true });
  // Module overview — H3, NOT bold, Arial 15pt, black
  const ovPara = add(`Module ${modNum} Overview`);
  ovPara.setHeading(H3);
  _fmt(ovPara.editAsText(), { font: FONT, size: 15, bold: false, color: BLACK });
  // Refer-to note — Normal text, Arial 11pt, black, not italic
  const refPara = add('[Refer to the Template Blueprint Customization by Program document to populate this section.]');
  refPara.setHeading(NORMAL);
  _fmt(refPara.editAsText(), { font: FONT, size: 11, bold: false, italic: false, color: BLACK });
  // Activity slots — one per activity in the course pattern table
  for (let slot = 1; slot <= activities.length; slot++) {
    const prefix = params.numbered ? `${modNum}.${String(slot).padStart(2,'0')} ` : '';
    // Activity title — H4, Arial 15pt, black, NOT bold, NOT italic
    const aPara = add(`${prefix}Activity Title`);
    aPara.setHeading(H4);
    aPara.setIndentStart(indent);
    _fmt(aPara.editAsText(), { font: FONT, size: 15, bold: false, italic: false, color: BLACK });
    // Estimated time
    const ePara = add('Estimated time:');
    ePara.setHeading(NORMAL);
    ePara.setIndentStart(indent);
    _fmt(ePara.editAsText(), { font: FONT, size: 11, italic: true });
    // Select Tool ⏺ Link to settings tab
    const tPara = add(TOOL_PLACEHOLDER + TOOL_SUFFIX);
    tPara.setHeading(NORMAL);
    tPara.setIndentStart(indent);
    _fmtToolLine(tPara);
    // Directions
    const dPara = add('Directions go here\u2026');
    dPara.setHeading(NORMAL);
    dPara.setIndentStart(indent);
    _fmt(dPara.editAsText(), { font: FONT, size: 11 });
  }
  // Spacer between modules
  const spacer = add('');
  spacer.setHeading(NORMAL);
  return idx;
}
// ── DELETE MODULE ─────────────────────────────────────────────────
function deleteModule(body, modNum) {
  const paras = body.getParagraphs();
  const H2    = DocumentApp.ParagraphHeading.HEADING2;
  const modRe = new RegExp(`^Module\\s+${modNum}[:\\s]`, 'i');
  let start = -1, end = paras.length;
  for (let i = 0; i < paras.length; i++) {
    if (paras[i].getHeading() !== H2) continue;
    const t = paras[i].getText().trim();
    if (modRe.test(t) && start < 0) { start = i; continue; }
    if (start >= 0) { end = i; break; }
  }
  if (start < 0) return;
  for (let i = end - 1; i >= start; i--) {
    try { paras[i].removeFromParent(); } catch(e) {}
  }
}
// ── PROCESS ONE MODULE ────────────────────────────────────────────
function processModule(body, modNum, activities, params, indent, stats) {
  const slots     = getSlotsInModule(body, modNum);
  const slotParas = [];
  for (const { slotNum, para } of slots) {
    if (slotNum <= activities.length) {
      fillSlot(body, para, modNum, slotNum, activities[slotNum - 1], params, indent, stats);
      slotParas.push(para);
    } else {
      removeSlot(body, para);
      stats.slotsDeleted++;
    }
  }
  if (activities.length > slots.length) {
    let idx = getIndexAfterSlots(body, modNum, slots);
    for (let s = slots.length + 1; s <= activities.length; s++) {
      const result = insertActivitySlot(body, modNum, s, activities[s - 1], params, indent, idx, stats);
      idx = result.idx;
      slotParas.push(result.h4Para);
    }
  }
  stats.headers += placeDueHeaders(body, slotParas, activities, params, stats);
}
// ── GET INDEX AFTER LAST SLOT ─────────────────────────────────────
function getIndexAfterSlots(body, modNum, slots) {
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const H4 = DocumentApp.ParagraphHeading.HEADING4;
  if (slots.length > 0) {
    const lastH4 = slots[slots.length - 1].para;
    const start  = body.getChildIndex(lastH4);
    let lastIdx  = start;
    // Hoisted: this loop only walks forward and never inserts, so re-reading the
    // child count every iteration is a round trip that returns the same number.
    const n = body.getNumChildren();
    for (let i = start + 1; i < n; i++) {
      const child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) break;
      const para = child.asParagraph();
      const h    = para.getHeading();
      if (h === H2 || h === H3 || h === H4) break;
      if (!para.getText()) break; // stop before empty spacer between modules
      lastIdx = i;
    }
    return lastIdx + 1;
  }
  // No existing slots: insert before next module's H2 (or end of body)
  const modRe = new RegExp(`^Module\\s+${modNum}[:\\s]`, 'i');
  const n = body.getNumChildren();
  let inModule = false, lastIdx = 0;
  for (let i = 0; i < n; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      if (inModule) lastIdx = i;
      continue;
    }
    const para = child.asParagraph();
    if (para.getHeading() === H2) {
      if (modRe.test(para.getText().trim())) { inModule = true; lastIdx = i; continue; }
      if (inModule) return i;
    }
    if (inModule) lastIdx = i;
  }
  return lastIdx + 1;
}
// ── INSERT A NEW ACTIVITY SLOT AT INDEX ───────────────────────────
function insertActivitySlot(body, modNum, slotNum, activity, params, indent, insertIdx, stats) {
  const H4     = DocumentApp.ParagraphHeading.HEADING4;
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  let idx = insertIdx;
  function ins(text) { return body.insertParagraph(idx++, text); }
  let title = activity.name;
  if (params.timeEstimates && activity.time) title += ` (${activity.time})`;
  const prefix = params.numbered ? `${modNum}.${String(slotNum).padStart(2,'0')} ` : '';
  const aPara  = ins(prefix + title);
  aPara.setHeading(H4);
  aPara.setIndentStart(indent);
  _fmt(aPara.editAsText(), { font: FONT, size: 15, bold: false, italic: false, color: BLACK });
  stats.filled++;
  const estText = (!params.timeEstimates && activity.time) ? 'Estimated time: ' + activity.time : 'Estimated time:';
  const ePara = ins(estText);
  ePara.setHeading(NORMAL);
  ePara.setIndentStart(indent);
  _fmt(ePara.editAsText(), { font: FONT, size: 11, italic: true });
  const tPara = ins(TOOL_PLACEHOLDER + TOOL_SUFFIX);
  tPara.setHeading(NORMAL);
  tPara.setIndentStart(indent);
  _fmtToolLine(tPara);
  const dPara = ins('Directions go here…');
  dPara.setHeading(NORMAL);
  dPara.setIndentStart(indent);
  _fmt(dPara.editAsText(), { font: FONT, size: 11 });
  if (activity.tool && setNearbyTool(body, aPara, activity.tool)) stats.tools++;
  return { idx, h4Para: aPara };
}
// ── GET SLOTS IN MODULE ───────────────────────────────────────────
function getSlotsInModule(body, modNum) {
  const H2    = DocumentApp.ParagraphHeading.HEADING2;
  const H4    = DocumentApp.ParagraphHeading.HEADING4;
  const modRe = new RegExp(`^Module\\s+${modNum}[:\\s]`, 'i');
  let inModule = false, counter = 0;
  const slots  = [];
  for (const para of body.getParagraphs()) {
    const h = para.getHeading();
    // Only H2 rows need their text (the module-boundary test); deferring
    // getText() past the heading check drops a round-trip on every other row.
    if (h === H2) {
      const t = para.getText().trim();
      if (modRe.test(t)) { inModule = true; continue; }
      if (inModule) break;
    }
    if (!inModule || h !== H4) continue;
    slots.push({ slotNum: ++counter, para });
  }
  return slots;
}
// ── FILL ONE SLOT ─────────────────────────────────────────────────
function fillSlot(body, headingPara, modNum, slotNum, activity, params, indent, stats) {
  let title = activity.name;
  if (params.timeEstimates && activity.time) title += ` (${activity.time})`;
  const prefix   = params.numbered ? `${modNum}.${String(slotNum).padStart(2,'0')} ` : '';
  const fullText = prefix + title;
  headingPara.setText(fullText);
  headingPara.setIndentStart(indent);
  _fmt(headingPara.editAsText(), { font: FONT, size: 15, bold: false, italic: false, color: BLACK });
  stats.filled++;
  if (activity.tool && setNearbyTool(body, headingPara, activity.tool)) stats.tools++;
  setNearbyEstimate(body, headingPara, activity.time, params.timeEstimates);
}
// ── SET TOOL WITH FORMATTING ──────────────────────────────────────
function setNearbyTool(body, headingPara, toolValue) {
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const H4 = DocumentApp.ParagraphHeading.HEADING4;
  const startIdx   = body.getChildIndex(headingPara);
  if (startIdx < 0) return false;
  const numChildren = body.getNumChildren();
  for (let i = startIdx + 1; i < Math.min(startIdx + 8, numChildren); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = child.asParagraph();
    const h    = para.getHeading();
    if (h === H2 || h === H3 || h === H4) break;
    const text          = para.getText();
    const hasSelectTool = text.includes(TOOL_PLACEHOLDER);
    // Match on the wording, not the separator. A Blueprint built before the
    // marker landed reads "<Tool>; Link to settings tab"; one this tool has
    // already touched reads "<Tool> ⏺ Link to settings tab". Keying on the
    // semicolon alone would silently stop finding the line on a second run.
    const hasSuffix     = text.includes('Link to settings tab');
    if (!hasSelectTool && !hasSuffix) continue;
    try {
      if (hasSelectTool) {
        para.replaceText(TOOL_PLACEHOLDER, toolValue);
        // Bring a legacy line onto the marker, so an old Blueprint ends up
        // looking exactly like a newly built one.
        para.replaceText(TOOL_SUFFIX_LEGACY, TOOL_SUFFIX);
        _fmtToolLine(para);
      } else if (hasSuffix) {
        // Only rewrite the text when it is actually wrong. setText() replaces
        // the whole run and resets every character attribute with it —
        // background colour included — so an unconditional rewrite would strip
        // QA's light-green "ready for Canvas" highlight off every tool line on
        // each re-run. That highlight is the very thing the marker change exists
        // to protect. _fmtToolLine still runs: it only sets foreground colour,
        // font and weight, so it repairs formatting without touching the green.
        if (text !== toolValue + TOOL_SUFFIX) para.setText(toolValue + TOOL_SUFFIX);
        _fmtToolLine(para);
      }
      Logger.log(`  Tool → ${toolValue}`);
      return true;
    } catch(e) {
      Logger.log(`  Tool error: ${e.message}`);
      return false;
    }
  }
  return false;
}
// ── SET ESTIMATED TIME LINE ───────────────────────────────────────
function setNearbyEstimate(body, headingPara, time, timeInTitle) {
  // If time goes in the title, or there's no time value, leave the line as-is
  if (timeInTitle || !time) return;
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const H4 = DocumentApp.ParagraphHeading.HEADING4;
  const startIdx    = body.getChildIndex(headingPara);
  if (startIdx < 0) return;
  const numChildren = body.getNumChildren();
  for (let i = startIdx + 1; i < Math.min(startIdx + 5, numChildren); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) break;
    const para = child.asParagraph();
    const h    = para.getHeading();
    if (h === H2 || h === H3 || h === H4) break;
    if (/^Estimated time/i.test(para.getText())) {
      para.setText('Estimated time: ' + time);
      _fmt(para.editAsText(), { font: FONT, size: 11, italic: true });
      return;
    }
  }
}
// ── REMOVE EXTRA SLOT ─────────────────────────────────────────────
// FIX: Use body.getChildIndex() instead of paras.indexOf()
function removeSlot(body, headingPara) {
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const H4 = DocumentApp.ParagraphHeading.HEADING4;
  const start       = body.getChildIndex(headingPara);
  if (start < 0) return;
  const numChildren = body.getNumChildren();
  const toRemove    = [headingPara];
  for (let i = start + 1; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) break;
    const para = child.asParagraph();
    const h    = para.getHeading();
    if (h === H4 || h === H3 || h === H2) break;
    toRemove.push(para);
  }
  for (let i = toRemove.length - 1; i >= 0; i--) {
    try { toRemove[i].removeFromParent(); } catch(e) {}
  }
}
// ── MODULE SPAN ───────────────────────────────────────────────────
// Child-index range [start, end) covering one module's content, derived from the
// module's own activity slots: walk back from the first slot to its H2 heading,
// forward from the last slot to the next one.
//
// Deriving the span from the slots beats rescanning the body for "Module N:".
// A full-body scan per module is O(modules x body), and at the 52-module ceiling
// that is real time against the 6-minute limit. This costs a handful of round
// trips instead. Callers must pass a non-empty slotParas.
//
// Returns null — meaning "do not clear anything" — when the module's own H2
// cannot be found above the first slot. FAILING CLOSED IS THE WHOLE POINT: an
// earlier version defaulted start to 0, so a module heading that had been hand
// restyled, or that lived inside a table (getSlotsInModule walks getParagraphs,
// which descends into cells; this walks top-level children, which does not),
// silently produced a span covering the ENTIRE tab. clearDueHeaders would then
// sweep every module in the document, chips and all, unrecoverably.
//
// The end bound stops at the last slot's own content block rather than running
// on to the next H2 or to the end of the body. Markers only ever sit ABOVE a
// slot, so nothing that needs clearing lives past the last one — and for the
// final module in the tab, running to the end would put trailing course-wrap
// content inside the deletion span for no benefit at all.
function getModuleSpan(body, slotParas) {
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const H4 = DocumentApp.ParagraphHeading.HEADING4;
  const n  = body.getNumChildren();
  const firstIdx = body.getChildIndex(slotParas[0]);
  const lastIdx  = body.getChildIndex(slotParas[slotParas.length - 1]);
  if (firstIdx < 0 || lastIdx < 0) {
    Logger.log('getModuleSpan: slot is not a top-level child; skipping clear.');
    return null;
  }
  let start = -1;
  for (let i = firstIdx - 1; i >= 0; i--) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    if (child.asParagraph().getHeading() === H2) { start = i + 1; break; }
  }
  if (start < 0) {
    Logger.log('getModuleSpan: no module heading above the first slot; skipping clear.');
    return null;
  }
  let end = n;
  for (let i = lastIdx + 1; i < n; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) { end = i; break; }
    const para = child.asParagraph();
    const h    = para.getHeading();
    if (h === H2 || h === H3 || h === H4) { end = i; break; }
    if (!para.getText()) { end = i; break; }   // the spacer between modules
  }
  return { start, end };
}
// ── COLLECT EXISTING DUE-BY MARKERS ───────────────────────────────
// Find every due-by marker in a module span — the template's chip placeholders
// and any header a previous run inserted alike — and return the ELEMENTS rather
// than their indices, so the caller can insert replacements before removing
// them without the indices going stale underneath it.
//
// Two shapes are handled: the due-by line itself, and a "Display header as:"
// paragraph immediately beside it, for the case where that chip is its own
// paragraph rather than a right-tabbed run on the same line.
//
// Tables are NOT touched. An earlier version removed any table containing the
// due-by wording, which would have destroyed an entire rubric or schedule table
// over a single matching row. That branch was speculative — written before we
// knew how the band is laid out — and speculation is not worth an unrecoverable
// deletion. If a band ever does turn out to be a table, the Logger line below
// reports it and we can handle that shape once we have a real example.
//
// A dropdown chip reads as empty string through Apps Script, so a paragraph made
// up ENTIRELY of chips cannot be told from a blank spacer and is deliberately
// left alone rather than guessed at.
function collectDueMarkers(body, span) {
  const marks = [];
  let tablesSeen = 0;
  for (let i = span.start; i < span.end; i++) {
    const child = body.getChild(i);
    const type  = child.getType();
    if (type === DocumentApp.ElementType.TABLE) {
      if (DUE_HEADER_RE.test(child.asTable().getText())) tablesSeen++;
      continue;
    }
    if (type !== DocumentApp.ElementType.PARAGRAPH) continue;
    if (DUE_MARKER_RE.test(child.asParagraph().getText())) marks.push(i);
  }
  if (tablesSeen > 0) {
    Logger.log('collectDueMarkers: %s table(s) in this module contain due-by ' +
               'wording and were LEFT IN PLACE. If the placeholder band is a ' +
               'table, that is why it survived.', tablesSeen);
  }
  // Pick up a "Display header as:" paragraph adjacent to a marked line. Keying
  // on adjacency keeps the blast radius tight: the phrase belongs to this band
  // and nothing else, but there is no reason to hunt for it across the module.
  const idxs = marks.slice();
  for (const i of marks) {
    for (const j of [i - 1, i + 1]) {
      if (j < span.start || j >= span.end || idxs.indexOf(j) !== -1) continue;
      const sib = body.getChild(j);
      if (sib.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      if (DISPLAY_AS_RE.test(sib.asParagraph().getText())) idxs.push(j);
    }
  }
  return idxs.map(i => body.getChild(i));
}
// Detach previously collected marker elements. Element references stay valid
// across insertions elsewhere in the body, which is why collectDueMarkers hands
// back elements: no index arithmetic survives an insert, but these do.
function removeDueMarkers(elements) {
  let removed = 0;
  for (const el of elements) {
    try { el.removeFromParent(); removed++; }
    catch (e) {
      Logger.log('removeDueMarkers: could not remove element: %s', (e && e.message) || e);
    }
  }
  Logger.log('removeDueMarkers: removed %s old due-by marker(s).', removed);
  return removed;
}
// ── PLACE DUE-DAY HEADERS ─────────────────────────────────────────
// Insert fresh due-by headers from the Course Pattern Table, and retire the
// module's existing markers.
//
// The template ships every module with placeholder markers — "Due by
// [Select Day] at 11:59 p.m. Mountain Time" plus a "Display header as:" chip.
// Leaving those in place stacked them directly above the headers this tool
// inserts, so QA read every due date twice.
//
// Wiping the module rather than trying to tell a placeholder from a previously
// inserted header is deliberate. Sniffing for the placeholder leaves stale
// headers behind the moment a due day changes between runs: the old day no
// longer matches the new one, so nothing removes it and the doubling returns —
// this time in a form that looks legitimate, which is worse. Rebuilding makes
// the module match the pattern table exactly on every run, by construction, and
// it is how the rest of this tool already treats modules 1..N (fillSlot
// overwrites titles outright; removeSlot deletes surplus slots).
//
// GATE: nothing is removed unless this module actually has headers going in.
// A run made while the pattern table's due-day column is still blank — an
// ordinary mid-build state — would otherwise strip the placeholders and put
// nothing back. Apps Script cannot recreate a dropdown chip, so that is not
// undoable inside the tool; recovery means re-copying the module by hand.
function placeDueHeaders(body, slotParas, activities, params, stats) {
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  // hasOwnProperty, not a bare lookup: google.script.run is callable directly,
  // so params.canvasOption can be anything. A bare lookup of "constructor" or
  // "toString" resolves up the prototype chain to a truthy function, which the
  // || fallback does not catch — the function body would then be stringified
  // into the heading and throw the annotation's styling offsets off. The same
  // reasoning already guards numModules and toolOverrides upstream.
  const CANVAS_TEXT = {
    display:     'Text Header in Canvas',
    doNotBuild:  'Do not build in Canvas',
    unpublished: 'Unpublished text header in Canvas'
  };
  const canvasText = Object.prototype.hasOwnProperty.call(CANVAS_TEXT, params.canvasOption)
    ? CANVAS_TEXT[params.canvasOption]
    : 'Text Header in Canvas';
  const groups = getDueDayGroups(activities);
  if (groups.length === 0) return 0;
  const targets = [];
  for (const { day, startIndex } of groups) {
    const targetPara = startIndex < slotParas.length ? slotParas[startIndex] : null;
    if (targetPara) targets.push({ day, targetPara });
  }
  // The gate. Also covers a module with fewer slots than the pattern expects:
  // no slot to hang a header on means nothing lands here, so nothing is cleared.
  if (targets.length === 0) return 0;
  // Collect the old markers, INSERT the replacements, and only then remove what
  // was collected. Ordering matters against the 6-minute execution limit: Docs
  // persists as it goes, so a run that dies mid-loop after clearing but before
  // inserting would leave a module stripped with nothing put back — and the
  // chips it stripped cannot be recreated by script. This ordering makes the
  // worst case a duplicate instead, which is merely the bug this change set out
  // to fix, and is recoverable by re-running.
  const span     = getModuleSpan(body, slotParas);
  const oldMarks = span ? collectDueMarkers(body, span) : [];
  for (let i = targets.length - 1; i >= 0; i--) {
    const { day, targetPara } = targets[i];
    const childIdx = body.getChildIndex(targetPara);
    // The Canvas annotation rides on the SAME paragraph as the due-by header,
    // separated by one space. As its own paragraph it wrapped to the next line
    // and fell outside the header's blue top/bottom borders.
    const headerText = `Due by ${day} at 11:59 p.m. Mountain Time`;
    const headerPara = body.insertParagraph(childIdx, `${headerText} ${canvasText}`);
    headerPara.setHeading(H3);
    const headerTxt = headerPara.editAsText();
    _fmt(headerTxt, { font: FONT, size: 15, bold: true, italic: false, color: DEEP_BLUE });
    // Re-style just the annotation run (header + the separating space).
    const annotStart = headerText.length + 1;
    const annotEnd   = annotStart + canvasText.length - 1;
    headerTxt.setFontSize(annotStart, annotEnd, 11);
    headerTxt.setBold(annotStart, annotEnd, false);
    headerTxt.setItalic(annotStart, annotEnd, true);
    headerTxt.setForegroundColor(annotStart, annotEnd, RED);
  }
  // Replacements are in; now retire the originals.
  stats.headersRemoved += removeDueMarkers(oldMarks);
  return targets.length;
}
// ── BLUE BORDERS ON DUE-BY HEADERS ────────────────────────────────
// QA wants a blue horizontal line above and below each "Due by …" header,
// matching the eCampus template. The DocumentApp API cannot set paragraph
// borders, so this runs as a Phase-2 post-pass via the Docs advanced service.
//
// It borders EVERY due-by header in the Development tab. A paragraph border is
// a single-valued style attribute (one borderTop / one borderBottom per
// paragraph), so re-applying the same spec to a header that already has it is
// a harmless no-op — no stacking, no visible change. That lets us skip any
// new-vs-existing tracking and simply normalize them all to the same look.
//
// Border spec (sampled from the template): #0000E7, 1.5 pt, 2 pt padding,
// solid, top + bottom.
function applyDueHeaderBorders(doc, devTabTitle) {
  // Persist the DocumentApp inserts so the Docs API sees the new headers.
  const docId = doc.getId();
  doc.saveAndClose();

  // The module edits are now saved. The border pass is purely cosmetic, so any
  // Docs-API failure here (advanced service disabled, permissions, transient
  // error) must NOT surface as a full-run failure — log it and return.
  try {
    borderDueHeaders_(docId, devTabTitle);
  } catch (e) {
    Logger.log('applyDueHeaderBorders: border pass failed (run already saved): ' +
               ((e && e.message) || e));
  }
}
// Docs-API worker for applyDueHeaderBorders. Kept separate so the caller can
// guard it in try/catch without deep-nesting the whole body.
function borderDueHeaders_(docId, devTabTitle) {
  // Re-read via the Docs API. Locate the Development tab BY TITLE (the
  // DocumentApp tab id and the Docs API tabId are not guaranteed to match, so
  // we don't rely on that) and read its own tabId for the ranges.
  const structured = Docs.Documents.get(docId, { includeTabsContent: true });
  const allTabs    = flattenDocsTabs_(structured);
  // Log counts/ids only — tab titles can encode course/program identifiers and
  // would then persist in Cloud Logging beyond the document's sharing scope.
  Logger.log('applyDueHeaderBorders: Docs API returned %s tab(s).', allTabs.length);

  let tab = null;
  for (const t of allTabs) {                 // exact title first
    if (t.tabProperties && t.tabProperties.title === devTabTitle) { tab = t; break; }
  }
  if (!tab) {                                // fall back to a loose match
    for (const t of allTabs) {
      if (t.tabProperties && /\bdevelopment\b/i.test(t.tabProperties.title || '')) { tab = t; break; }
    }
  }
  if (!tab) {
    Logger.log('applyDueHeaderBorders: Development tab not found in Docs API response.');
    return;
  }

  const tabId   = tab.tabProperties.tabId;
  const content = (tab.documentTab && tab.documentTab.body && tab.documentTab.body.content) || [];

  const BLUE = {
    color:     { color: { rgbColor: { red: 0, green: 0, blue: 0.90588 } } },
    width:     { magnitude: 1.5, unit: 'PT' },
    padding:   { magnitude: 2,   unit: 'PT' },
    dashStyle: 'SOLID'
  };

  // 3. One border request per due-by header in the tab.
  const requests = [];
  for (const el of content) {
    if (!el.paragraph || !el.paragraph.elements) continue;
    const text = el.paragraph.elements
      .map(e => (e.textRun && e.textRun.content) || '').join('');
    if (!DUE_HEADER_RE.test(text)) continue;
    requests.push({
      updateParagraphStyle: {
        range: { startIndex: el.startIndex, endIndex: el.endIndex, tabId: tabId },
        paragraphStyle: { borderTop: BLUE, borderBottom: BLUE },
        fields: 'borderTop,borderBottom'
      }
    });
  }

  // 4. Apply. updateParagraphStyle does not change text length, so there is
  //    no index drift and request order is irrelevant.
  Logger.log('applyDueHeaderBorders: bordering %s due-by header(s) (tab id %s).',
             requests.length, tabId);
  if (requests.length > 0) {
    Docs.Documents.batchUpdate({ requests: requests }, docId);
    Logger.log('applyDueHeaderBorders: batchUpdate applied.');
  }
}
// Flatten a Docs API document's tabs (including nested child tabs) into a
// single ordered array.
function flattenDocsTabs_(structuredDoc) {
  const out = [];
  function walk(list) {
    (list || []).forEach(function (t) { out.push(t); walk(t.childTabs); });
  }
  walk(structuredDoc.tabs);
  return out;
}
// ── FORMAT HELPER ─────────────────────────────────────────────────
function _fmt(textEl, opts) {
  if (opts.font   !== undefined) textEl.setFontFamily(opts.font);
  if (opts.size   !== undefined) textEl.setFontSize(opts.size);
  if (opts.bold   !== undefined) textEl.setBold(opts.bold);
  if (opts.italic !== undefined) textEl.setItalic(opts.italic);
  if (opts.color  !== undefined) textEl.setForegroundColor(opts.color);
  return textEl;
}
// Paint a Canvas-tool line: Arial 11 bold red throughout, except the marker,
// which is black. Also strips the old grey chip if the line is carrying one.
//
// Every place that writes this line goes through here, so the two paths that
// fill a slot — building it fresh and swapping a real tool into an existing
// "Select Tool" line — can no longer drift apart in formatting.
function _fmtToolLine(para) {
  const pt   = para.editAsText();
  const text = pt.getText();
  if (text.length === 0) return;
  _fmt(pt, { font: FONT, size: 11, bold: true, italic: false, color: RED });
  const sepIdx = text.indexOf(TOOL_MARKER);
  if (sepIdx !== -1) pt.setForegroundColor(sepIdx, sepIdx, TOOL_MARKER_COLOR);
  _clearGreyChip(pt, text.length);
}
// Remove the old grey chip from a tool line, and ONLY the old grey chip.
//
// A blanket setBackgroundColor(null) would also wipe the light-green "ready for
// Canvas" highlight QA applies by hand — the very marking this whole change
// exists to protect. So the colour is checked first and the line is left alone
// unless it is carrying our grey.
//
// The chip always started at offset 0, and a highlight is applied to the whole
// line, so testing the first character is enough to tell grey from green. That
// keeps this to two DocumentApp round trips on affected lines and one on the
// rest, rather than a per-character scan on every activity in the course.
// Takes the Text element and its length rather than the Paragraph: the only
// caller has both in hand already, so re-deriving them here was two wasted
// round trips on every tool line in the course.
function _clearGreyChip(pt, len) {
  if (len === 0) return;
  if (pt.getBackgroundColor(0) !== GREY_CHIP) return;
  pt.setBackgroundColor(0, len - 1, null);
}
// ── BUILD SUMMARY ─────────────────────────────────────────────────
function buildSummary(stats, params, activities, numModules) {
  // Tool cells we still couldn't interpret — either the user chose "leave
  // unchanged" in the pre-flight panel or skipped it. These activities keep the
  // "Select Tool" placeholder, so name the cell text here: the fix belongs in
  // the Course Pattern Table, and this run's pick doesn't persist back to the
  // Design tab.
  const unknownTools = [];
  for (const act of activities) {
    if (act.tool || !act.rawTool) continue;
    if (unknownTools.indexOf(act.rawTool) === -1) unknownTools.push(act.rawTool);
  }
  const unknownNote = unknownTools.length === 0 ? null :
    '\n⚠ Unrecognized Canvas Tool, still "Select Tool": ' +
    unknownTools.map(t => `"${t}"`).join(', ') +
    '\n   Correct these in the Course Pattern Table to fix them permanently.';

  return [
    '✅ Blueprint Development Tab Updated!',
    '',
    `Modules: ${numModules} total`,
    stats.created      > 0 ? `  + ${stats.created} new module(s) created`  : null,
    stats.deleted      > 0 ? `  − ${stats.deleted} module(s) removed`       : null,
    `Activities set: ${stats.filled}`,
    `Tools assigned: ${stats.tools}`,
    stats.slotsDeleted > 0 ? `Extra slots removed: ${stats.slotsDeleted}`   : null,
    `Due-day headers inserted: ${stats.headers}`,
    stats.headersRemoved > 0 ? `  − ${stats.headersRemoved} old due-date marker(s) replaced` : null,
    '',
    `Numbered: ${params.numbered ? 'Yes' : 'No'}`,
    `Time estimates: ${params.timeEstimates ? 'Yes' : 'No'}`,
    `Canvas option: ${params.canvasOption}`,
    unknownNote
  ].filter(l => l !== null).join('\n');
}
