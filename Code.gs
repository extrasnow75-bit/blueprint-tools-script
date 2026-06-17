/**
 * ================================================================
 * BLUEPRINT TOOLS  |  Google Apps Script  v4.2 + removeSlot fix + formatting fixes
 * ================================================================
 */
const RED       = '#ff0000';
const DEEP_BLUE = '#0033a0';
const BLACK     = '#000000';
const FONT      = 'Arial';
// ── MENU ──────────────────────────────────────────────────────────
function onOpen() {
  DocumentApp.getUi()
    .createMenu('🎓 Blueprint Tools')
    .addItem('Populate Development Tab', 'showSidebar')
    .addSeparator()
    .addItem('Add Activity Directions', 'showDirectionsSidebar')
    .addSeparator()
    .addItem('Time Estimator', 'showTimeEstimatorSidebar')
    .addToUi();
}
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('🎓 Blueprint Tools')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}
// ── MAIN ──────────────────────────────────────────────────────────
function processBlueprint(params) {
  const doc  = DocumentApp.getActiveDocument();
  const tabs = collectTabs(doc);
  const designTab = tabs.find(t => /\bdesign\b/i.test(t.title));
  const devTab    = tabs.find(t => /\bdevelopment\b/i.test(t.title));
  if (!designTab || !devTab)
    throw new Error('Could not find "Design" and/or "Development" tabs.');
  const activities = parseCoursePattern(designTab.body);
  if (activities.length === 0)
    throw new Error('No activities found in the course pattern table.');
  const stats      = { created: 0, deleted: 0, filled: 0, tools: 0, slotsDeleted: 0, headers: 0 };
  const numModules = params.numModules;
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
function parseCoursePattern(body) {
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
      activities.push({
        name,
        tool:   normalizeToolName(toolCol < n ? row.getCell(toolCol).getText().trim() : ''),
        dueDay: parseDueDay(dayCol  < n ? row.getCell(dayCol).getText().trim()  : ''),
        time:   timeCol < n ? row.getCell(timeCol).getText().trim() : ''
      });
    }
  }
  return activities;
}
// ── NORMALIZE TOOL NAME ───────────────────────────────────────────
function normalizeToolName(raw) {
  const t = raw.toLowerCase().trim();
  if (!t) return null;
  if (t.includes('not graded') || t.includes('(not'))   return 'Assignment (Not Graded)';
  if (t.includes('assignment'))                          return 'Assignment';
  if (t.includes('discussion'))                          return 'Discussion';
  if (t.includes('page'))                                return 'Page';
  if (t.includes('quiz') && t.includes('new'))           return 'Quiz (New)';
  if (t.includes('quiz'))                                return 'Quiz (Classic)';
  return null;
}
// ── PARSE DUE DAY ─────────────────────────────────────────────────
const DAY_ALIASES = [
  ['thursday','Thursday'],['tuesday','Tuesday'],['saturday','Saturday'],
  ['wednesday','Wednesday'],['monday','Monday'],['sunday','Sunday'],
  ['friday','Friday'],['thurs','Thursday'],['tues','Tuesday'],
  ['thur','Thursday'],['wed','Wednesday'],['mon','Monday'],
  ['fri','Friday'],['sat','Saturday'],['sun','Sunday'],
  ['thu','Thursday'],['tue','Tuesday']
];
function parseDueDay(text) {
  if (!text) return null;
  for (const [alias, full] of DAY_ALIASES) {
    if (new RegExp(`\\b${alias}\\b`, 'i').test(text)) return full;
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
    // Select Tool; Link to settings tab
    const tPara = add('Select Tool; Link to settings tab');
    tPara.setHeading(NORMAL);
    tPara.setIndentStart(indent);
    _fmt(tPara.editAsText(), { font: FONT, size: 11, bold: true, color: RED });
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
  const slots = getSlotsInModule(body, modNum);
  for (const { slotNum, para } of slots) {
    if (slotNum <= activities.length) {
      fillSlot(body, para, modNum, slotNum, activities[slotNum - 1], params, indent, stats);
    } else {
      removeSlot(body, para);
      stats.slotsDeleted++;
    }
  }
  stats.headers += placeDueHeaders(body, modNum, activities, params);
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
    const t = para.getText().trim();
    if (h === H2) {
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
}
// ── SET TOOL WITH FORMATTING ──────────────────────────────────────
function setNearbyTool(body, headingPara, toolValue) {
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const H4 = DocumentApp.ParagraphHeading.HEADING4;
  const startIdx   = body.getChildIndex(headingPara);
  if (startIdx < 0) return false;
  const suffix      = '; Link to settings tab';
  const numChildren = body.getNumChildren();
  for (let i = startIdx + 1; i < Math.min(startIdx + 8, numChildren); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = child.asParagraph();
    const h    = para.getHeading();
    if (h === H2 || h === H3 || h === H4) break;
    const text          = para.getText();
    const hasSelectTool = text.includes('Select Tool');
    const hasSuffix     = text.includes(suffix);
    if (!hasSelectTool && !hasSuffix) continue;
    try {
      if (hasSelectTool) {
        para.replaceText('Select Tool', toolValue);
        const updated = para.getText();
        const toolEnd = updated.includes(';') ? updated.indexOf(';') - 1 : updated.length - 1;
        if (toolEnd >= 0) {
          const pt = para.editAsText();
          pt.setFontFamily(0, toolEnd, FONT);
          pt.setFontSize(0, toolEnd, 11);
          pt.setBold(0, toolEnd, true);
          pt.setForegroundColor(0, toolEnd, RED);
        }
      } else if (hasSuffix) {
        para.setText(toolValue + suffix);
        _fmt(para.editAsText(), { font: FONT, size: 11, bold: true, color: RED });
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
// ── PLACE DUE-DAY HEADERS ─────────────────────────────────────────
function placeDueHeaders(body, modNum, activities, params) {
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const canvasText = {
    display:     'Text Header in Canvas',
    doNotBuild:  'Do not build in Canvas',
    unpublished: 'Unpublished text header in Canvas'
  }[params.canvasOption] || 'Text Header in Canvas';
  // NOTE: This function used to scan the module for existing "Due by … Mountain
  // Time" headers and delete them before inserting new ones. Per request, the
  // template's original due-date markers are now left in place untouched —
  // only new headers are inserted at the correct activity positions below.
  const groups = getDueDayGroups(activities);
  if (groups.length === 0) return 0;
  const targets = [];
  for (const { day, startIndex } of groups) {
    const targetPara = getNthSlotPara(body, modNum, startIndex + 1);
    if (targetPara) targets.push({ day, targetPara });
  }
  for (let i = targets.length - 1; i >= 0; i--) {
    const { day, targetPara } = targets[i];
    const childIdx = body.getChildIndex(targetPara);
    const annotPara = body.insertParagraph(childIdx, canvasText);
    annotPara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    _fmt(annotPara.editAsText(), { font: FONT, size: 11, italic: true, color: RED });
    const headerPara = body.insertParagraph(childIdx, `Due by ${day} at 11:59 p.m. Mountain Time`);
    headerPara.setHeading(H3);
    _fmt(headerPara.editAsText(), { font: FONT, size: 15, bold: true, color: DEEP_BLUE });
  }
  return targets.length;
}
// ── GET NTH H4 SLOT IN MODULE ─────────────────────────────────────
function getNthSlotPara(body, modNum, slotNum) {
  const H2    = DocumentApp.ParagraphHeading.HEADING2;
  const H4    = DocumentApp.ParagraphHeading.HEADING4;
  const modRe = new RegExp(`^Module\\s+${modNum}[:\\s]`, 'i');
  let inModule = false, count = 0;
  for (const para of body.getParagraphs()) {
    const h = para.getHeading();
    const t = para.getText().trim();
    if (h === H2) {
      if (modRe.test(t)) { inModule = true; continue; }
      if (inModule) return null;
    }
    if (!inModule || h !== H4) continue;
    if (++count === slotNum) return para;
  }
  return null;
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
// ── BUILD SUMMARY ─────────────────────────────────────────────────
function buildSummary(stats, params, activities, numModules) {
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
    '',
    `Numbered: ${params.numbered ? 'Yes' : 'No'}`,
    `Time estimates: ${params.timeEstimates ? 'Yes' : 'No'}`,
    `Canvas option: ${params.canvasOption}`
  ].filter(l => l !== null).join('\n');
}
