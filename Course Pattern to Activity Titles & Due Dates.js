/**
 * ====================================================================
 * COURSE PATTERN → ACTIVITY TITLES & DUE DAYS  |  Apps Script  v2.0
 * ====================================================================
 *
 * WHAT IT DOES
 *   • Replaces "Activity Title" placeholders with activity names
 *   • Sets the "Select Tool" dropdown for each activity
 *   • Deletes extra activity slots beyond the course pattern
 *   • Places "Due by [Day]" headers before each due-day group
 *     (creates additional headers if more due days than the 2 in template)
 *
 * INSTALLATION
 *   1. Extensions → Apps Script in your Blueprint Google Doc
 *   2. Replace all existing code with this script, Save (Ctrl+S)
 *   3. Rename the project to "Course Pattern → Activity Titles & Due Days"
 *   4. Reload the Google Doc — "🎓 Blueprint Tools" menu appears
 *   5. Click  🎓 Blueprint Tools → Populate Development Tab
 *
 * NOTES
 *   • Slots already filled in (title ≠ "Activity Title") are skipped.
 *   • If the course is > 7 weeks, a warning is shown — manually copy
 *     the module block for extra weeks.
 *   • New "Due by" headers are plain text (the original dropdown chips
 *     can't be recreated by script). Re-add the "Display header as:"
 *     dropdown chip manually via Insert → Dropdown if needed.
 * ====================================================================
 */


// ── TOOL CLASSIFICATION ─────────────────────────────────────────────
const TOOL_RULES = [
  {
    tool: 'Page',
    keywords: ['overview', 'reading', 'readings', 'video', 'videos',
               'watch', 'lecture', 'content', 'introduction', 'intro',
               'module overview', 'week overview']
  },
  {
    tool: 'Discussion',
    keywords: ['discussion', 'forum', 'board']
  },
  {
    tool: 'Quiz (Classic)',
    keywords: ['quiz', 'test', 'exam', 'midterm', 'final exam', 'knowledge check']
  },
  {
    tool: 'Assignment',
    keywords: ['assignment', 'paper', 'reflection', 'project', 'essay',
               'report', 'writing', 'submission', 'lab', 'homework',
               'journal', 'portfolio', 'worksheet', 'response']
  }
];

function mapToTool(activityName) {
  const lower = activityName.toLowerCase().trim();
  for (const { keywords, tool } of TOOL_RULES) {
    for (const kw of keywords) {
      if (lower.includes(kw)) return tool;
    }
  }
  return null; // Unknown — leave dropdown unchanged
}


// ── DAY NORMALIZATION ────────────────────────────────────────────────
// Sorted longest-first to avoid "thu" matching before "thurs", etc.
const DAY_ALIASES = [
  ['thursday',  'Thursday'],
  ['tuesday',   'Tuesday'],
  ['saturday',  'Saturday'],
  ['wednesday', 'Wednesday'],
  ['monday',    'Monday'],
  ['tuesday',   'Tuesday'],
  ['sunday',    'Sunday'],
  ['friday',    'Friday'],
  ['thurs',     'Thursday'],
  ['tues',      'Tuesday'],
  ['thur',      'Thursday'],
  ['wed',       'Wednesday'],
  ['mon',       'Monday'],
  ['fri',       'Friday'],
  ['sat',       'Saturday'],
  ['sun',       'Sunday'],
  ['thu',       'Thursday'],
  ['tue',       'Tuesday'],
];

function parseDueDay(text) {
  if (!text) return null;
  for (const [alias, full] of DAY_ALIASES) {
    if (new RegExp(`\\b${alias}\\b`, 'i').test(text)) return full;
  }
  return null;
}


// ── MENU ────────────────────────────────────────────────────────────
function onOpen() {
  DocumentApp.getUi()
    .createMenu('🎓 Blueprint Tools')
    .addItem('Populate Development Tab', 'populateDevelopmentTab')
    .addToUi();
}


// ── MAIN FUNCTION ────────────────────────────────────────────────────
function populateDevelopmentTab() {
  const ui  = DocumentApp.getUi();
  const doc = DocumentApp.getActiveDocument();

  const go = ui.alert(
    '🎓 Blueprint Activity Populator',
    'This will:\n' +
    '  • Replace "Activity Title" placeholders\n' +
    '  • Set "Select Tool" dropdowns\n' +
    '  • Delete extra activity slots\n' +
    '  • Add "Due by [Day]" headers from the course pattern\n\n' +
    'Slots already filled in will be skipped.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (go !== ui.Button.YES) return;

  // ── 1. Find tabs ─────────────────────────────────────────────────
  const allTabs    = collectTabs(doc);
  const dashboard  = allTabs.find(t => /dashboard|project/i.test(t.title));
  const designTab  = allTabs.find(t => /^design$/i.test(t.title) || /\bdesign\b/i.test(t.title));
  const devTab     = allTabs.find(t => /^development$/i.test(t.title) || /\bdevelopment\b/i.test(t.title));

  if (!designTab || !devTab) {
    ui.alert('Tab Error',
      'Could not find "Design" and/or "Development" tabs.\n' +
      'Make sure the tab names match (case-insensitive).',
      ui.ButtonSet.OK);
    return;
  }

  // ── 2. Detect course length ──────────────────────────────────────
  const numWeeks = detectCourseLength((dashboard || designTab).body);
  Logger.log('Course length: ' + numWeeks + ' weeks');

  // ── 3. Parse course pattern (all matching tables combined) ───────
  const activities = parseCoursePattern(designTab.body);
  if (activities.length === 0) {
    ui.alert('Parse Error',
      'No activities found in the course pattern table(s) in the Design tab.\n' +
      'Make sure the table header contains "Activity" or "Assessment".',
      ui.ButtonSet.OK);
    return;
  }
  Logger.log('Activities: ' + activities.map(a => `${a.name} (${a.dueDay || 'no day'} → ${a.tool || '?'})`).join(' | '));

  // ── 4. Populate modules ──────────────────────────────────────────
  const stats  = { filled: 0, tools: 0, deleted: 0, headers: 0 };
  const inDoc  = Math.min(numWeeks, 7); // Template has 7 modules max

  for (let mod = 1; mod <= inDoc; mod++) {
    processModule(devTab.body, mod, activities, stats);
  }

  // ── 5. Summary ───────────────────────────────────────────────────
  let msg = '✅ Done!\n\n' +
            `  • ${stats.filled} activity title(s) updated\n` +
            `  • ${stats.tools} tool dropdown(s) set\n` +
            `  • ${stats.deleted} extra slot(s) removed\n` +
            `  • ${stats.headers} due-day header(s) placed\n\n` +
            `ℹ️ New "Due by" headers are plain text.\n` +
            `Re-add the "Display header as:" chip manually\n` +
            `(Insert → Dropdown) if needed.`;

  if (numWeeks > 7) {
    msg += `\n\n⚠️ This course is ${numWeeks} weeks but the template only has\n` +
           `7 modules. Manually duplicate the module block for weeks 8–${numWeeks}.`;
  }

  ui.alert('Blueprint Tools', msg, ui.ButtonSet.OK);
}


// ── COLLECT TABS ────────────────────────────────────────────────────
function collectTabs(doc) {
  const result = [];
  function walk(tab) {
    result.push({ title: tab.getTitle(), body: tab.asDocumentTab().getBody() });
    tab.getChildTabs().forEach(walk);
  }
  doc.getTabs().forEach(walk);
  return result;
}


// ── DETECT COURSE LENGTH ─────────────────────────────────────────────
function detectCourseLength(body) {
  const text = body.getText();
  const m1 = text.match(/\b(\d+)W\d*/i);       // "15W1", "7W", "5W2" …
  if (m1) return parseInt(m1[1], 10);
  const m2 = text.match(/\b(\d+)[- ]?week/i);  // "15-week", "15 week" …
  if (m2) return parseInt(m2[1], 10);
  Logger.log('Course length not detected — defaulting to 7');
  return 7;
}


// ── PARSE COURSE PATTERN TABLE(S) ───────────────────────────────────
/**
 * Scans ALL tables in the Design tab body. A table is treated as a
 * course pattern table when its first row has:
 *   col 0 → contains "activity" or "assessment"
 *   col 1 → contains "criteria", "due", or "day"
 * This avoids accidentally matching the Detailed Course Design Map.
 * Multiple matching tables are combined in order (handles the common
 * case where the course pattern is split across two tables).
 */
function parseCoursePattern(body) {
  const activities = [];

  for (const table of body.getTables()) {
    if (table.getNumRows() < 2 || table.getNumColumns() < 2) continue;

    const col0 = table.getRow(0).getCell(0).getText().toLowerCase();
    const col1 = table.getRow(0).getCell(1).getText().toLowerCase();

    const col0Match = col0.includes('activity') || col0.includes('assessment');
    const col1Match = col1.includes('criteria') || col1.includes('due') || col1.includes('day');
    if (!col0Match || !col1Match) continue;

    // This is a course pattern table — read every data row
    for (let r = 1; r < table.getNumRows(); r++) {
      const row  = table.getRow(r);
      const name = row.getCell(0).getText().trim();
      if (!name) continue;

      const dueDayText = row.getCell(1).getText().trim();
      activities.push({
        name,
        tool:   mapToTool(name),
        dueDay: parseDueDay(dueDayText)
      });
    }
    // No break — intentionally processes ALL matching tables
  }

  return activities;
}


// ── DUE DAY GROUPS ───────────────────────────────────────────────────
/**
 * Returns [{ day, startSlot }, …] — one entry per unique due day,
 * in the order that day first appears in the activities list.
 * startSlot is the 1-based index of the first activity with that day.
 */
function getDueDayGroups(activities) {
  const groups = [];
  const seen   = new Set();
  for (let i = 0; i < activities.length; i++) {
    const day = activities[i].dueDay;
    if (!day || seen.has(day)) continue;
    seen.add(day);
    groups.push({ day, startSlot: i + 1 });
  }
  return groups;
}


// ── PROCESS ONE MODULE ───────────────────────────────────────────────
function processModule(body, modNum, activities, stats) {
  const H4     = DocumentApp.ParagraphHeading.HEADING4;
  const prefix = modNum + '.';

  // Collect all HEADING4 paragraphs belonging to this module
  const slots = [];
  for (const para of body.getParagraphs()) {
    if (para.getHeading() !== H4) continue;
    const text = para.getText().trim();
    if (!text.startsWith(prefix)) continue;
    const slotNum = parseInt(text.split(' ')[0].split('.')[1], 10);
    if (!isNaN(slotNum)) slots.push({ slotNum, para });
  }
  slots.sort((a, b) => a.slotNum - b.slotNum);

  // Fill matching slots; queue extras for deletion
  const toDelete = [];
  for (const { slotNum, para } of slots) {
    if (slotNum <= activities.length) {
      fillSlot(body, para, activities[slotNum - 1], stats);
    } else {
      toDelete.push(para);
    }
  }

  // Delete extra slots bottom-up (avoids index drift)
  for (let i = toDelete.length - 1; i >= 0; i--) {
    removeSlot(body, toDelete[i]);
    stats.deleted++;
  }

  // Place due-day headers
  stats.headers += placeDueHeaders(body, modNum, activities);
}


// ── FILL ONE SLOT ────────────────────────────────────────────────────
function fillSlot(body, headingPara, activity, stats) {
  const text      = headingPara.getText().trim();
  const spaceIdx  = text.indexOf(' ');
  if (spaceIdx < 0) return;

  const numCode      = text.substring(0, spaceIdx);
  const currentTitle = text.substring(spaceIdx + 1).trim();

  // Only replace if the placeholder is still there
  if (currentTitle === 'Activity Title') {
    headingPara.replaceText('Activity Title', activity.name);
    stats.filled++;
    Logger.log(`Filled: ${numCode} → "${activity.name}"`);
  }

  // Set the nearest "Select Tool" dropdown
  if (activity.tool) {
    if (setNearbyTool(body, headingPara, activity.tool)) stats.tools++;
  }
}


// ── SET TOOL DROPDOWN ────────────────────────────────────────────────
/**
 * Scans up to 6 paragraphs after the activity heading for a line
 * containing "Select Tool" and replaces it with the correct tool name.
 *
 * Dropdown chips store their value as readable text; replaceText()
 * can update this. If a chip reverts when clicked, set it manually
 * (a future v3 could use the advanced Docs REST API for a harder fix).
 */
function setNearbyTool(body, headingPara, toolValue) {
  const paras = body.getParagraphs();
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const H4 = DocumentApp.ParagraphHeading.HEADING4;

  let idx = -1;
  for (let i = 0; i < paras.length; i++) {
    if (paras[i] === headingPara) { idx = i; break; }
  }
  if (idx < 0) return false;

  for (let i = idx + 1; i < Math.min(idx + 7, paras.length); i++) {
    const para = paras[i];
    const h    = para.getHeading();
    if (h === H2 || h === H3 || h === H4) break; // Left this activity's block
    if (para.getText().includes('Select Tool')) {
      try {
        para.replaceText('Select Tool', toolValue);
        Logger.log(`  Tool → ${toolValue}`);
        return true;
      } catch (e) {
        Logger.log(`  Could not set tool: ${e.message}`);
        return false;
      }
    }
  }
  return false;
}


// ── REMOVE EXTRA SLOT ────────────────────────────────────────────────
/**
 * Removes the HEADING4 line and all following body paragraphs
 * (estimated time, tool line, directions) up to the next heading.
 */
function removeSlot(body, headingPara) {
  const paras = body.getParagraphs();
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const H4 = DocumentApp.ParagraphHeading.HEADING4;

  let startIdx = -1;
  for (let i = 0; i < paras.length; i++) {
    if (paras[i] === headingPara) { startIdx = i; break; }
  }
  if (startIdx < 0) return;

  const toRemove = [paras[startIdx]];
  for (let i = startIdx + 1; i < paras.length; i++) {
    const h = paras[i].getHeading();
    if (h === H4 || h === H3 || h === H2) break;
    toRemove.push(paras[i]);
  }
  for (let i = toRemove.length - 1; i >= 0; i--) {
    try { toRemove[i].removeFromParent(); }
    catch (e) { Logger.log('Could not remove paragraph: ' + e.message); }
  }
  Logger.log(`Deleted slot: ${headingPara.getText()}`);
}


// ── PLACE DUE-DAY HEADERS ────────────────────────────────────────────
/**
 * For the given module:
 *   1. Finds and removes existing "Due by … Mountain Time" headers
 *   2. Inserts a fresh "Due by [Day] at 11:59 p.m. Mountain Time"
 *      header (HEADING3) immediately before the first activity of
 *      each due-day group found in the course pattern.
 *
 * Insertions are done in reverse group order so that earlier
 * paragraphs' child indices stay stable during the operation.
 *
 * Returns the number of headers inserted.
 */
function placeDueHeaders(body, modNum, activities) {
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;
  const H4 = DocumentApp.ParagraphHeading.HEADING4;

  // ── Find the paragraph span for this module ──────────────────────
  const allParas  = body.getParagraphs();
  const modRegex  = new RegExp(`^(Module|Week)\\s+${modNum}[:\\s]`, 'i');
  const nextRegex = new RegExp(`^(Module|Week)\\s+${modNum + 1}[:\\s]`, 'i');

  let modStart = -1, modEnd = allParas.length;
  for (let i = 0; i < allParas.length; i++) {
    if (allParas[i].getHeading() !== H2) continue;
    const t = allParas[i].getText().trim();
    if (modRegex.test(t)  && modStart < 0)         modStart = i;
    else if (nextRegex.test(t) && modStart >= 0)  { modEnd = i; break; }
  }
  if (modStart < 0) return 0;

  // ── Remove existing "Due by … Mountain Time" headers ────────────
  const oldHeaders = allParas.slice(modStart, modEnd).filter(p =>
    p.getHeading() === H3 &&
    /due\s+by\b/i.test(p.getText()) &&
    /mountain\s+time/i.test(p.getText())
  );
  for (let i = oldHeaders.length - 1; i >= 0; i--) {
    oldHeaders[i].removeFromParent();
  }
  Logger.log(`Module ${modNum}: removed ${oldHeaders.length} old header(s)`);

  // ── Build due-day groups ─────────────────────────────────────────
  const groups = getDueDayGroups(activities);
  if (groups.length === 0) return 0;

  // ── Collect target activity paragraphs (fresh fetch after deletions)
  const freshParas = body.getParagraphs();
  const targets = [];

  for (const { day, startSlot } of groups) {
    const slotCode = `${modNum}.${String(startSlot).padStart(2, '0')}`;
    for (const para of freshParas) {
      if (para.getHeading() === H4 &&
          para.getText().trim().startsWith(slotCode + ' ')) {
        targets.push({ day, targetPara: para });
        break;
      }
    }
  }

  // ── Insert headers in reverse order (last group first) ───────────
  for (let i = targets.length - 1; i >= 0; i--) {
    const { day, targetPara } = targets[i];
    const childIdx  = body.getChildIndex(targetPara);
    const headerTxt = `Due by ${day} at 11:59 p.m. Mountain Time`;
    const newPara   = body.insertParagraph(childIdx, headerTxt);
    newPara.setHeading(H3);
    Logger.log(`Module ${modNum}: inserted "Due by ${day}" before slot ${targets[i].day}`);
  }

  return targets.length;
}
