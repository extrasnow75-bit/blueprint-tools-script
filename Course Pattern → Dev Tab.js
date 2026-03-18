/**
 * ====================================================================
 * BLUEPRINT ACTIVITY POPULATOR  |  Google Apps Script  v1.0
 * ====================================================================
 *
 * PURPOSE
 *   Reads the course pattern table from the Design tab, then fills in:
 *     • Activity titles  (replaces "Activity Title" placeholders)
 *     • Tool dropdowns   (replaces "Select Tool" with Canvas tool type)
 *   Extra activity slots beyond the course pattern are deleted.
 *
 * INSTALLATION
 *   1. In the Blueprint Google Doc: Extensions → Apps Script
 *   2. Delete any existing code, paste this entire script, Save (Ctrl+S)
 *   3. Reload the Google Doc — a "🎓 Blueprint Tools" menu appears
 *   4. Click  🎓 Blueprint Tools → Populate Development Tab
 *   5. Authorize the script when prompted, then run it again
 *
 * NOTES
 *   • Slots already filled in (title ≠ "Activity Title") are skipped.
 *   • If the course is > 7 weeks, a warning is shown. Copy the module
 *     block manually for extra weeks — dropdown chips cannot be created
 *     programmatically in v1.
 *   • "Select Tool" dropdowns: replaceText() works if the chip exposes
 *     its value as editable text. If a chip reverts when clicked,
 *     select it manually (or request a v2 with the advanced Docs API).
 * ====================================================================
 */

// ── TOOL CLASSIFICATION ─────────────────────────────────────────────
// Add or adjust keywords here to tune tool detection.
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
    keywords: ['quiz', 'test', 'exam', 'midterm', 'final exam',
               'knowledge check']
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

  // Confirm before making changes
  const go = ui.alert(
    '🎓 Blueprint Activity Populator',
    'This will:\n' +
    '  • Replace "Activity Title" placeholders in the Development tab\n' +
    '  • Set each "Select Tool" dropdown to the appropriate type\n' +
    '  • Delete extra activity slots with no matching course activity\n\n' +
    'Slots already filled in will be skipped.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (go !== ui.Button.YES) return;

  // ── 1. Find tabs ─────────────────────────────────────────────────
  const allTabs = collectTabs(doc);

  const dashboard = allTabs.find(t => /dashboard|project/i.test(t.title));
  const designTab = allTabs.find(t => /^design$/i.test(t.title) || /\bdesign\b/i.test(t.title));
  const devTab    = allTabs.find(t => /^development$/i.test(t.title) || /\bdevelopment\b/i.test(t.title));

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

  // ── 3. Parse course pattern ──────────────────────────────────────
  const activities = parseCoursePattern(designTab.body);
  if (activities.length === 0) {
    ui.alert('Parse Error',
      'No activities found in the course pattern table in the Design tab.\n' +
      'Make sure the table header contains "Activity" or "Assessment".',
      ui.ButtonSet.OK);
    return;
  }
  Logger.log('Activities: ' + activities.map(a => `${a.name} → ${a.tool || '?'}`).join(' | '));

  // ── 4. Populate modules ──────────────────────────────────────────
  const stats  = { filled: 0, tools: 0, deleted: 0 };
  const inDoc  = Math.min(numWeeks, 7); // Template has 7 modules max

  for (let mod = 1; mod <= inDoc; mod++) {
    processModule(devTab.body, mod, activities, stats);
  }

  // ── 5. Report ────────────────────────────────────────────────────
  let msg = '✅ Done!\n\n' +
            `  • ${stats.filled} activity title(s) updated\n` +
            `  • ${stats.tools} tool dropdown(s) set\n` +
            `  • ${stats.deleted} extra slot(s) removed`;

  if (numWeeks > 7) {
    msg += `\n\n⚠️ This course is ${numWeeks} weeks long but the template only ` +
           `has 7 modules.\nPlease manually duplicate the module block in the ` +
           `Development tab for weeks 8–${numWeeks}.`;
  }

  ui.alert('Blueprint Tools', msg, ui.ButtonSet.OK);
}


// ── COLLECT ALL TABS ─────────────────────────────────────────────────
/**
 * Returns flat array of { title, body } for every tab (including children).
 */
function collectTabs(doc) {
  const result = [];
  function walk(tab) {
    result.push({
      title: tab.getTitle(),
      body:  tab.asDocumentTab().getBody()
    });
    tab.getChildTabs().forEach(walk);
  }
  doc.getTabs().forEach(walk);
  return result;
}


// ── DETECT COURSE LENGTH ─────────────────────────────────────────────
function detectCourseLength(body) {
  const text = body.getText();

  // "15W1", "7W", "5W2", "8W" …
  const m1 = text.match(/\b(\d+)W\d*/i);
  if (m1) return parseInt(m1[1], 10);

  // "15-week", "15 week", "15week" …
  const m2 = text.match(/\b(\d+)[- ]?week/i);
  if (m2) return parseInt(m2[1], 10);

  Logger.log('Course length not detected — defaulting to 7');
  return 7;
}


// ── PARSE COURSE PATTERN TABLE ───────────────────────────────────────
/**
 * Finds the first table whose header row contains "activity" or
 * "assessment" and returns rows 1-N as activity objects.
 */
function parseCoursePattern(body) {
  const activities = [];

  for (const table of body.getTables()) {
    if (table.getNumRows() < 2) continue;
    const header = table.getRow(0).getCell(0).getText().toLowerCase();
    if (!header.includes('activity') && !header.includes('assessment')) continue;

    for (let r = 1; r < table.getNumRows(); r++) {
      const name = table.getRow(r).getCell(0).getText().trim();
      if (name) activities.push({ name, tool: mapToTool(name) });
    }
    break; // Only use the first matching table
  }

  return activities;
}


// ── PROCESS ONE MODULE ───────────────────────────────────────────────
function processModule(body, modNum, activities, stats) {
  const H4     = DocumentApp.ParagraphHeading.HEADING4;
  const prefix = modNum + '.';

  // Gather all HEADING4 paragraphs that belong to this module (e.g. "3.02 …")
  const slots = [];
  for (const para of body.getParagraphs()) {
    if (para.getHeading() !== H4) continue;
    const text = para.getText().trim();
    if (!text.startsWith(prefix)) continue;

    const code    = text.split(' ')[0];           // "3.02"
    const slotNum = parseInt(code.split('.')[1], 10);
    if (!isNaN(slotNum)) slots.push({ slotNum, para });
  }

  slots.sort((a, b) => a.slotNum - b.slotNum);

  // Split into fill vs. delete
  const toDelete = [];
  for (const { slotNum, para } of slots) {
    if (slotNum <= activities.length) {
      fillSlot(body, para, activities[slotNum - 1], stats);
    } else {
      toDelete.push(para);
    }
  }

  // Delete extra slots from the bottom up (avoids index-shift issues)
  for (let i = toDelete.length - 1; i >= 0; i--) {
    removeSlot(body, toDelete[i]);
    stats.deleted++;
  }
}


// ── FILL ONE SLOT ────────────────────────────────────────────────────
function fillSlot(body, headingPara, activity, stats) {
  const text      = headingPara.getText().trim();
  const spaceIdx  = text.indexOf(' ');
  if (spaceIdx < 0) return;

  const numCode      = text.substring(0, spaceIdx);          // "1.04"
  const currentTitle = text.substring(spaceIdx + 1).trim();

  // Only replace if the placeholder is still there
  if (currentTitle === 'Activity Title') {
    headingPara.replaceText('Activity Title', activity.name);
    stats.filled++;
    Logger.log(`Filled: ${numCode} → "${activity.name}"`);
  }

  // Set the nearest "Select Tool" dropdown
  if (activity.tool) {
    const toolSet = setNearbyTool(body, headingPara, activity.tool);
    if (toolSet) stats.tools++;
  }
}


// ── SET TOOL DROPDOWN ────────────────────────────────────────────────
/**
 * Scans up to 6 paragraphs after the activity heading for a paragraph
 * containing "Select Tool" and replaces it.
 *
 * Dropdown chips in Google Docs store their current value as readable
 * text. replaceText() can update this text; if the chip's selectedIndex
 * is managed separately and reverts on click, set the value manually
 * or request a v2 update using the advanced Docs REST API (batchUpdate
 * with a replaceNamedRangeContent or chip-update operation).
 */
function setNearbyTool(body, headingPara, toolValue) {
  const paras = body.getParagraphs();
  const H2    = DocumentApp.ParagraphHeading.HEADING2;
  const H3    = DocumentApp.ParagraphHeading.HEADING3;
  const H4    = DocumentApp.ParagraphHeading.HEADING4;

  // Find index of the heading para
  let idx = -1;
  for (let i = 0; i < paras.length; i++) {
    if (paras[i] === headingPara) { idx = i; break; }
  }
  if (idx < 0) return false;

  const limit = Math.min(idx + 7, paras.length);
  for (let i = idx + 1; i < limit; i++) {
    const para = paras[i];
    const h    = para.getHeading();
    if (h === H2 || h === H3 || h === H4) break; // Entered next section

    if (para.getText().includes('Select Tool')) {
      try {
        para.replaceText('Select Tool', toolValue);
        Logger.log(`  Tool → ${toolValue}`);
        return true;
      } catch (e) {
        Logger.log(`  Could not set tool dropdown (${e.message}) — set manually`);
        return false;
      }
    }
  }
  return false;
}


// ── REMOVE EXTRA SLOT ────────────────────────────────────────────────
/**
 * Removes an activity slot: the HEADING4 line plus all following
 * body paragraphs up to (but not including) the next heading.
 */
function removeSlot(body, headingPara) {
  const paras = body.getParagraphs();
  const H2    = DocumentApp.ParagraphHeading.HEADING2;
  const H3    = DocumentApp.ParagraphHeading.HEADING3;
  const H4    = DocumentApp.ParagraphHeading.HEADING4;

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

  // Remove bottom-up to prevent index shifting
  for (let i = toRemove.length - 1; i >= 0; i--) {
    try { toRemove[i].removeFromParent(); }
    catch (e) { Logger.log('Could not remove: ' + e.message); }
  }

  Logger.log(`Deleted slot: ${headingPara.getText()}`);
}
