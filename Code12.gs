// ============================================================
// Blueprint Tools — Code12.gs
// Draft Settings Tables
// Adds draft tabs to the project's settings table (a Google Sheet) by copying
// the template tabs (Assignment, Discussion, Classic Quiz, New Quiz) once per
// kind of activity in the Development tab. Everything after that (filling the
// tabs in) stays manual.
// ------------------------------------------------------------
// Last updated on 2026-11-03 at 09:10 MST
// ------------------------------------------------------------
//
// Spec: docs/dev-handoff/HANDOFF.md section 3.3. Mockups: DraftTabs.png,
// A1Title.png. UI: Sidebar12.html. Runs BEFORE Deploy Settings Tables
// (Code11.gs), which then links each activity to its tab.
//
// CONTRACT.
//  - Preview first: previewDraft12 only reads. Nothing is created until the
//    user confirms the list of tab names and presses "Create".
//  - A whole template tab is copied (Sheet.copyTo), never rebuilt, so the
//    conditional formatting that depends on fixed cell positions comes along.
//  - NO CELL IS EDITED. Cell A1 stays exactly as the template has it (Greg chose
//    option A in A1Title.png). Only the new tab's NAME is set.
//  - Existing tabs are never changed, renamed, moved or deleted. A proposed
//    name that matches an existing tab is skipped.
//  - The "Check when settings are finalized:" box is never ticked or unticked.
//    If a template tab arrives with it ticked, the result warns about it.
//
// GAS FLAT NAMESPACE. Every name in this file ends in 12 (or 12_ for private
// helpers). Reuses, does NOT redefine:
//   scanActivities11_, loadSettingsTable11, getDeployInit11   (Code11.gs)
//   normBase9_                                                (Code9.gs)
//
// SCOPES. Needs the Sheets scope "spreadsheets" (read and write), approved by
// Greg on 2026-11-03 and already in appsscript.json.
//
// FLAGS. Open questions are marked "DEFAULT:" in a comment.
// ============================================================


// ── CONSTANTS ────────────────────────────────────────────────

var MAX_TAB_NAME_12 = 100;          // Greg: names stop at 100 characters
var TEMPLATE_TAB_RE_12 = /template/i;
var FINALIZE_RE_12 = /settings are finalized/i;

// Which template tab a Canvas tool is copied from: a regex on the template
// tab's name. DEFAULT: tabs are found by name, since the template names are not
// stored anywhere else. "Page" has no settings tab.
var TOOL_TEMPLATE_RE_12 = {
  'Assignment':              /assignment/i,
  'Assignment (Not Graded)': /assignment/i,
  'Discussion':              /discussion/i,
  'Quiz (Classic)':          /classic/i,
  'Quiz (New)':              /\bnew\b/i
};

// "Topic 3", "Week 2" ... labels that make otherwise identical activities look
// different. Removed when grouping and when suggesting a name.
var TOPIC_LABEL_RE_12 = /\b(?:topic|week|module|unit|part)\s*\d+\b/gi;


// ── MENU ENTRY ───────────────────────────────────────────────

function showDraftSidebar12() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar12')
    .setTitle('Draft Settings Tables')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}


// ── SIDEBAR API ──────────────────────────────────────────────
// The sidebar opens with getDeployInit11 (finds the settings table through the
// Dashboard link) and loadSettingsTable11 ("Choose a different file"), both in
// Code11.gs.

/**
 * Read-only. Proposes the tabs to create.
 * @param {{tableId: string, mode: 'kind'|'tool'|'activity'}} state
 * @returns {{table, templates, existing, proposals, skipped, warnings}}
 */
function previewDraft12(state) {
  var doc = DocumentApp.getActiveDocument();
  var table = loadSettingsTable11(state.tableId);
  var acts = scanActivities11_(doc);
  var mode = state.mode === 'tool' || state.mode === 'activity' ? state.mode : 'kind';

  var templates = table.tabs.filter(function (t) { return TEMPLATE_TAB_RE_12.test(t.name); });
  var existingNames = table.tabs.map(function (t) { return t.name; });
  var warnings = [];

  // Which template serves each tool.
  var tmplFor = {};
  Object.keys(TOOL_TEMPLATE_RE_12).forEach(function (tool) {
    for (var i = 0; i < templates.length; i++) {
      if (TOOL_TEMPLATE_RE_12[tool].test(templates[i].name) && !(tool === 'Quiz (New)' && /classic/i.test(templates[i].name))) { tmplFor[tool] = templates[i]; break; }
    }
  });

  // Group activities.
  var groups = [], byKey = {}, skipped = { page: 0, unknown: 0 };
  var order = 0;
  acts.forEach(function (m) {
    m.activities.forEach(function (a) {
      if (a.isPage) { skipped.page++; return; }
      if (!TOOL_TEMPLATE_RE_12[a.tool]) { skipped.unknown++; return; }
      var key = mode === 'tool' ? a.tool
              : mode === 'activity' ? a.id
              : a.tool + '|' + normBase9_(stripTopic12_(a.title));
      var g = byKey[key];
      if (!g) {
        g = byKey[key] = { key: key, tool: a.tool, name: mode === 'tool' ? a.tool : suggestName12_(a.title, mode), activities: [], order: order++ };
        groups.push(g);
      }
      g.activities.push(m.label + ' › ' + a.title);
    });
  });

  // Unique, valid names. A name that matches an existing tab is flagged and
  // skipped; two proposals with the same name get a number on the later one.
  var existingSet = {}, used = {};
  existingNames.forEach(function (n) { existingSet[n.toLowerCase()] = true; used[n.toLowerCase()] = true; });
  var proposals = groups.map(function (g) {
    var name = g.name.slice(0, MAX_TAB_NAME_12), low = name.toLowerCase(), exists = !!existingSet[low];
    if (!exists && used[low]) {
      var n = 2, cand;
      do { cand = (name.slice(0, MAX_TAB_NAME_12 - 4) + ' ' + n).trim(); n++; } while (used[cand.toLowerCase()]);
      name = cand; low = name.toLowerCase();
    }
    if (!exists) used[low] = true;
    var tm = tmplFor[g.tool] || null;
    return {
      key: g.key, name: name, tool: g.tool, count: g.activities.length, activities: g.activities,
      templateGid: tm ? tm.gid : '', templateName: tm ? tm.name : '',
      exists: exists,
      noTemplate: !tm
    };
  });

  var missing = [];
  proposals.forEach(function (p) { if (p.noTemplate && missing.indexOf(p.tool) === -1) missing.push(p.tool); });
  if (missing.length) warnings.push('No template tab found for: ' + missing.join(', ') + '. Those tabs cannot be created. Template tabs are found by name (the tab name must contain “Template” and the tool’s name).');
  if (!templates.length) warnings.push('This settings table has no tab with “Template” in its name, so there is nothing to copy.');

  // A template that arrives with the finalize box already ticked: warn, never untick.
  var ticked = [];
  var usedTemplates = {};
  proposals.forEach(function (p) { if (p.templateGid) usedTemplates[p.templateGid] = p.templateName; });
  Object.keys(usedTemplates).forEach(function (gid) {
    if (isFinalizedTicked12_(sheetByGid12_(SpreadsheetApp.openById(table.id), gid))) ticked.push(usedTemplates[gid]);
  });
  if (ticked.length) warnings.push('The “Check when settings are finalized:” box is already ticked on ' + ticked.join(', ') + '. New copies will arrive ticked. This tool never changes that box; untick it on each new tab by hand.');

  return {
    table: { id: table.id, name: table.name, url: table.url },
    templates: templates.map(function (t) { return { gid: t.gid, name: t.name }; }),
    existing: existingNames,
    proposals: proposals,
    skipped: skipped,
    warnings: warnings,
    activityCount: acts.reduce(function (s, m) { return s + m.activities.length; }, 0)
  };
}

/**
 * Creates the tabs the user confirmed. The sidebar sends the final list; the
 * server re-checks every name and template before copying anything.
 * @param {{tableId: string, tabs: Array<{name: string, templateGid: string}>}} state
 * @returns {{created: Array<{name, gid, url}>, skipped: string[], warnings: string[]}}
 */
function applyDraft12(state) {
  var table = loadSettingsTable11(state.tableId);
  var ss = SpreadsheetApp.openById(table.id);
  var taken = {};
  table.tabs.forEach(function (t) { taken[t.name.toLowerCase()] = true; });
  var res = { created: [], skipped: [], warnings: [] };

  (state.tabs || []).forEach(function (req) {
    var name = String(req.name || '').trim();
    if (!name) { res.skipped.push('A tab with no name was skipped.'); return; }
    if (name.length > MAX_TAB_NAME_12) { res.skipped.push('“' + name.slice(0, 30) + '…” is longer than ' + MAX_TAB_NAME_12 + ' characters.'); return; }
    if (taken[name.toLowerCase()]) { res.skipped.push('“' + name + '” already exists (or appears twice), so it was not created.'); return; }
    var tmpl = sheetByGid12_(ss, req.templateGid);
    if (!tmpl || !TEMPLATE_TAB_RE_12.test(tmpl.getName())) { res.skipped.push('“' + name + '”: the template tab could not be found.'); return; }
    try {
      var copy = tmpl.copyTo(ss);        // whole tab, formatting and all; cells untouched
      copy.setName(name);
      if (copy.isSheetHidden()) copy.showSheet();
      taken[name.toLowerCase()] = true;
      res.created.push({ name: name, gid: String(copy.getSheetId()), url: table.url + '#gid=' + copy.getSheetId() });
      if (isFinalizedTicked12_(copy)) res.warnings.push('“' + name + '” arrived with “Check when settings are finalized:” ticked (it is ticked on the template). It was left as it is.');
    } catch (e) {
      res.skipped.push('“' + name + '”: ' + (e && e.message || e));
    }
  });
  return res;
}


// ── HELPERS ──────────────────────────────────────────────────

function sheetByGid12_(ss, gid) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) if (String(sheets[i].getSheetId()) === String(gid)) return sheets[i];
  return null;
}

/** "Discussion: Topic 1" -> "Discussion: " with the topic label and stray numbers removed. */
function stripTopic12_(title) {
  return String(title || '').replace(TOPIC_LABEL_RE_12, ' ').replace(/\b\d+\b/g, ' ').replace(/\s+/g, ' ').trim();
}

/** A readable tab name from an activity title: topic labels and numbers dropped. */
function suggestName12_(title, mode) {
  if (mode === 'activity') return String(title).trim() || 'Activity';
  var s = stripTopic12_(title).replace(/\s*:\s*$/, '').replace(/^\s*:\s*/, '').replace(/\s+:/g, ':').trim();
  return s || String(title).trim() || 'Tab';
}

/**
 * Is the "Check when settings are finalized:" box ticked on this tab? Finds the
 * row by its text, then looks for a checkbox value (true) in the same row.
 * READ ONLY. Returns false when the row or box is not found.
 */
function isFinalizedTicked12_(sheet) {
  if (!sheet) return false;
  var range = sheet.getDataRange(), vals = range.getValues();
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      if (typeof vals[r][c] === 'string' && FINALIZE_RE_12.test(vals[r][c])) {
        // The box may sit in the same cell's row to the right, or on the next row.
        for (var k = c; k < vals[r].length; k++) if (vals[r][k] === true) return true;
        if (r + 1 < vals.length) for (var k2 = 0; k2 < vals[r + 1].length; k2++) if (vals[r + 1][k2] === true && k2 >= c) return true;
        return false;
      }
    }
  }
  return false;
}
