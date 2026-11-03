// ============================================================
// Blueprint Tools — Spikes.gs   (THROWAWAY: delete this file after the spikes)
// Last updated on 2026-11-03 at 01:21 MST
// ------------------------------------------------------------
// Read-only checks for Link Project Documents. Nothing here writes to any
// document or file. Run each function from the Apps Script editor (pick it in
// the function dropdown, press Run), then open View > Logs / Execution log and
// copy what it printed back to Claude Code.
// ============================================================

/**
 * Spike A. Does findText accept a "^" anchor, and does it find the Dashboard
 * labels? Run this from the blueprint's script project.
 * Expected: each label prints a hit with its current link (or null).
 */
function spikeLabelsSpike10() {
  var doc = DocumentApp.getActiveDocument();
  var tabs = collectTabs(doc), body = tabs[0].body;
  tabs.forEach(function (t) { if (/dashboard/i.test(t.title)) body = t.body; });
  ['^\\s*Google Folder', '^\\s*Schedule of Deliverables', '^\\s*Settings Tables?', '^\\s*Syllabus',
   '^\\s*Instructor Guide', '^\\s*Course Suggestions?', '^\\s*End[\\s-]*of[\\s-]*Project Report'].forEach(function (re) {
    var r = body.findText(re), n = 0, first = null;
    while (r) { n++; if (!first) { var el = r.getElement().asText(); first = el.getText().slice(0, 40) + ' | link=' + el.getLinkUrl(r.getStartOffset()); } r = body.findText(re, r); }
    Logger.log('%s -> %s hit(s). First: %s', re, n, first);
  });
  // Instructor Resources heading levels.
  collectTabs(doc).forEach(function (t) {
    if (!/development/i.test(t.title)) return;
    var b = t.body, n = b.getNumChildren(), shown = 0;
    for (var i = 0; i < n && shown < 12; i++) {
      var c = b.getChild(i);
      if (c.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      var p = c.asParagraph();
      if (/Instructor Resources|Instructor Guide|Course Suggestions/i.test(p.getText())) {
        Logger.log('child %s | %s | %s', i, p.getHeading(), p.getText().slice(0, 70)); shown++;
      }
    }
  });
}

/**
 * Spike B. Can DriveApp list this blueprint's folder? Prints folder name and
 * each file's name and type. This is what Link Project Documents does on open.
 */
function spikeFolderSpike10() {
  var id = DocumentApp.getActiveDocument().getId();
  var parents = DriveApp.getFileById(id).getParents();
  if (!parents.hasNext()) { Logger.log('No parent folder.'); return; }
  var folder = parents.next();
  Logger.log('Folder: %s (%s)', folder.getName(), folder.getId());
  var it = folder.getFiles();
  while (it.hasNext()) { var f = it.next(); Logger.log('%s | %s | %s', f.getName(), f.getMimeType(), f.getId()); }
}

/**
 * Spike C (catalog). The catalogs run on Kuali (found in Greg's Network tab).
 * This checks the two things the lookup depends on:
 *   1. Does /api/v1/catalog/public/catalogs list the catalogs, and what are
 *      their names and IDs? (Needed to find the graduate catalog's ID and to
 *      confirm "newest undergraduate / graduate" discovery.)
 *   2. Does a search for a course return its description? Tries CS 321
 *      (undergraduate) and APPAI 505 (graduate) in every catalog whose name
 *      mentions "graduate".
 * Copy everything it logs back to Claude Code.
 */
function spikeCatalogSpike10() {
  var base = 'https://boisestate.kuali.co/api/v1/catalog/';
  var r = UrlFetchApp.fetch(base + 'public/catalogs', { muteHttpExceptions: true, followRedirects: true });
  Logger.log('public/catalogs -> %s, %s chars', r.getResponseCode(), r.getContentText().length);
  var list = [];
  try { list = JSON.parse(r.getContentText()); } catch (e) { Logger.log('Not JSON: %s', r.getContentText().slice(0, 400)); return; }
  Logger.log('First entry keys: %s', list.length ? Object.keys(list[0]).join(', ') : '(empty)');
  list.forEach(function (c) {
    Logger.log('CATALOG | %s | id=%s | %s', c.name || c.title || c.fullName, c._id || c.id,
               ['activeDate', 'archived', 'status', 'startDate', 'endDate'].map(function (k) { return k + '=' + c[k]; }).join(' '));
  });
  list.forEach(function (c) {
    var name = String(c.name || c.title || '');
    if (!/graduate/i.test(name)) return;
    var q = /\bundergraduate\b/i.test(name) ? 'cs 321' : 'appai 505';
    var s = UrlFetchApp.fetch(base + 'search/' + (c._id || c.id) + '?q=' + encodeURIComponent(q) + '&limit=6', { muteHttpExceptions: true });
    var body = s.getContentText();
    Logger.log('SEARCH "%s" in %s -> %s: %s', q, name, s.getResponseCode(), body.slice(0, 500));
  });
}
