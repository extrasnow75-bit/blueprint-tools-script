// ============================================================
// Blueprint Tools — Spikes.gs   (THROWAWAY: delete this file after the spikes)
// Last updated on 2026-11-06 at 00:24 MST
// ------------------------------------------------------------
// Read-only checks for Setup Assistant. Nothing here writes to any
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
 * each file's name and type. This is what Setup Assistant does on open.
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


/**
 * Spike D (for Development Status Check, Code13.gs). READ-ONLY. Run it from a
 * blueprint's script project. It prints:
 *  1. every highlight colour used in the Development tab, with how Code13
 *     classifies it and how many text runs use it (so you can confirm green,
 *     purple and yellow are what you expect);
 *  2. the Home and Syllabus sections it finds;
 *  3. what it reads for Credits and Course Length on the Dashboard;
 *  4. the raw checkbox / status / link data of the first rows of the Creative
 *     Brief "Media Deliverables" table.
 * Copy the log back to Claude Code.
 */
function spikeStatusSpike13() {
  var doc = DocumentApp.getActiveDocument();
  var body = getDevelopmentTabBody(doc);
  var counts = {}, fcounts = {}, strikes = 0, runs = 0;
  (function walk(el) {
    var T = DocumentApp.ElementType, ty = el.getType(), t = null;
    if (ty === T.PARAGRAPH) t = el.asParagraph().editAsText(); else if (ty === T.LIST_ITEM) t = el.asListItem().editAsText();
    if (t) {
      var text = t.getText(), idx = t.getTextAttributeIndices();
      for (var i = 0; i < idx.length; i++) {
        var end = i + 1 < idx.length ? idx[i + 1] : text.length;
        if (!text.substring(idx[i], end).trim()) continue;
        runs++;
        var bg = t.getBackgroundColor(idx[i]); if (t.isStrikethrough(idx[i])) strikes++;
        if (bg) counts[bg] = (counts[bg] || 0) + 1;
        var fg = t.getForegroundColor(idx[i]); if (fg) fcounts[fg] = (fcounts[fg] || 0) + 1;
      }
      return;
    }
    if (ty === T.BODY_SECTION || ty === T.TABLE || ty === T.TABLE_ROW || ty === T.TABLE_CELL) for (var c = 0; c < el.getNumChildren(); c++) walk(el.getChild(c));
  })(body);
  Logger.log('Development tab: ' + runs + ' text runs, ' + strikes + ' struck through.');
  Object.keys(counts).forEach(function (k) { Logger.log('  highlight ' + k + ' x' + counts[k] + ' -> ' + stageOfColor13_(k)); });

  Object.keys(fcounts).forEach(function (k) { Logger.log('  font colour ' + k + ' x' + fcounts[k] + (isInstructionRed13_(k) ? ' -> treated as red builder instruction (ignored)' : '')); });

  var reg = homeSyllabusRegions13_(body);
  ['home', 'syllabus'].forEach(function (k) {
    Logger.log(k + ': ' + (reg[k] ? reg[k].elements.length + ' elements -> ' + JSON.stringify(describeRegion13_(reg[k], k === 'syllabus')) : 'NOT FOUND'));
    if (reg[k]) reg[k].elements.slice(0, 4).forEach(function (e) { var tx = e.getType() === DocumentApp.ElementType.TABLE ? e.asTable().getText() : (e.asParagraph ? e.asParagraph().getText() : ''); Logger.log('    ' + e.getType() + ': ' + String(tx).slice(0, 80)); });
  });

  var dash = dashboardTab10_(doc).body;
  Logger.log('Dashboard Credits -> "' + readDashField13_(dash, /^\s*credits?\b/i) + '"; Course Length -> "' + readDashField13_(dash, /^\s*course\s*length\b/i) + '"');

  var s = Docs.Documents.get(doc.getId(), { includeTabsContent: true }), tab = null;
  flattenDocsTabs_(s).forEach(function (t) { if (!tab && /^\s*creative brief\s*$/i.test((t.tabProperties || {}).title || '')) tab = t; });
  if (!tab) { Logger.log('No Creative Brief tab.'); return; }
  var content = tab.documentTab.body.content, table = findMediaTable13_(content);
  if (!table) { Logger.log('No table under Media Deliverables. Tables found in the Creative Brief tab:'); describeTables13_(content, 0, []).forEach(function (l) { Logger.log('  ' + l); }); return; }
  Logger.log('Header row: ' + JSON.stringify(table.tableRows[0].tableCells.map(cellText13_)));
  for (var r = 1; r < Math.min(table.tableRows.length, 4); r++) {
    table.tableRows[r].tableCells.forEach(function (cell, c) {
      Logger.log('row ' + r + ' col ' + c + ': ' + JSON.stringify(cell.content).slice(0, 700));
    });
  }
}
