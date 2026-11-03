// ============================================================
// Blueprint Tools — Spikes.gs   (THROWAWAY: delete this file after the spikes)
// Last updated on 2026-11-02 at 19:05 MST
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
 * Spike C (catalog endpoint). Fetches the two catalog home pages and a few
 * guesses at a data endpoint, and logs status and the start of each response.
 * The catalogs are hash-routed single-page apps, so the home HTML is probably
 * just a shell. THE REAL ANSWER is in the browser: open a course page in the
 * catalog, open DevTools > Network > Fetch/XHR, reload, and look for the request
 * whose response contains the course description. Copy that request's URL and
 * paste it to Claude Code. This function only checks the cheap guesses.
 */
function spikeCatalogSpike10() {
  var urls = [
    'https://www.boisestate.edu/registrar-catalog/',
    'https://www.boisestate.edu/graduatecatalog/',
    'https://boisestate.kuali.co/api/v1/catalog/public/catalogs',
    'https://boisestate.kuali.co/api/v1/catalog/courses'
  ];
  urls.forEach(function (u) {
    try {
      var r = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true });
      var body = r.getContentText();
      Logger.log('%s -> %s, %s chars, type %s\n%s', u, r.getResponseCode(), body.length, r.getHeaders()['Content-Type'], body.slice(0, 400));
      var scripts = body.match(/<script[^>]+src="[^"]+"/g);
      if (scripts) Logger.log('  script tags: %s', scripts.slice(0, 6).join(' ; '));
    } catch (e) { Logger.log('%s -> ERROR %s', u, e.message); }
  });
}
