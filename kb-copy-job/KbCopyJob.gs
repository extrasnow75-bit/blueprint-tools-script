// ============================================================
// KB Copy Job — KbCopyJob.gs   (PROTOTYPE)
// Copies eCampus Knowledge Base (Confluence) pages into Google Docs so
// the Blueprint Tools' "Set Up Gemini Notebook" can list them as
// notebook sources. Gemini Notebook can't read pages that need a
// Confluence sign-in; it can read Google Docs.
// ------------------------------------------------------------
// Last updated on 2026-11-22 at 11:30 MST
// ------------------------------------------------------------
//
// THIS IS NOT PART OF THE BLUEPRINT TOOLS SCRIPT.
//   It lives in its own Apps Script project, bound to a private Google
//   Sheet that only its owner (Greg) can edit. See README.md beside it.
//   The 14 IDCs never run it and never see the Confluence token; they
//   only need VIEW access to the output folder.
//
// WHERE THE TOKEN LIVES
//   In the owner's User Properties for this script (Script menu →
//   "Save Confluence sign-in"). Never in code, never in the Sheet,
//   never in logs. The public GitHub repo holds no secret.
//
// WHICH PAGES
//   1. Every Confluence page linked in the "Gemini Notebook Reference
//      List" Sheet (REFERENCE_SHEET_ID), and
//   2. any extra page links pasted into column A of this Sheet's
//      "Pages to copy" tab (handy for testing).
//
// HOW EACH PAGE IS COPIED
//   One Google Doc per page in OUTPUT_FOLDER_ID, named "KB – <title>".
//   The Doc keeps the SAME file (same link) on every refresh, so a
//   Gemini Notebook source made from it can simply be re-synced. A page
//   is only rewritten when its Confluence version number changes.
//   Each Doc's Drive description holds "confluence-page-id: <id>",
//   which is how Set Up Gemini Notebook (Code17.gs) finds it.
//   Images are not copied (Confluence images need a sign-in); each one
//   becomes "[Image – see the original page]".
//
// NEEDS
//   • Advanced service: Drive API (v3), identifier "Drive".
//   • Scopes in appsscript.json (beside this file).
// ============================================================

var CONFLUENCE_SITE  = 'https://boisestateecampus.atlassian.net';
// "Gemini Notebook Reference List" Sheet: every Knowledge Base link in it
// gets copied. Same Sheet Set Up Gemini Notebook reads (REFERENCE_SHEET_ID_17).
var REFERENCE_SHEET_ID = '1991TCf1UoWGRyl_3YdDR9cBMrdS4eouRhr8tnwpjGz4';
// "Gemini Notebook Project Work Station Stuff"
var OUTPUT_FOLDER_ID = '15_H0_97_6ZqU-Pi8nW79PSf2_ARNtxzz';

var PAGES_TAB  = 'Pages to copy';
var COPIES_TAB = 'Copies';
var COPIES_HEADER = ['Page ID', 'Title', 'Confluence page', 'Google Doc',
                     'Version copied', 'Last checked', 'Result'];
var NIGHTLY_HOUR = 2;           // 2-3 a.m. in the script's time zone
var PROP_EMAIL = 'KB_CONFLUENCE_EMAIL';
var PROP_TOKEN = 'KB_CONFLUENCE_TOKEN';
var PROP_BASE  = 'KB_CONFLUENCE_BASE';   // which address worked last time


// ── MENU ─────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📚 KB Copy Job')
    .addItem('1. Save Confluence sign-in',   'menuSaveSignIn')
    .addItem('2. Test connection',            'menuTestConnection')
    .addItem('3. Copy pages now',             'menuCopyNow')
    .addSeparator()
    .addItem('Turn ON nightly copy',          'menuNightlyOn')
    .addItem('Turn OFF nightly copy',         'menuNightlyOff')
    .addSeparator()
    .addItem('Forget Confluence sign-in',     'menuForgetSignIn')
    .addToUi();
  ensureTabs_();
}

function menuSaveSignIn() {
  var ui = SpreadsheetApp.getUi();
  var e = ui.prompt('Confluence sign-in (1 of 2)',
    'Your Atlassian email address (the one you sign in to Confluence with):',
    ui.ButtonSet.OK_CANCEL);
  if (e.getSelectedButton() !== ui.Button.OK) return;
  var email = e.getResponseText().trim();
  var t = ui.prompt('Confluence sign-in (2 of 2)',
    'Paste the API token. It is stored only in your private script settings ' +
    'and is never shown again.', ui.ButtonSet.OK_CANCEL);
  if (t.getSelectedButton() !== ui.Button.OK) return;
  var token = t.getResponseText().trim();
  if (!email || !token) { ui.alert('Nothing was saved: both boxes are needed.'); return; }
  var p = PropertiesService.getUserProperties();
  p.setProperty(PROP_EMAIL, email);
  p.setProperty(PROP_TOKEN, token);
  p.deleteProperty(PROP_BASE);
  ui.alert('Saved. Next: 📚 KB Copy Job → 2. Test connection.');
}

function menuForgetSignIn() {
  var p = PropertiesService.getUserProperties();
  [PROP_EMAIL, PROP_TOKEN, PROP_BASE].forEach(function (k) { p.deleteProperty(k); });
  SpreadsheetApp.getUi().alert('The Confluence sign-in was removed from this script. ' +
    'If you think the token leaked, also delete it at id.atlassian.com → Security → API tokens.');
}

function menuTestConnection() {
  var ui = SpreadsheetApp.getUi();
  var pages = collectPages_();
  if (!pages.length) {
    ui.alert('No Confluence page links found yet. Paste one into column A of the "' +
      PAGES_TAB + '" tab (or add some to the Gemini Notebook Reference List), then test again.');
    return;
  }
  var r = fetchPage_(pages[0].id);
  ui.alert(r.ok
    ? 'Connected. Confluence returned the page "' + r.page.title + '" (version ' +
      r.page.version + ').\n\nNext: 📚 KB Copy Job → 3. Copy pages now.'
    : 'Not connected.\n\n' + r.message);
}

function menuCopyNow() {
  var res = copyAllPages();
  SpreadsheetApp.getUi().alert(res.summary + '\n\nDetails are in the "' + COPIES_TAB + '" tab.');
}

function menuNightlyOn() {
  menuNightlyOff_(true);
  ScriptApp.newTrigger('copyAllPages').timeBased().everyDays(1).atHour(NIGHTLY_HOUR).create();
  SpreadsheetApp.getUi().alert('Nightly copy is ON. It runs between ' + NIGHTLY_HOUR + ' and ' +
    (NIGHTLY_HOUR + 1) + ' a.m. under your account, even when this Sheet is closed.');
}

function menuNightlyOff() { menuNightlyOff_(false); }

function menuNightlyOff_(silent) {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'copyAllPages') { ScriptApp.deleteTrigger(tr); n++; }
  });
  if (!silent) SpreadsheetApp.getUi().alert(n ? 'Nightly copy is OFF.' : 'Nightly copy was already off.');
}


// ── MAIN JOB (menu and nightly trigger) ──────────────────────

/** Copies every listed page that is new or changed. No UI calls here,
 *  because the nightly trigger runs it with no one watching. */
function copyAllPages() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return { summary: 'A copy is already running. Try again in a minute.' };
  try {
    ensureTabs_();
    var pages = collectPages_();
    var log = readCopies_();
    var folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    var counts = { copied: 0, unchanged: 0, failed: 0 };
    var now = new Date();

    pages.forEach(function (pg) {
      var row = log.byId[pg.id] || { pageId: pg.id };
      row.checked = now;
      var r = fetchPage_(pg.id);
      if (!r.ok) {
        row.result = 'Failed: ' + r.message;
        if (!row.confluenceUrl) row.confluenceUrl = pg.url;
        counts.failed++;
        log.put(row);
        return;
      }
      var page = r.page;
      row.title = page.title;
      row.confluenceUrl = page.webUrl || pg.url;
      var docOk = row.docId && docExists_(row.docId);
      if (docOk && String(row.version) === String(page.version)) {
        row.result = 'Unchanged';
        counts.unchanged++;
        log.put(row);
        return;
      }
      try {
        row.docId = writeDoc_(docOk ? row.docId : null, page, folder);
        row.version = page.version;
        row.result = docOk ? 'Updated' : 'Created';
        counts.copied++;
      } catch (e) {
        row.result = 'Failed to write the Google Doc: ' + (e && e.message ? e.message : e);
        counts.failed++;
      }
      log.put(row);
    });

    log.save();
    return { summary: pages.length + ' page(s) checked: ' + counts.copied + ' copied or updated, ' +
      counts.unchanged + ' unchanged, ' + counts.failed + ' failed.' };
  } finally {
    lock.releaseLock();
  }
}


// ── WHICH PAGES ──────────────────────────────────────────────

/** Confluence page links from the reference list Sheet and the Pages tab,
 *  one entry per page ID. */
function collectPages_() {
  var urls = [];
  try {
    var ref = SpreadsheetApp.openById(REFERENCE_SHEET_ID);
    ref.getSheets().forEach(function (tab) {
      var range = tab.getDataRange();
      var vals = range.getDisplayValues(), rich = range.getRichTextValues();
      vals.forEach(function (row, i) {
        row.forEach(function (v, j) {
          var rt = rich[i][j];
          if (rt) {
            if (rt.getLinkUrl()) urls.push(rt.getLinkUrl());
            rt.getRuns().forEach(function (run) { if (run.getLinkUrl()) urls.push(run.getLinkUrl()); });
          }
          if (/^https?:/i.test(v)) urls.push(v);
        });
      });
    });
  } catch (e) { /* list unreadable: carry on with this Sheet's own tab */ }

  var sh = SpreadsheetApp.getActive().getSheetByName(PAGES_TAB);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getRichTextValues().forEach(function (r) {
      var cell = r[0];
      var link = cell.getLinkUrl();
      urls.push(link || cell.getText());
    });
  }

  var seen = {}, out = [];
  urls.forEach(function (u) {
    var id = confluencePageId(u);
    if (id && !seen[id]) { seen[id] = true; out.push({ id: id, url: String(u).trim() }); }
  });
  return out;
}


/** PURE. The numeric page ID in a Confluence page link on our site, or null. */
function confluencePageId(url) {
  var s = String(url || '');
  if (!/^https?:\/\/[^\/]*atlassian\.net\/wiki\//i.test(s)) return null;
  var m = s.match(/\/pages\/(\d+)/) || s.match(/[?&]pageId=(\d+)/);
  return m ? m[1] : null;
}


// ── CONFLUENCE ───────────────────────────────────────────────

/**
 * Reads one page. Returns {ok:true, page:{id,title,version,html,webUrl}}
 * or {ok:false, message} in plain language. Never includes the token.
 */
function fetchPage_(pageId) {
  var p = PropertiesService.getUserProperties();
  var email = p.getProperty(PROP_EMAIL), token = p.getProperty(PROP_TOKEN);
  if (!email || !token) {
    return { ok: false, message: 'No Confluence sign-in is saved. Use 📚 KB Copy Job → 1. Save Confluence sign-in.' };
  }
  var auth = 'Basic ' + Utilities.base64Encode(email + ':' + token);

  // Scoped tokens only work through Atlassian's api.atlassian.com
  // address; classic tokens through the site address. Try the one that
  // worked last time first, then the other.
  var bases = candidateBases_();
  var last = null;
  for (var b = 0; b < bases.length; b++) {
    var formats = ['view', 'export_view', 'storage'];
    for (var f = 0; f < formats.length; f++) {
      var url = bases[b] + '/wiki/api/v2/pages/' + encodeURIComponent(pageId) +
                '?body-format=' + formats[f];
      var resp;
      try {
        resp = UrlFetchApp.fetch(url, { headers: { Authorization: auth, Accept: 'application/json' },
                                       muteHttpExceptions: true });
      } catch (e) {
        last = { code: 0 };
        break;
      }
      var code = resp.getResponseCode();
      if (code === 200) {
        p.setProperty(PROP_BASE, bases[b]);
        var j = JSON.parse(resp.getContentText());
        var body = j.body && j.body[formats[f]] ? j.body[formats[f]].value : '';
        return { ok: true, page: {
          id: String(j.id || pageId),
          title: j.title || ('Page ' + pageId),
          version: j.version ? j.version.number : '',
          html: body || '',
          webUrl: j._links && j._links.webui ? CONFLUENCE_SITE + '/wiki' + j._links.webui : ''
        } };
      }
      last = { code: code };
      if (code === 400) continue;      // that body format isn't offered; try the next
      break;                           // anything else: try the other address
    }
  }
  return { ok: false, message: explainStatus_(last ? last.code : 0, pageId) };
}

function candidateBases_() {
  var list = [];
  var saved = PropertiesService.getUserProperties().getProperty(PROP_BASE);
  if (saved) list.push(saved);
  var cloud = cloudId_();
  if (cloud) list.push('https://api.atlassian.com/ex/confluence/' + cloud);
  list.push(CONFLUENCE_SITE);
  return list.filter(function (x, i) { return list.indexOf(x) === i; });
}

/** The site's cloud ID (public; needed for scoped tokens). Cached 6 h. */
function cloudId_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('kb_cloud_id');
  if (hit) return hit;
  try {
    var r = UrlFetchApp.fetch(CONFLUENCE_SITE + '/_edge/tenant_info', { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      var id = JSON.parse(r.getContentText()).cloudId;
      if (id) { cache.put('kb_cloud_id', id, 6 * 60 * 60); return id; }
    }
  } catch (e) { /* fall back to the site address */ }
  return null;
}

/** PURE. Plain-language reason for a failed request. */
function explainStatus_(code, pageId) {
  switch (code) {
    case 401: return 'Confluence did not accept the email and token. Check that the email is ' +
      'your Atlassian sign-in and that the token hasn\'t expired, then save the sign-in again.';
    case 403: return 'The token works but is not allowed to read page ' + pageId + '. If it is a ' +
      'scoped token, check it includes the "read:page:confluence" permission; otherwise check ' +
      'you can open this page in Confluence.';
    case 404: return 'Page ' + pageId + ' was not found. It may have been deleted or moved, or ' +
      'you may not have access to it.';
    case 429: return 'Confluence asked us to slow down. The next run will try again.';
    case 0:   return 'Couldn\'t reach Confluence (network problem). The next run will try again.';
    default:  return 'Confluence answered with an unexpected error (' + code + ').';
  }
}


// ── GOOGLE DOC ───────────────────────────────────────────────

/** Creates or overwrites the Doc for one page; returns its file ID. */
function writeDoc_(docId, page, folder) {
  var html = buildDocHtml(page, CONFLUENCE_SITE,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a'));
  var blob = Utilities.newBlob(html, 'text/html', 'page.html');
  var meta = {
    name: 'KB – ' + page.title,
    description: 'confluence-page-id: ' + page.id + '\nCopied automatically from ' +
      (page.webUrl || CONFLUENCE_SITE) + '. Edit the page in Confluence, not here: ' +
      'changes made in this Doc are overwritten.'
  };
  if (docId) {
    Drive.Files.update(meta, docId, blob);
    return docId;
  }
  meta.parents = [folder.getId()];
  meta.mimeType = 'application/vnd.google-apps.document';   // convert the HTML to a Doc
  return Drive.Files.create(meta, blob).id;
}

function docExists_(id) {
  try { return !DriveApp.getFileById(id).isTrashed(); } catch (e) { return false; }
}

/** PURE. The HTML that becomes the Google Doc. */
function buildDocHtml(page, site, stamp) {
  var body = String(page.html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<img\b[^>]*>/gi, '<em>[Image – see the original page]</em>')
    // Confluence uses site-relative links; make them work from Drive.
    .replace(/(href=["'])\/(?!\/)/gi, '$1' + site + '/');
  var link = page.webUrl || site;
  return '<html><head><meta charset="utf-8"></head><body>' +
    '<p style="color:#555;font-size:9pt">Copied from the eCampus Knowledge Base: ' +
    '<a href="' + escHtml_(link) + '">' + escHtml_(link) + '</a><br>' +
    'Page version ' + escHtml_(page.version) + ', copied ' + escHtml_(stamp) + '. ' +
    'Edit the original page in Confluence; changes made here are overwritten.</p>' +
    '<h1>' + escHtml_(page.title) + '</h1>' + body + '</body></html>';
}

function escHtml_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}


// ── THE SHEET ────────────────────────────────────────────────

function ensureTabs_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(PAGES_TAB)) {
    var p = ss.insertSheet(PAGES_TAB);
    p.getRange(1, 1, 1, 2).setValues([['Confluence page link', 'Note (optional)']]).setFontWeight('bold');
    p.setColumnWidth(1, 480);
    p.setFrozenRows(1);
  }
  if (!ss.getSheetByName(COPIES_TAB)) {
    var c = ss.insertSheet(COPIES_TAB);
    c.getRange(1, 1, 1, COPIES_HEADER.length).setValues([COPIES_HEADER]).setFontWeight('bold');
    c.setFrozenRows(1);
  }
}

/** The Copies tab as an editable map keyed by page ID. */
function readCopies_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(COPIES_TAB);
  var byId = {}, order = [];
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, COPIES_HEADER.length).getValues().forEach(function (v) {
      if (!v[0]) return;
      var docUrl = String(v[3] || '');
      var m = docUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      var row = { pageId: String(v[0]), title: v[1], confluenceUrl: v[2],
                  docId: m ? m[1] : '', version: v[4], checked: v[5], result: v[6] };
      byId[row.pageId] = row; order.push(row.pageId);
    });
  }
  return {
    byId: byId,
    put: function (row) { if (!byId[row.pageId]) order.push(row.pageId); byId[row.pageId] = row; },
    save: function () {
      var rows = order.map(function (id) {
        var r = byId[id];
        return [r.pageId, r.title || '', r.confluenceUrl || '',
                r.docId ? 'https://docs.google.com/document/d/' + r.docId + '/edit' : '',
                r.version || '', r.checked || '', r.result || ''];
      });
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, COPIES_HEADER.length).clearContent();
      if (rows.length) sh.getRange(2, 1, rows.length, COPIES_HEADER.length).setValues(rows);
    }
  };
}
