// ============================================================
// Gemini Notebook API — personal-account test   (THROWAWAY TEST)
// Answers: can a script create a notebook and fill it with the kinds
// of sources Set Up Gemini Notebook would add?
// ------------------------------------------------------------
// Last updated on 2026-11-22 at 14:00 MST
// ------------------------------------------------------------
//
// RUN THIS IN A PERSONAL GOOGLE ACCOUNT ONLY. The API is in preview and
// limited to personal accounts. Use copies of test documents, never
// Boise State originals. Setup steps: README.md beside this file.
//
// WHAT THE API OFFERS (discovery document v1, revision 20261120):
//   notebooks: create, list, get, patch (rename)
//   sources:   create (text/Markdown or a web page URL), upload (a file,
//              e.g. a PDF), list, get
//   NOT offered: a Google Drive source type (the kind that auto-syncs),
//   and deleting notebooks or sources. Test notebooks must be deleted by
//   hand in Gemini Notebook.
//
// THE TESTS (runAllTests runs them in order and prints a summary):
//   1. List notebooks                → is setup/permission right?
//   2. Create a test notebook
//   3. Add a plain text source       → baseline
//   4. Add a public web page         → baseline
//   5. Add your test Google Doc as a PDF upload       (snapshot)
//   6. Add your test Google Doc as Markdown text      (snapshot)
//   7. Add your test Google Doc as a web page, three ways: its normal
//      /edit link, its plain-text export link, and (optional) its
//      "Publish to web" link. Set the Doc to "Anyone with the link" first.
//   8. Wait, then report each source's state: ACTIVE / FAILED / PROCESSING
// ============================================================

// ── SET THESE ────────────────────────────────────────────────
// Your Google Cloud project's ID (not its number), e.g. "notebook-api-test-123456".
// Sent as the x-goog-user-project header so requests are billed/counted to it.
var GCP_PROJECT_ID = 'PASTE_YOUR_PROJECT_ID_HERE';
// A Google Doc in THIS personal account (a copy of a test document).
var TEST_DOC_ID = 'PASTE_A_GOOGLE_DOC_ID_HERE';
// Any public web page.
var PUBLIC_URL = 'https://en.wikipedia.org/wiki/Instructional_design';
// Optional: for tests 7a–7c, set the test Doc's sharing to "Anyone with
// the link". For test 7c, also use File → Share → Publish to web and paste
// the published link here (it contains /pub). Leave blank to skip 7c.
var PUBLISHED_URL = '';
// ─────────────────────────────────────────────────────────────

var API = 'https://notebook.googleapis.com/v1/';
var UPLOAD_API = 'https://notebook.googleapis.com/upload/v1/';


/** Runs every test and logs a plain-language summary at the end. */
function runAllTests() {
  var results = [];
  function record(name, r) {
    results.push((r.ok ? '✅ ' : '❌ ') + name + (r.note ? ' — ' + r.note : ''));
    return r;
  }

  var list = record('1. List notebooks (setup check)', call_('get', API + 'notebooks?pageSize=5'));
  if (!list.ok) { finish_(results); return; }

  var nb = record('2. Create a test notebook', call_('post', API + 'notebooks', {
    displayName: 'Blueprint Tools API test ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  }));
  if (!nb.ok) { finish_(results); return; }
  var notebook = nb.json.name;                       // "notebooks/abc123"
  PropertiesService.getUserProperties().setProperty('TEST_NOTEBOOK', notebook);
  results.push('   Notebook: ' + notebook);

  record('3. Text source', call_('post', API + notebook + '/sources', {
    displayName: 'API test – plain text',
    mimeType: 'text/plain',
    textContent: { content: 'This is a test source created by the Blueprint Tools API test.' }
  }));

  record('4. Public web page source', call_('post', API + notebook + '/sources', {
    displayName: 'API test – public web page',
    webContent: { url: PUBLIC_URL }
  }));

  if (TEST_DOC_ID.indexOf('PASTE') === 0) {
    results.push('⏭️  5–7 skipped: set TEST_DOC_ID at the top of the script.');
  } else {
    var docName = DriveApp.getFileById(TEST_DOC_ID).getName();

    var pdf = DriveApp.getFileById(TEST_DOC_ID).getAs('application/pdf');
    pdf.setName(docName + '.pdf');
    record('5. Google Doc uploaded as a PDF', upload_(notebook, pdf, docName + ' (PDF upload)'));

    var md = exportDoc_(TEST_DOC_ID, 'text/markdown');
    record('6. Google Doc added as Markdown text', md.ok
      ? call_('post', API + notebook + '/sources', {
          displayName: docName + ' (Markdown text)',
          mimeType: 'text/markdown',
          textContent: { content: md.text }
        })
      : md);

    var shared = DriveApp.getFileById(TEST_DOC_ID).getSharingAccess() === DriveApp.Access.ANYONE_WITH_LINK;
    results.push('   (Test Doc sharing: ' + (shared ? 'anyone with the link' : 'NOT public') + ')');

    record('7a. Doc\'s normal /edit link as a web page', call_('post', API + notebook + '/sources', {
      displayName: docName + ' (edit link as web page)',
      webContent: { url: 'https://docs.google.com/document/d/' + TEST_DOC_ID + '/edit' }
    }));
    record('7b. Doc\'s plain-text export link as a web page', call_('post', API + notebook + '/sources', {
      displayName: docName + ' (text export link as web page)',
      webContent: { url: 'https://docs.google.com/document/d/' + TEST_DOC_ID + '/export?format=txt' }
    }));
    if (PUBLISHED_URL) {
      record('7c. Doc\'s "Publish to web" link as a web page', call_('post', API + notebook + '/sources', {
        displayName: docName + ' (published link as web page)',
        webContent: { url: PUBLISHED_URL }
      }));
    } else {
      results.push('⏭️  7c skipped: no PUBLISHED_URL set.');
    }
  }

  // Sources are processed in the background; give them a moment.
  Utilities.sleep(20 * 1000);
  results.push('8. Source states after 20 seconds:');
  results = results.concat(sourceStates_(notebook));
  finish_(results);
}

/** Run this a few minutes later to see whether PROCESSING sources finished. */
function checkSourcesAgain() {
  var notebook = PropertiesService.getUserProperties().getProperty('TEST_NOTEBOOK');
  if (!notebook) { Logger.log('Run runAllTests first.'); return; }
  finish_(['Source states for ' + notebook + ':'].concat(sourceStates_(notebook)));
}


// ── HELPERS ──────────────────────────────────────────────────

function authHeaders_() {
  var h = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  if (GCP_PROJECT_ID && GCP_PROJECT_ID.indexOf('PASTE') !== 0) h['x-goog-user-project'] = GCP_PROJECT_ID;
  return h;
}

/** One JSON API call. Returns {ok, json, note} and never throws. */
function call_(method, url, body) {
  var opts = {
    method: method,
    headers: authHeaders_(),
    muteHttpExceptions: true
  };
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  try {
    var r = UrlFetchApp.fetch(url, opts);
    return describe_(r);
  } catch (e) {
    return { ok: false, note: 'Could not reach the API: ' + e.message };
  }
}

/** Multipart upload of a file as a new source. */
function upload_(notebook, blob, displayName) {
  var boundary = 'bpt' + Utilities.getUuid().replace(/-/g, '');
  var meta = JSON.stringify({ source: { displayName: displayName } });
  var head = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
             meta + '\r\n--' + boundary + '\r\nContent-Type: ' + blob.getContentType() + '\r\n\r\n';
  var tail = '\r\n--' + boundary + '--';
  var bytes = Utilities.newBlob(head).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(tail).getBytes());
  try {
    var r = UrlFetchApp.fetch(UPLOAD_API + notebook + '/sources:upload?uploadType=multipart', {
      method: 'post',
      contentType: 'multipart/related; boundary=' + boundary,
      payload: bytes,
      headers: authHeaders_(),
      muteHttpExceptions: true
    });
    return describe_(r);
  } catch (e) {
    return { ok: false, note: 'Upload failed: ' + e.message };
  }
}

/** A Google Doc's contents in another format (via the Drive API). */
function exportDoc_(docId, mimeType) {
  var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + docId +
    '/export?mimeType=' + encodeURIComponent(mimeType), {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
  return r.getResponseCode() === 200
    ? { ok: true, text: r.getContentText() }
    : { ok: false, note: 'Could not export the Doc as ' + mimeType + ' (' + r.getResponseCode() + ')' };
}

function describe_(r) {
  var code = r.getResponseCode(), text = r.getContentText(), json = null;
  try { json = JSON.parse(text); } catch (e) { /* not JSON */ }
  if (code >= 200 && code < 300) return { ok: true, json: json, note: '' };
  var msg = json && json.error ? json.error.status + ': ' + json.error.message : text.slice(0, 300);
  var hint = '';
  if (code === 403 && /SERVICE_DISABLED|has not been used|is disabled/i.test(msg)) {
    hint = ' → Turn on the Gemini Notebook API in your Google Cloud project (README step 2).';
  } else if (code === 403 && /scope|insufficient/i.test(msg)) {
    hint = ' → Check appsscript.json has the notebook scope, then re-authorize.';
  } else if (code === 403 && /user.?project|serviceusage|USER_PROJECT_DENIED/i.test(msg)) {
    hint = ' → Check GCP_PROJECT_ID at the top of the script (the project ID, not the number).';
  } else if (code === 403) {
    hint = ' → The API may not be available to this account (preview: personal accounts only; ' +
           'Workspace accounts such as boisestate.edu get PERMISSION_DENIED).';
  } else if (code === 404) {
    hint = ' → Open notebook.google.com once in this account and sign in, then run again.';
  } else if (code === 401) {
    hint = ' → Re-run and approve the permissions prompt.';
  }
  return { ok: false, json: json, note: 'HTTP ' + code + ' ' + msg + hint };
}

function sourceStates_(notebook) {
  var r = call_('get', API + notebook + '/sources?pageSize=50');
  if (!r.ok) return ['   Could not list sources: ' + r.note];
  var list = (r.json && r.json.sources) || [];
  if (!list.length) return ['   (no sources listed)'];
  return list.map(function (s) {
    var icon = s.state === 'ACTIVE' ? '✅' : (s.state === 'FAILED' ? '❌' : '⏳');
    return '   ' + icon + ' ' + (s.displayName || s.name) + ': ' + s.state +
      (s.error && s.error.message ? ' — ' + s.error.message : '');
  });
}

function finish_(lines) {
  Logger.log('\n' + lines.join('\n') +
    '\n\nOpen Gemini Notebook to see the test notebook. The API cannot delete it; ' +
    'delete it there by hand when you are done.');
}
