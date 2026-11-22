// ============================================================
// Blueprint Tools — Code17.gs
// Set Up Gemini Notebook
// Builds the checklist of sources an IDC adds to a project's Gemini
// Notebook: every usable file in the project folder (the folder the
// Blueprint sits in, including subfolders) plus the shared reference
// documents listed in the "How To Set Up And Use a Project Work Station"
// doc. Also shows the project's current design-cycle week (from the
// Dashboard's "Design Cycle:" line and the Course Design and
// Development Schedule calendar).
//
// Gemini Notebook has no API for Workspace accounts yet, so the user
// adds the sources by hand; this tool tells them exactly which ones.
// Read-only: it changes nothing in Drive or in the Blueprint.
// ------------------------------------------------------------
// Last updated on 2026-11-22 at 11:00 MST
// ------------------------------------------------------------
//
// WHERE THE SOURCES COME FROM
//   • Project folder: the Blueprint's parent folder and all subfolders,
//     NOT the Dashboard links (the tool may run before anything is
//     linked). Shortcuts are followed to the real file. Google Docs,
//     Slides, Sheets, PDFs and Word files are offered; anything else is
//     listed under "can't be added".
//   • Reference documents: read live from the "Gemini Notebook Reference
//     List" Sheet (REFERENCE_SHEET_ID_17), one row per document:
//     Category | Name | Link | Notes. Rows without a link show as
//     "no link yet". Edit the Sheet to change the list.
//     Knowledge Base (Confluence) links are swapped for their Google Doc
//     copy in KB_COPY_FOLDER_ID_17, made by the separate KB Copy Job
//     (kb-copy-job/ in this repo). Without a copy they show "not copied".
//   • Signed agreement: found anywhere in the user's Drive by course code
//     + "Development Agreement" (see SIGNED AGREEMENT below).
//   • Program Canvas Specifications: the Dashboard link, else a Drive
//     search by course prefix / program title (see PROGRAM CANVAS
//     SPECIFICATIONS below).
//
// REUSES (Code10.gs): MIME_DOC_10, MIME_SHEET_10, MIME_DOCX_10,
//   MIME_SHORTCUT_10, idFromUrl10_, findDashboardTitle10_, dashboardTab10_,
//   findDashLabel10_. (Code.gs): collectTabs.
//
// HOW THE CYCLE IS WORKED OUT
//   IDCs write the Design Cycle line in many ways ("Fall Cycle 1",
//   "F1", "Fall1", "Fall 1", "FA 01", with or without dates written as
//   "August 31", "Aug 31" or "08/31"). The calendar is the reference, so
//   the line only has to give enough to pick one cycle:
//     1. term + cycle number  ("Fall" + 1), and/or
//     2. a start date, matched to that cycle's Week 1 event.
//   When both are present and disagree, the tool stops and says so
//   rather than guessing. A lone "S" (Spring or Summer?) is treated as
//   unknown and the date decides.
//
// SPLIT FOR TESTING
//   parseDesignCycleLine17, buildCycleInstances17, pickInstance17,
//   describeToday17 and parseReferenceRows17 are pure functions (no Google services), so they can
//   be tested outside Apps Script. Everything that touches DocumentApp or
//   CalendarApp lives in the functions near the top.
//
// TO DO
//   • Multi-Semester List: when the live Google Sheets version is
//     available, put its ID in MULTI_SEMESTER_SHEET_ID_17. It can then
//     (a) cross-check a project's term/cycle by course code and
//     (b) supply the QA and Build start dates for the Gantt chart tool.
//     The copy reviewed on 2026-11-21 was an .xlsx snapshot dated
//     11/02/2026, which Apps Script cannot read directly and which goes
//     stale, so it is intentionally NOT used yet.
// ============================================================

// Course Design and Development Schedule (shared Boise State calendar).
var CDD_CALENDAR_ID_17 =
  'boisestate.edu_cl32njs703rfu5grpft5n4l5d4@group.calendar.google.com';

// Live Multi-Semester List (Google Sheet). Leave blank until available.
var MULTI_SEMESTER_SHEET_ID_17 = '';

// How far back/forward to read the calendar. Cycles run ~15-17 weeks and
// Fall Cycle 2 crosses New Year, so +/- 13 months covers any project an
// IDC is likely to open.
var CAL_WINDOW_DAYS_17 = 400;

// A date on the Design Cycle line counts as "the cycle's start" when it
// is within this many days of the Week 1 event's first day.
var START_DATE_TOLERANCE_DAYS_17 = 6;

var DAY_MS_17 = 24 * 60 * 60 * 1000;

// "Gemini Notebook Reference List" Google Sheet (tab "References":
// Category | Name | Link | Notes). The shared reference documents are
// read from it on every run; edit the Sheet to change the list. Every
// user needs at least view access to it. (Until 2026-11-22 the list was
// read from Greg's planning doc, which is now planning-only.)
var REFERENCE_SHEET_ID_17 = '1991TCf1UoWGRyl_3YdDR9cBMrdS4eouRhr8tnwpjGz4';
var REFERENCE_TAB_17 = 'References';

// Stop scanning the project folder after this many files so a huge
// folder can't make the tool time out. The sidebar says when it stopped.
var MAX_PROJECT_FILES_17 = 500;

var GEMINI_NOTEBOOK_URL_17 = 'https://notebook.google.com/';

// Folder where the KB Copy Job (kb-copy-job/ in this repo, a separate
// script Greg owns) keeps Google Doc copies of Knowledge Base pages.
// Each copy's Drive description holds "confluence-page-id: <id>".
var KB_COPY_FOLDER_ID_17 = '15_H0_97_6ZqU-Pi8nW79PSf2_ARNtxzz';

// File kinds Gemini Notebook accepts as sources (Greg's list, 2026-11-20).
var NOTEBOOK_KINDS_17 = {
  'application/vnd.google-apps.document':     'Google Doc',
  'application/vnd.google-apps.presentation': 'Google Slides',
  'application/vnd.google-apps.spreadsheet':  'Google Sheet',
  'application/pdf':                          'PDF',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word',
  'application/msword':                       'Word'
};


// ── SIDEBAR OPENER ───────────────────────────────────────────

function showSetupNotebookSidebar17() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar17')
    .setTitle('🎓 Set Up Gemini Notebook')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}


// ── MAIN SERVER FUNCTION (called from Sidebar17) ─────────────

/**
 * Everything the sidebar shows: the design-cycle week, the project-folder
 * sources and the reference documents. Each part fails on its own, so a
 * problem with the calendar never hides the source list.
 */
function getNotebookSetup17() {
  var doc = DocumentApp.getActiveDocument();
  var out = {
    notebookUrl: GEMINI_NOTEBOOK_URL_17,
    week: safe17_(getProjectWeek17),
    project: safe17_(function () { return scanProjectFolder17_(doc); }),
    reference: safe17_(readReferenceDocs17_)
  };
  out.agreement = safe17_(function () { return findAgreement17_(doc, out.week); });
  out.specs = safe17_(function () { return findCanvasSpecs17_(doc, out.agreement); });
  if (out.specs.ok && out.specs.pick && out.project.ok) {
    out.project.files = out.project.files.filter(function (f) {
      return f.id !== out.specs.pick.id;            // list it once, as the specs
    });
  }
  if (out.agreement.ok && out.agreement.pick && out.project.ok) {
    out.project.files = out.project.files.filter(function (f) {
      return f.id !== out.agreement.pick.id;        // list it once, as the agreement
    });
  }
  // A reference document that is also in the project folder is listed
  // once, under the project folder.
  if (out.project.ok && out.reference.ok) {
    var inProject = {};
    out.project.files.forEach(function (f) { inProject[f.id] = true; });
    out.reference.groups.forEach(function (g) {
      g.items.forEach(function (it) {
        if (it.id && inProject[it.id]) it.alsoInProject = true;
      });
    });
  }
  return out;
}

function safe17_(fn) {
  try { return fn(); }
  catch (e) {
    return { ok: false, message: 'Something went wrong: ' + (e && e.message ? e.message : e) };
  }
}


// ── PROJECT FOLDER ───────────────────────────────────────────

/**
 * Lists the Blueprint's parent folder and every subfolder. Returns
 * { ok, folderName, folderUrl, files:[...], unsupported:[...], truncated }.
 * Each file: { id, name, kind, url, path, isBlueprint, shortcut }.
 */
function scanProjectFolder17_(doc) {
  var parents = DriveApp.getFileById(doc.getId()).getParents();
  if (!parents.hasNext()) {
    return { ok: false, message:
      'This Blueprint isn\'t inside a Drive folder, so there is no project ' +
      'folder to read. Move it into the project folder and run this again.' };
  }
  var root = parents.next();
  var files = [], unsupported = [], seen = {}, visited = {};
  var queue = [{ folder: root, path: '' }];
  var count = 0, truncated = false;

  while (queue.length && !truncated) {
    var cur = queue.shift();
    var fid = cur.folder.getId();
    if (visited[fid]) continue;
    visited[fid] = true;

    var it = cur.folder.getFiles();
    while (it.hasNext()) {
      if (count >= MAX_PROJECT_FILES_17) { truncated = true; break; }
      var f = it.next(); count++;
      var id = f.getId(), mime = f.getMimeType(), shortcut = false;
      if (mime === MIME_SHORTCUT_10) {
        var tid = f.getTargetId();
        if (!tid) continue;
        if (f.getTargetMimeType() === 'application/vnd.google-apps.folder') {
          unsupported.push({ name: f.getName(), path: cur.path,
            why: 'Shortcut to another folder (not opened)' });
          continue;
        }
        id = tid; mime = f.getTargetMimeType(); shortcut = true;
      }
      if (seen[id]) continue;          // same file reached twice
      seen[id] = true;
      var kind = NOTEBOOK_KINDS_17[mime];
      if (!kind) {
        unsupported.push({ name: f.getName(), path: cur.path,
                           why: friendlyMime17_(mime) });
        continue;
      }
      files.push({ id: id, name: f.getName(), kind: kind,
                   url: fileUrl17_(id, mime), path: cur.path,
                   isBlueprint: id === doc.getId(), shortcut: shortcut });
    }

    var subs = cur.folder.getFolders();
    while (subs.hasNext()) {
      var sub = subs.next();
      queue.push({ folder: sub,
                   path: (cur.path ? cur.path + ' / ' : '') + sub.getName() });
    }
  }

  files.sort(function (a, b) {
    if (a.isBlueprint !== b.isBlueprint) return a.isBlueprint ? -1 : 1;
    if (a.path !== b.path) return a.path < b.path ? -1 : 1;
    return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
  });
  return { ok: true, folderName: root.getName(), folderUrl: root.getUrl(),
           files: files, unsupported: unsupported, truncated: truncated };
}

function fileUrl17_(id, mime) {
  if (mime === MIME_DOC_10)   return 'https://docs.google.com/document/d/' + id + '/edit';
  if (mime === MIME_SHEET_10) return 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
  if (mime === 'application/vnd.google-apps.presentation')
    return 'https://docs.google.com/presentation/d/' + id + '/edit';
  return 'https://drive.google.com/file/d/' + id + '/view';
}

function friendlyMime17_(mime) {
  if (/^video\//.test(mime)) return 'Video';
  if (/^image\//.test(mime)) return 'Image';
  if (/^audio\//.test(mime)) return 'Audio';
  if (/zip|compressed/.test(mime)) return 'Zip file';
  if (/presentationml|powerpoint/.test(mime)) return 'PowerPoint (save as Google Slides to add it)';
  if (/spreadsheetml|ms-excel/.test(mime)) return 'Excel (save as a Google Sheet to add it)';
  if (mime === 'application/vnd.google-apps.form') return 'Google Form';
  if (mime === 'application/vnd.google-apps.drawing') return 'Google Drawing';
  return 'File type not accepted by Gemini Notebook';
}


// ── SIGNED AGREEMENT ─────────────────────────────────────────
//
// Every project has an "Online Course Design & Development Agreement for
// <Term> <Year>" signed by the faculty developer. It is NOT in the project
// folder and its location varies, so the tool searches all of Drive the
// user can see: file name containing the course code, and text containing
// "Development Agreement". When several match, the one whose term matches
// the Blueprint's design cycle wins, then an original over a "Copy of",
// then the most recently changed. It contains the faculty developer's ID
// number and contact details, so the sidebar shows a caution.

/** Returns { ok, course, pick, others, message }. */
function findAgreement17_(doc, week) {
  var titleLine = findDashboardTitle10_(doc);
  var course = parseCourseCode17((titleLine && titleLine.text) || '') ||
               parseCourseCode17(doc.getName());
  if (!course) {
    return { ok: false, message:
      'I couldn\'t tell the course code (for example "CS 321") from the Dashboard ' +
      'title or the Blueprint\'s name, so I couldn\'t look for the signed agreement.' };
  }

  var q = "title contains '" + course.number + "' and fullText contains " +
          "'Development Agreement' and trashed = false";
  var it = DriveApp.searchFiles(q), cands = [], n = 0;
  while (it.hasNext() && n < 40) {
    var f = it.next(); n++;
    var name = f.getName(), mime = f.getMimeType();
    if (f.getId() === doc.getId()) continue;
    if (!nameMatchesCourse17(name, course)) continue;
    if (/blueprint/i.test(name)) continue;
    if (mime !== 'application/vnd.google-apps.document' && mime !== 'application/pdf') continue;
    cands.push({ id: f.getId(), name: name, mime: mime,
                 kind: NOTEBOOK_KINDS_17[mime], updated: f.getLastUpdated().getTime(),
                 url: fileUrl17_(f.getId(), mime) });
  }
  if (!cands.length) {
    return { ok: true, course: course.code, prefix: course.prefix, pick: null, others: [], message:
      'No signed agreement for ' + course.code + ' was found in the Drive files you can open. ' +
      'It may be named differently, or you may not have access to it.' };
  }

  // Read the term from the first few candidates that are Google Docs.
  cands.slice(0, 6).forEach(function (c) {
    if (c.mime !== 'application/vnd.google-apps.document') return;
    try {
      var text = DocumentApp.openById(c.id).getBody().getText().slice(0, 4000);
      c.term = agreementTerm17(text);
      c.program = agreementProgram17(text);
    } catch (e) { /* no access to the content: rank without the term */ }
  });
  var target = (week && week.ok) ? { term: week.term, year: week.termYear } : null;
  var ranked = rankAgreements17(cands, target);
  var pick = ranked[0];
  pick.termMatches = !!(target && pick.term &&
    pick.term.term === target.term && pick.term.year === target.year);
  pick.termLabel = pick.term ? pick.term.term + ' ' + pick.term.year : '';
  return { ok: true, course: course.code, prefix: course.prefix, pick: pick,
           program: pick.program || '',
           others: ranked.slice(1, 5).map(function (c) {
             return { name: c.name, url: c.url,
                      termLabel: c.term ? c.term.term + ' ' + c.term.year : '' };
           }),
           target: target ? target.term + ' ' + target.year : '' };
}

/** PURE. Course code from a title like "CS 321: Data Structures",
 *  "UX/UX-PRO 300", "SOCWRK 481/499" or "CS 321 Blueprint". */
function parseCourseCode17(text) {
  var m = String(text || '').match(/\b([A-Z]{2,8}(?:[\/-][A-Z]{2,8})*)\s*-?\s*(\d{3}[A-Z]?)(?!\d)/);
  if (!m) return null;
  return { prefix: m[1], number: m[2], code: m[1] + ' ' + m[2] };
}

/** PURE. Does a file name name this course? ("CS 321-...", "CS321 ...") */
function nameMatchesCourse17(name, course) {
  var prefix = course.prefix.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&');
  return new RegExp('(^|[^A-Za-z])' + prefix + '\\s*-?\\s*' + course.number + '(?![0-9])', 'i')
    .test(String(name || ''));
}

/** PURE. {term, year} from "...Development Agreement for Fall 2026". */
function agreementTerm17(text) {
  var m = String(text || '').match(/agreement\s+for\s+(fall|spring|summer)\s+(\d{4})/i);
  if (!m) return null;
  return { term: m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase(), year: +m[2] };
}

/** PURE. Program title from the agreement's "Program Title (if applicable):"
 *  row, or '' for blank / "n/a". Table cells come through on separate lines. */
function agreementProgram17(text) {
  var m = String(text || '').match(/program\s+title[^:\n]*:[ \t]*\n?[ \t]*([^\n]*)/i);
  var v = m ? m[1].trim() : '';
  return /^(n\/?a|none|-+)?$/i.test(v) ? '' : v;
}

/** PURE. Best agreement first. */
function rankAgreements17(cands, target) {
  function score(c) {
    var s = 0;
    if (target && c.term && c.term.term === target.term && c.term.year === target.year) s += 4;
    if (!/^copy of /i.test(c.name)) s += 2;
    if (c.mime === 'application/vnd.google-apps.document') s += 1;
    return s;
  }
  return cands.slice().sort(function (a, b) {
    return (score(b) - score(a)) || (b.updated - a.updated);
  });
}


// ── PROGRAM CANVAS SPECIFICATIONS ────────────────────────────
//
// Program courses follow a program's "Canvas Specifications" doc; faculty
// courses (not part of a program) have none, and some programs have none.
// 1. If the Dashboard's "Program Canvas Specifications" line is linked,
//    use that.
// 2. Otherwise, unless the Dashboard's Project Type says "Faculty Course",
//    search Drive for files named "...Canvas Spec..." and match them to the
//    course prefix ("RESPCARE Canvas Specifications", "Canvas Specifications
//    - BUSMGT, ...", "BAS/IPS Canvas Specifications For...") or to words of
//    the program title on the signed agreement ("Canvas Specifications -
//    Public Service"). Names are inconsistent, so a search result is always
//    shown as "check this is right", never as certain. Archives, copies,
//    templates and examples are skipped.

/** Returns { ok, status:'linked'|'found'|'none', pick, others, message }. */
function findCanvasSpecs17_(doc, agreement) {
  var dash = dashboardTab10_(doc);
  var lab = findDashLabel10_(dash.body, '^\\s*Program Canvas Specifications?');
  if (lab && lab.existing) {
    var id = idFromUrl10_(lab.existing);
    var file = null;
    try { file = id ? DriveApp.getFileById(id) : null; } catch (e) { file = null; }
    if (file) {
      var mime = file.getMimeType();
      return { ok: true, status: 'linked', others: [],
               pick: { id: id, name: file.getName(), url: fileUrl17_(id, mime),
                       kind: NOTEBOOK_KINDS_17[mime] || friendlyMime17_(mime) } };
    }
    return { ok: true, status: 'none', others: [], message:
      'The Dashboard links Program Canvas Specifications, but you can\'t open that file ' +
      '(no access, or it was deleted).' };
  }

  var projectType = dashValue17_(dash.body, /project\s*type\s*:/i);
  var facultyCourse = /faculty\s+course/i.test(projectType);

  var titleLine = findDashboardTitle10_(doc);
  var course = parseCourseCode17((titleLine && titleLine.text) || '') ||
               parseCourseCode17(doc.getName());
  var program = agreement && agreement.ok ? (agreement.program || '') : '';

  var cands = [];
  var it = DriveApp.searchFiles("title contains 'Canvas Spec' and trashed = false");
  var n = 0;
  while (it.hasNext() && n < 200) {
    var f = it.next(); n++;
    var mime2 = f.getMimeType();
    if (!NOTEBOOK_KINDS_17[mime2]) continue;
    cands.push({ id: f.getId(), name: f.getName(), mime: mime2,
                 kind: NOTEBOOK_KINDS_17[mime2], updated: f.getLastUpdated().getTime(),
                 url: fileUrl17_(f.getId(), mime2) });
  }
  var ranked = rankSpecs17(cands, course ? course.prefix : '', program);
  var why = facultyCourse ? 'This is a faculty course, so no program specifications are expected.' : '';

  if (!ranked.length) {
    return { ok: true, status: 'none', others: [], facultyCourse: facultyCourse, message: why ||
      'No Canvas specifications doc was found for ' + (course ? course.prefix : 'this course') +
      (program ? ' or the "' + program + '" program' : '') +
      '. The program may not have one. If it does, link it on the Dashboard\'s ' +
      '"Program Canvas Specifications" line.' };
  }
  if (facultyCourse) {
    // A faculty course normally has none; offer matches without picking one.
    return { ok: true, status: 'none', facultyCourse: true, message: why,
             others: ranked.slice(0, 4).map(slimSpec17_) };
  }
  return { ok: true, status: 'found', pick: ranked[0],
           others: ranked.slice(1, 5).map(slimSpec17_) };
}

function slimSpec17_(c) { return { name: c.name, url: c.url, matchedOn: c.matchedOn }; }

/** Text after a Dashboard label such as "Project Type:" on the same line. */
function dashValue17_(body, labelRe) {
  var lines = body.getText().split(/\n/);
  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(labelRe);
    if (m) return lines[i].slice(m.index + m[0].length).trim();
  }
  return '';
}

/**
 * PURE. Canvas specification files that match the course prefix or the
 * program title, best first. Each kept file gets `matchedOn`.
 */
function rankSpecs17(cands, prefix, program) {
  var skip = /^(archive|copy of|old\b)|\b(template|example|sample|draft)\b/i;
  var prefixParts = String(prefix || '').toUpperCase().split(/[\/-]/)
    .filter(function (p) { return p.length >= 2; });
  var stop = { program: 1, programs: 1, online: 1, degree: 1, certificate: 1,
               the: 1, and: 1, of: 1, for: 1, in: 1, bachelor: 1, master: 1,
               masters: 1, science: 1, arts: 1, applied: 1, studies: 1 };
  var progWords = String(program || '').toLowerCase().split(/[^a-z]+/)
    .filter(function (w) { return w.length >= 4 && !stop[w]; });
  var small = { of: 1, and: 1, the: 1, for: 1, in: 1 };
  var initials = String(program || '').split(/[^A-Za-z]+/)
    .filter(function (w) { return w && !small[w.toLowerCase()]; })
    .map(function (w) { return w.charAt(0).toUpperCase(); }).join('');
  if (initials.length > 6) initials = '';

  var out = [];
  cands.forEach(function (c) {
    if (!/canvas\s*spec/i.test(c.name) || skip.test(c.name)) return;
    var tokens = c.name.toUpperCase().split(/[^A-Z0-9]+/);
    var score = 0, on = '';
    if (prefixParts.length && prefixParts.every(function (p) { return tokens.indexOf(p) >= 0; })) {
      score += 10; on = 'course prefix ' + prefix;
    } else if (prefixParts.some(function (p) { return p.length >= 3 && tokens.indexOf(p) >= 0; })) {
      score += 6; on = 'part of course prefix ' + prefix;
    }
    var lower = c.name.toLowerCase();
    var hits = progWords.filter(function (w) {
      return lower.indexOf(w) >= 0 || lower.indexOf(w.slice(0, 5)) >= 0;
    });
    if (hits.length) {
      score += 3 * hits.length;
      on = on ? on + ' and program title' : 'program title "' + program + '"';
    } else if (initials.length >= 2 && tokens.indexOf(initials) >= 0) {
      // "Online Degree Program" -> ODP, "Bachelor of Applied Science" -> BAS
      score += 5;
      on = on ? on + ' and program initials' : 'program initials ' + initials + ' ("' + program + '")';
    }
    if (!score) return;
    out.push({ id: c.id, name: c.name, mime: c.mime, kind: c.kind, url: c.url,
               updated: c.updated, score: score, matchedOn: on });
  });
  return out.sort(function (a, b) { return (b.score - a.score) || (b.updated - a.updated); });
}


// ── REFERENCE DOCUMENTS (from the reference list Sheet) ──────

/**
 * Reads the reference list Sheet and checks each linked file.
 * Returns { ok, docUrl, groups:[{category, items:[...]}] }.
 * Each item: { name, url, id, kind, problem, noLink }.
 */
function readReferenceDocs17_() {
  var sheet;
  try {
    var ss = SpreadsheetApp.openById(REFERENCE_SHEET_ID_17);
    sheet = ss.getSheetByName(REFERENCE_TAB_17) || ss.getSheets()[0];
  } catch (e) {
    return { ok: false, message:
      'I couldn\'t open the "Gemini Notebook Reference List" Sheet, which lists the ' +
      'reference documents. Ask Greg Snow for view access.' };
  }
  var range = sheet.getDataRange();
  var values = range.getDisplayValues();
  var rich = range.getRichTextValues();
  // A link may be typed as a URL or put on the cell's text as a hyperlink.
  var rows = values.map(function (r, i) {
    return r.map(function (v, j) {
      var link = rich[i][j] && rich[i][j].getLinkUrl();
      return link && !/^https?:/i.test(v) ? link : v;
    });
  });

  var groups = parseReferenceRows17(rows);
  var kbCopies = null;   // read the copies folder only if a KB link appears
  groups.forEach(function (g) {
    g.items.forEach(function (it) {
      if (!it.url) return;
      var pageId = confluencePageId17_(it.url);
      if (pageId) {
        if (!kbCopies) kbCopies = listKbCopies17_();
        var copy = kbCopies[pageId];
        if (copy) {
          it.id = copy.id; it.url = copy.url; it.kind = 'Google Doc'; it.kbCopy = true;
        } else {
          it.kind = 'Knowledge Base page';
          it.problem = 'Not copied to a Google Doc yet (the KB Copy Job hasn\'t run for it). ' +
                       'Gemini Notebook can\'t open pages that need a Confluence sign-in.';
        }
        return;
      }
      var id = /docs\.google\.com|drive\.google\.com/.test(it.url)
        ? idFromUrl10_(it.url) : null;
      if (!id) { it.kind = 'Web page'; return; }   // any other website
      it.id = id;
      try {
        var f = DriveApp.getFileById(id);
        var mime = f.getMimeType();
        if (mime === MIME_SHORTCUT_10) { mime = f.getTargetMimeType(); it.id = f.getTargetId(); }
        it.kind = NOTEBOOK_KINDS_17[mime] || friendlyMime17_(mime);
        if (!NOTEBOOK_KINDS_17[mime]) it.problem = 'Gemini Notebook can\'t add this file type.';
        it.url = fileUrl17_(it.id, mime);
      } catch (e) {
        it.problem = 'You don\'t have access to this file, or it was deleted.';
      }
    });
  });
  return { ok: true, groups: groups,
           docUrl: 'https://docs.google.com/spreadsheets/d/' + REFERENCE_SHEET_ID_17 + '/edit' };
}

/** Confluence page ID in a Knowledge Base link, or null. */
function confluencePageId17_(url) {
  var s = String(url || '');
  if (!/^https?:\/\/[^\/]*atlassian\.net\/wiki\//i.test(s)) return null;
  var m = s.match(/\/pages\/(\d+)/) || s.match(/[?&]pageId=(\d+)/);
  return m ? m[1] : null;
}

/** The KB Copy Job's Docs, keyed by Confluence page ID. */
function listKbCopies17_() {
  var map = {};
  try {
    var it = DriveApp.getFolderById(KB_COPY_FOLDER_ID_17)
      .getFilesByType('application/vnd.google-apps.document');
    while (it.hasNext()) {
      var f = it.next();
      var m = String(f.getDescription() || '').match(/confluence-page-id:\s*(\d+)/);
      if (m) map[m[1]] = { id: f.getId(),
        url: 'https://docs.google.com/document/d/' + f.getId() + '/edit' };
    }
  } catch (e) { /* no access to the folder: links show as "not copied yet" */ }
  return map;
}


/**
 * PURE. Turns the reference Sheet's rows into groups by category.
 * Row 1 is the header; columns are found by name (Category, Name, Link),
 * so they can be reordered. Blank rows are skipped; a row with a name but
 * no link shows as "no link yet".
 * @param {string[][]} rows  display values, header first
 */
function parseReferenceRows17(rows) {
  if (!rows || !rows.length) return [];
  var head = rows[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  var col = function (name, fallback) {
    var i = head.indexOf(name); return i >= 0 ? i : fallback;
  };
  var cCat = col('category', 0), cName = col('name', 1), cLink = col('link', 2);
  var groups = [], byName = {};
  rows.slice(1).forEach(function (r) {
    var name = String(r[cName] || '').replace(/\s+/g, ' ').trim();
    var link = String(r[cLink] || '').trim();
    if (!name && !link) return;
    if (!name) name = link;
    var category = String(r[cCat] || '').trim() || 'Other';
    if (!byName[category]) {
      byName[category] = { category: category, items: [] };
      groups.push(byName[category]);
    }
    var url = /^https?:\/\//i.test(link) ? link : null;
    byName[category].items.push({ name: name, url: url, noLink: !url });
  });
  return groups;
}


/**
 * Works out the project's design cycle and where "today" falls in it.
 * Returns a plain object the sidebar renders. Never throws for expected
 * problems — it returns { ok:false, message } in plain language instead.
 */
function getProjectWeek17() {
  var doc = DocumentApp.getActiveDocument();
  var tz  = Session.getScriptTimeZone();

  // 1. Read the Design Cycle line.
  var line = readDesignCycleLine17(doc);
  if (line === null) {
    return { ok: false, message:
      'I couldn\'t find a "Design Cycle:" line on the Dashboard tab. ' +
      'Check that the Blueprint has a Dashboard tab and that the line ' +
      'starts with "Design Cycle:".' };
  }
  var parsed = parseDesignCycleLine17(line);
  if (!parsed.term && !parsed.cycle && !parsed.date) {
    return { ok: false, line: line, message:
      'The Design Cycle line doesn\'t say which cycle this is. Add the ' +
      'term and cycle (for example "Fall Cycle 1") or the start date, ' +
      'then run this again.' };
  }

  // 2. Read the calendar.
  var cal = CalendarApp.getCalendarById(CDD_CALENDAR_ID_17);
  if (!cal) {
    return { ok: false, line: line, message:
      'I couldn\'t open the "Course Design and Development Schedule" ' +
      'calendar from your account. In Google Calendar, check that it is ' +
      'listed under "Other calendars". If it isn\'t, add it with ' +
      '"Subscribe to calendar" using this ID: ' + CDD_CALENDAR_ID_17 };
  }
  var today = startOfDay17(new Date());
  var from  = new Date(today.getTime() - CAL_WINDOW_DAYS_17 * DAY_MS_17);
  var to    = new Date(today.getTime() + CAL_WINDOW_DAYS_17 * DAY_MS_17);
  var events = cal.getEvents(from, to).map(function (ev) {
    return {
      title: ev.getTitle(),
      start: ev.getStartTime(),
      // All-day events end at midnight AFTER the last day; store the last
      // day itself so "Nov 16 – Nov 20" reads the way people expect.
      end: ev.isAllDayEvent()
        ? new Date(ev.getEndTime().getTime() - DAY_MS_17)
        : ev.getEndTime()
    };
  });
  var instances = buildCycleInstances17(events);
  if (!instances.length) {
    return { ok: false, line: line, message:
      'The calendar opened, but I found no "… Cycle N: Week N (…)" ' +
      'events in the 13 months before or after today.' };
  }

  // 3. Pick the cycle and describe today.
  var pick = pickInstance17(parsed, instances, today);
  if (!pick.instance) {
    return { ok: false, line: line, parsed: summariseParsed17(parsed, tz),
             message: pick.message };
  }
  var status = describeToday17(pick.instance, today);

  return {
    ok: true,
    docName: doc.getName(),
    line: line,
    parsed: summariseParsed17(parsed, tz),
    note: pick.note || '',
    cycleLabel: pick.instance.term + ' Cycle ' + pick.instance.cycle,
    term: pick.instance.term,
    termYear: pick.instance.start.getFullYear(),
    cycleDates: fmt17(pick.instance.start, tz) + ' – ' +
                fmt17(pick.instance.end, tz) + ', ' +
                Utilities.formatDate(pick.instance.end, tz, 'yyyy'),
    totalWeeks: pick.instance.weeks.length,
    status: formatStatus17(status, tz),
    weeks: pick.instance.weeks.map(function (w) {
      return {
        week: w.week,
        phase: w.phase,
        dates: fmt17(w.start, tz) + ' – ' + fmt17(w.end, tz),
        isCurrent: !!(status.current && status.current.week === w.week)
      };
    })
  };
}


// ── READING THE DASHBOARD ────────────────────────────────────

/**
 * Returns the text after "Design Cycle:" on the Dashboard tab, or null
 * if there is no such line. Handles three layouts:
 *   • "Design Cycle: Fall Cycle 1" in one paragraph (the template),
 *   • the label in one table cell and the value in the next cell,
 *   • the label on its own line and the value on the next paragraph.
 * Date smart chips are read through their display text.
 */
function readDesignCycleLine17(doc) {
  var dash = null;
  collectTabs(doc).forEach(function (t) {
    if (!dash && /dashboard/i.test(t.title)) dash = t;
  });
  if (!dash) return null;

  var hit = dash.body.findText('(?i)design\\s*cycle\\s*:?');
  if (!hit) return null;

  var para = hit.getElement().getParent();
  // findText can land inside a nested element; climb to the paragraph.
  while (para && para.getType() !== DocumentApp.ElementType.PARAGRAPH &&
         para.getType() !== DocumentApp.ElementType.LIST_ITEM) {
    para = para.getParent();
  }
  if (!para) return null;

  // A paragraph may hold several Dashboard lines joined by soft returns.
  var segments = elementText17(para).split(/[\n\r\u000b]/);
  var seg = '';
  for (var i = 0; i < segments.length; i++) {
    if (/design\s*cycle/i.test(segments[i])) { seg = segments[i]; break; }
  }
  var value = seg.replace(/^.*?design\s*cycle\s*:?\s*/i, '').trim();
  if (value) return value;

  // Value not on the same line: try the next table cell, then the next
  // paragraph.
  var parent = para.getParent();
  if (parent && parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
    var next = parent.getNextSibling();
    if (next) {
      value = elementText17(next).trim();
      if (value) return value;
    }
  }
  var sib = para.getNextSibling();
  if (sib) {
    value = elementText17(sib).trim();
    if (value) return value;
  }
  return '';
}

/** Text of an element, including date chips, people chips and link chips. */
function elementText17(el) {
  var T = DocumentApp.ElementType;
  var type = el.getType();
  try {
    if (type === T.TEXT) return el.getText();
    if (type === T.DATE) return el.getDisplayText();
    if (type === T.PERSON) return el.getName();
    if (type === T.RICH_LINK) return el.getTitle();
  } catch (e) { return ''; }
  if (typeof el.getNumChildren === 'function') {
    var out = '';
    for (var i = 0; i < el.getNumChildren(); i++) {
      out += elementText17(el.getChild(i));
    }
    if (type === T.TABLE_CELL || type === T.TABLE_ROW) return out + '\n';
    return out;
  }
  return typeof el.getText === 'function' ? el.getText() : '';
}


// ── PURE LOGIC (no Google services below this line except Utilities
//    in the formatters at the very bottom) ─────────────────────

var TERM_WORDS_17 = {
  f: 'Fall', fa: 'Fall', fal: 'Fall', fall: 'Fall',
  sp: 'Spring', spr: 'Spring', spg: 'Spring', spring: 'Spring',
  su: 'Summer', sum: 'Summer', smr: 'Summer', summer: 'Summer',
  s: null   // Spring or Summer — ambiguous, let the date decide.
};

var MONTHS_17 = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
  jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
};

/**
 * Reads term, cycle number, optional term year and the first date from
 * the free-text Design Cycle value.
 * @returns {{term:?string, termAmbiguous:boolean, cycle:?number,
 *            termYear:?number, date:?{month,day,year:?number}}}
 */
function parseDesignCycleLine17(text) {
  var s = String(text || '').replace(/[–—]/g, '-');
  var out = { term: null, termAmbiguous: false, cycle: null,
              termYear: null, date: null };

  // A: term and number together — "Fall Cycle 1", "F1", "Fall1",
  // "FA 01", "Sp-2", "Su Cycle #1". (?!\d) stops "FA27" or "Spring 2027"
  // being read as cycle 2.
  var combined = /\b(fall|fal|fa|f|spring|spr|spg|sp|summer|sum|smr|su|s)\s*[-:.]?\s*(?:cycle|cyc|c)?\s*[-:#.]?\s*0?([1-9])(?!\d)/i;
  var m = s.match(combined);
  if (m) {
    setTerm17(out, m[1]);
    out.cycle = parseInt(m[2], 10);
  }

  // B: pieces apart — "Fall 2026 Cycle 1", "Cycle 2 (Spring)".
  if (!out.cycle) {
    var c = s.match(/\bcyc(?:le)?\s*[-:#.]?\s*0?([1-9])(?!\d)/i);
    if (c) out.cycle = parseInt(c[1], 10);
  }
  if (!out.term && !out.termAmbiguous) {
    var t = s.match(/\b(fall|spring|summer)\b/i) ||
            s.match(/\b(fa|sp|su)\b/i);
    if (t) setTerm17(out, t[1]);
  }

  // Term year: "Fall 2026" or "FA26"/"FA 26" next to the term word.
  var y = s.match(/\b(?:fall|spring|summer|fa|sp|su)\s*'?(20\d{2}|\d{2})(?!\d)/i);
  if (y) {
    var yr = parseInt(y[1], 10);
    // Two digits next to a term could be a cycle typo ("F1" handled above
    // already took single digits), so only accept 2 digits >= 20.
    if (y[1].length === 4) out.termYear = yr;
    else if (yr >= 20) out.termYear = 2000 + yr;
  }

  out.date = firstDate17(s);
  return out;
}

function setTerm17(out, word) {
  var key = word.toLowerCase();
  if (TERM_WORDS_17.hasOwnProperty(key)) {
    out.term = TERM_WORDS_17[key];
    out.termAmbiguous = TERM_WORDS_17[key] === null;
  }
}

/** First recognisable date in the text, as {month, day, year|null}. */
function firstDate17(s) {
  var found = [];

  var named = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{4}))?(?!\d)/ig;
  var m;
  while ((m = named.exec(s)) !== null) {
    found.push({ at: m.index, month: MONTHS_17[m[1].toLowerCase()],
                 day: +m[2], year: m[3] ? +m[3] : null });
  }

  var numeric = /\b(\d{1,2})\/(\d{1,2})(?:\/(\d{2}|\d{4}))?(?!\d)/g;
  while ((m = numeric.exec(s)) !== null) {
    var yr = m[3] ? (m[3].length === 2 ? 2000 + +m[3] : +m[3]) : null;
    found.push({ at: m.index, month: +m[1], day: +m[2], year: yr });
  }

  found = found.filter(function (d) {
    return d.month >= 1 && d.month <= 12 && d.day >= 1 && d.day <= 31;
  }).sort(function (a, b) { return a.at - b.at; });

  return found.length
    ? { month: found[0].month, day: found[0].day, year: found[0].year }
    : null;
}

/**
 * Groups calendar events titled "<Term> Cycle <n>: Week <n> (<Phase>)"
 * into cycle instances (one per term/cycle/year).
 * @param {Array<{title,start:Date,end:Date}>} events  end = last day
 * @returns {Array<{term,cycle,start,end,weeks:Array<{week,phase,start,end}>}>}
 */
function buildCycleInstances17(events) {
  var re = /^\s*(fall|spring|summer)\s+cycle\s+(\d+)\s*:\s*week\s+(\d+)\s*(?:\((.*)\))?/i;
  var byKey = {};
  events.forEach(function (ev) {
    var m = re.exec(ev.title || '');
    if (!m) return;
    var term = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
    var key = term + '|' + m[2];
    (byKey[key] = byKey[key] || []).push({
      term: term, cycle: +m[2], week: +m[3],
      phase: (m[4] || '').replace(/\)\s*\(.*$/, '').trim(),
      start: ev.start, end: ev.end
    });
  });

  var instances = [];
  Object.keys(byKey).forEach(function (key) {
    var list = byKey[key].sort(function (a, b) { return a.start - b.start; });
    var cur = null;
    list.forEach(function (w) {
      var gap = cur ? (w.start - cur.end) / DAY_MS_17 : 0;
      var prevWeek = cur ? cur.weeks[cur.weeks.length - 1].week : 0;
      // New instance when the week count restarts or there is a long gap
      // (the same term/cycle a year later).
      if (!cur || w.week <= prevWeek || gap > 60) {
        cur = { term: w.term, cycle: w.cycle, start: w.start, end: w.end,
                weeks: [] };
        instances.push(cur);
      }
      // Skip duplicate copies of the same week (e.g. a re-posted event).
      if (prevWeek === w.week && cur.weeks.length) return;
      cur.weeks.push({ week: w.week, phase: w.phase,
                       start: w.start, end: w.end });
      if (w.end > cur.end) cur.end = w.end;
    });
  });
  // Instances whose first week on the calendar isn't Week 1 were cut off
  // by the read window; keep them but use their real first week's date.
  instances.forEach(function (inst) {
    inst.start = inst.weeks[0].start;
    inst.hasWeek1 = inst.weeks[0].week === 1;
  });
  return instances;
}

/**
 * Chooses the cycle instance the Design Cycle line refers to.
 * @returns {{instance:?Object, message?:string, note?:string}}
 */
function pickInstance17(parsed, instances, today) {
  var label = function (i) {
    return i.term + ' Cycle ' + i.cycle + ' (starts ' + shortDate17(i.start) + ')';
  };

  var pool = instances.slice();
  if (parsed.term)  pool = pool.filter(function (i) { return i.term === parsed.term; });
  // "S" can only be Spring or Summer, never Fall.
  if (parsed.termAmbiguous) pool = pool.filter(function (i) { return i.term !== 'Fall'; });
  if (parsed.cycle) pool = pool.filter(function (i) { return i.cycle === parsed.cycle; });
  if (parsed.termYear) {
    var yp = pool.filter(function (i) {
      return i.hasWeek1 && i.start.getFullYear() === parsed.termYear;
    });
    if (yp.length) pool = yp;
  }

  if (!pool.length) {
    var what = [parsed.term, parsed.cycle ? 'Cycle ' + parsed.cycle : '']
      .filter(Boolean).join(' ');
    return { instance: null, message:
      'The calendar has no "' + what + '" in the 13 months around today. ' +
      'Check the Design Cycle line on the Dashboard.' };
  }

  // A date was given: it must line up with a Week 1.
  if (parsed.date) {
    var dateHits = pool.filter(function (i) {
      return i.hasWeek1 && matchesStart17(parsed.date, i.start);
    });
    if (dateHits.length === 1) {
      return { instance: dateHits[0],
               note: parsed.term && parsed.cycle ? '' :
                 'Matched by the start date on the Design Cycle line.' };
    }
    // Same cycle in different years (a date written without a year, like
    // "August 31", fits every Fall Cycle 1). Take the one nearest today.
    var sameCycle = dateHits.every(function (i) {
      return i.term === dateHits[0].term && i.cycle === dateHits[0].cycle;
    });
    if (dateHits.length > 1 && sameCycle) {
      dateHits.sort(function (a, b) {
        return distance17(a, today) - distance17(b, today);
      });
      return { instance: dateHits[0],
               note: (parsed.date.year || (parsed.term && parsed.cycle)) ? '' :
        'The date on the Design Cycle line has no year, so I used the ' +
        dateHits[0].term + ' Cycle ' + dateHits[0].cycle + ' closest to today (starts ' +
        shortDate17(dateHits[0].start) + '). Add the year if that\'s wrong.' };
    }
    if (dateHits.length > 1) {
      return { instance: null, message:
        'The start date matches more than one cycle: ' +
        dateHits.map(label).join('; ') +
        '. Add the term and cycle number (for example "Fall Cycle 1").' };
    }
    // No hit. If the term and cycle were both given, the date disagrees.
    if (parsed.term && parsed.cycle) {
      return { instance: null, message:
        'The Design Cycle line says ' + parsed.term + ' Cycle ' + parsed.cycle +
        ', but its date (' + parsed.date.month + '/' + parsed.date.day +
        ') isn\'t when that cycle starts on the calendar. Calendar start ' +
        'date(s): ' + pool.map(label).join('; ') +
        '. Fix whichever part is wrong and run this again.' };
    }
    return { instance: null, message:
      'The date on the Design Cycle line (' + parsed.date.month + '/' +
      parsed.date.day + ') doesn\'t match the start of any cycle on the ' +
      'calendar. Add the term and cycle number (for example "Fall Cycle 1").' };
  }

  // No date: needs at least a cycle number, and a term unless only one
  // matching cycle is running.
  if (!parsed.cycle) {
    return { instance: null, message: (parsed.term || parsed.termAmbiguous)
      ? 'The Design Cycle line names a term but not which cycle. Add the ' +
        'cycle number (for example "Fall Cycle 1") or the start date.'
      : 'The Design Cycle line doesn\'t say which cycle this is. Add the ' +
        'term and cycle (for example "Fall Cycle 1") or the start date.' };
  }
  var running = pool.filter(function (i) {
    return i.start <= today && today <= new Date(i.end.getTime() + 7 * DAY_MS_17);
  });
  if (running.length === 1) return { instance: running[0] };
  if (running.length > 1) {
    return { instance: null, message:
      (parsed.termAmbiguous ? '"S" could mean Spring or Summer, and ' : '') +
      'more than one matching cycle is running now: ' +
      running.map(label).join('; ') +
      '. Write the term in full (for example "Spring Cycle 1") or add the ' +
      'start date.' };
  }
  // Nothing running: take the closest one in time, but say so.
  pool.sort(function (a, b) {
    return distance17(a, today) - distance17(b, today);
  });
  var terms = {};
  pool.forEach(function (i) { terms[i.term] = true; });
  if (Object.keys(terms).length > 1) {
    return { instance: null, message: (parsed.termAmbiguous
        ? '"S" could mean Spring or Summer, and both have a Cycle ' +
          parsed.cycle + ' near today. Write "Sp" or "Su" (for example ' +
          '"Sp' + parsed.cycle + '") or add the start date.'
        : null) ||
      'The Design Cycle line doesn\'t say the term, and more than one ' +
      'term has a Cycle ' + parsed.cycle + ' near today. Add the term ' +
      '(for example "Fall Cycle ' + parsed.cycle + '") or the start date.' };
  }
  return { instance: pool[0], note:
    'No date was on the Design Cycle line, so I used the ' + pool[0].term +
    ' Cycle ' + pool[0].cycle + ' closest to today. Add the start date to ' +
    'the line if that\'s the wrong year.' };
}

function matchesStart17(d, start) {
  var years = d.year ? [d.year]
    : [start.getFullYear() - 1, start.getFullYear(), start.getFullYear() + 1];
  return years.some(function (y) {
    var cand = new Date(y, d.month - 1, d.day);
    return Math.abs(dayDiff17(cand, start)) <= START_DATE_TOLERANCE_DAYS_17;
  });
}

function distance17(inst, today) {
  if (today < inst.start) return inst.start - today;
  if (today > inst.end) return today - inst.end;
  return 0;
}

/**
 * Where today falls in a cycle instance.
 * @returns {{kind:'before'|'during'|'between'|'after', current?, previous?,
 *            next?, daysUntil?}}
 */
function describeToday17(inst, today) {
  var weeks = inst.weeks;
  if (dayDiff17(today, weeks[0].start) < 0) {
    return { kind: 'before', next: weeks[0],
             daysUntil: -dayDiff17(today, weeks[0].start) };
  }
  for (var i = 0; i < weeks.length; i++) {
    var w = weeks[i];
    if (dayDiff17(today, w.start) >= 0 && dayDiff17(today, w.end) <= 0) {
      return { kind: 'during', current: w, next: weeks[i + 1] || null };
    }
    var nxt = weeks[i + 1];
    if (nxt && dayDiff17(today, w.end) > 0 && dayDiff17(today, nxt.start) < 0) {
      return { kind: 'between', previous: w, next: nxt };
    }
  }
  return { kind: 'after', previous: weeks[weeks.length - 1] };
}

/** Whole calendar days from b to a (a - b), ignoring time of day. */
function dayDiff17(a, b) {
  var ua = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  var ub = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((ua - ub) / DAY_MS_17);
}

function startOfDay17(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function shortDate17(d) {
  var names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return names[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}


// ── FORMATTERS (use Utilities) ───────────────────────────────

function fmt17(d, tz) {
  return Utilities.formatDate(d, tz, 'MMM d');
}

function formatStatus17(st, tz) {
  var wk = function (w) {
    return 'Week ' + w.week + (w.phase ? ' (' + w.phase + ')' : '');
  };
  switch (st.kind) {
    case 'before':
      return { headline: 'Starts in ' + st.daysUntil + ' day' +
                 (st.daysUntil === 1 ? '' : 's'),
               detail: wk(st.next) + ' begins ' + fmt17(st.next.start, tz) + '.' };
    case 'during':
      return { headline: wk(st.current),
               detail: fmt17(st.current.start, tz) + ' – ' +
                       fmt17(st.current.end, tz) +
                       (st.next ? '. Next: ' + wk(st.next) + ', starting ' +
                                  fmt17(st.next.start, tz) + '.'
                                : '. This is the last week of the cycle.') };
    case 'between':
      return { headline: 'Between Week ' + st.previous.week + ' and Week ' +
                         st.next.week,
               detail: wk(st.previous) + ' ended ' + fmt17(st.previous.end, tz) +
                       '. ' + wk(st.next) + ' starts ' +
                       fmt17(st.next.start, tz) + '.' };
    default:
      return { headline: 'Cycle finished',
               detail: 'The last week (' + wk(st.previous) + ') ended ' +
                       fmt17(st.previous.end, tz) + '.' };
  }
}

function summariseParsed17(p, tz) {
  var bits = [];
  bits.push('Term: ' + (p.term || (p.termAmbiguous ? 'unclear ("S")' : 'not given')));
  bits.push('Cycle: ' + (p.cycle || 'not given'));
  if (p.termYear) bits.push('Year: ' + p.termYear);
  bits.push('Start date: ' + (p.date
    ? p.date.month + '/' + p.date.day + (p.date.year ? '/' + p.date.year : '')
    : 'not given'));
  return bits.join(' · ');
}
