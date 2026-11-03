// ============================================================
// Blueprint Tools — Code10.gs
// Link Project Documents
// Sets up a new project from the blueprint: (A) links the project documents
// found in the blueprint's Drive folder onto the Dashboard and into the
// Development tab's Instructor Resources, (B) fills the course identity
// (department, number, name) into the title of each project document, and
// (C) places the catalog description in the Design tab.
// ------------------------------------------------------------
// Last updated on 2026-11-03 at 01:38 MST
// ------------------------------------------------------------
//
// Spec: docs/dev-handoff/HANDOFF.md section 3.1. Mockup: mockups/LinkDocs.png.
// UI: Sidebar10.html.
//
// CONTRACT. Nothing is written until the user has pressed "Preview all changes"
// and then "Apply". Apply re-scans every document (nothing is carried over from
// the preview except the user's choices), applies only what is ticked, and never
// replaces a link that is already there unless "Replace links" is ticked.
//
// GAS FLAT NAMESPACE. Every name in this file ends in 10 (or 10_ for private
// helpers). Reuses, does NOT redefine:
//   collectTabs                                   (Code.gs)
//   extractDocId                                  (Code2.gs)
//   normTitle9_                                   (Code9.gs)
//
// FLAGS. Places where the handoff left a question open are marked "DEFAULT:" in
// a comment. Grep for DEFAULT: to find every one.
//
// SCOPES. Uses DriveApp (needs https://www.googleapis.com/auth/drive, added in
// its own commit with Greg's OK) and DocumentApp.openById for the other project
// documents (existing "documents" scope; Greg must have edit access to them).
// ============================================================


// ── CONSTANTS ────────────────────────────────────────────────

var MIME_DOC_10    = 'application/vnd.google-apps.document';
var MIME_SHEET_10  = 'application/vnd.google-apps.spreadsheet';
var MIME_DOCX_10   = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
var MIME_SHORTCUT_10 = 'application/vnd.google-apps.shortcut';

// The six documents the identity step can touch. `label` is the sidebar name;
// `nameRe` finds the file in the blueprint's folder; `dash` is the Dashboard
// "Course Documents" label to link (a regex string for findText, anchored to
// the start of the paragraph, so "Syllabus" cannot match mid-sentence).
// `optional` slots may legitimately have no file.
var LINKDOCS_SLOTS_10 = [
  { key: 'schedule',    label: 'Agenda & Schedule of Deliverables', short: 'Schedule of Deliverables',
    type: 'doc',   nameRe: /agenda|schedule\s*of\s*deliverables/i,
    dash: '^\\s*(?:Schedule of Deliverables|Agenda)', optional: false },
  { key: 'settings',    label: 'Settings Tables', short: 'Settings Tables',
    type: 'sheet', nameRe: /settings?\s*tables?/i,
    dash: '^\\s*Settings Tables?', optional: false, noIdentity: true },
  { key: 'syllabus',    label: 'Syllabus', short: 'Syllabus',
    type: 'doc',   nameRe: /syllabus/i,
    dash: '^\\s*Syllabus', optional: false },
  { key: 'guide',       label: 'Instructor Guide', short: 'Instructor Guide',
    type: 'doc',   nameRe: /instructor\s*guide/i,
    dash: '^\\s*Instructor Guide', optional: false },
  { key: 'suggestions', label: 'Course Suggestions', short: 'Course Suggestions',
    type: 'doc',   nameRe: /course\s*suggestions?/i,
    dash: '^\\s*Course Suggestions?', optional: true },
  { key: 'eopr',        label: 'End-of-Project Report', short: 'End-of-Project Report',
    type: 'doc',   nameRe: /end[\s-]*of[\s-]*project|project\s*report/i,
    dash: '^\\s*End[\\s-]*of[\\s-]*Project Report', optional: true }
];

// Order shown in "3. Documents to update": the blueprint first, then these.
var IDENTITY_ORDER_10 = ['blueprint', 'schedule', 'suggestions', 'guide', 'syllabus', 'eopr'];

// A department code followed by a course number: "APPAI 505", "CS 321", "ENGL101".
var IDENT_RE_10 = /\b([A-Z]{2,8})\s*(\d{3}[A-Z]?)\b/;
// Template placeholders: "COURSE ###", "Course ###", "COURSE XXX", and "DEPT ###"
// (Schedule of Deliverables Template B: title "DEPT ###: Course Title" and a
// "DEPT ### Agreement" bullet).
var PLACEHOLDER_RE_10 = /\b(?:COURSE|Course|DEPT)\s*(?:###|XXX)/;
// RE2 versions for Body.findText (it takes a string pattern, case-sensitive).
var PLACEHOLDER_FIND_10 = '\\b(?:COURSE|Course|DEPT)\\s*(?:###|XXX)';

var MAX_OTHERS_SHOWN_10 = 12;


// ── MENU ENTRY ───────────────────────────────────────────────

function showLinkDocsSidebar10() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar10')
    .setTitle('Link Project Documents')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}


// ── SIDEBAR API ──────────────────────────────────────────────

/**
 * Everything the sidebar needs on open: the identity already on the Dashboard
 * (so a re-run is prefilled), the blueprint's folder with its files, the best
 * file guess per slot, and what the Dashboard already links.
 */
function getLinkDocsInit10() {
  var doc = DocumentApp.getActiveDocument();
  var out = {
    blueprintId: doc.getId(),
    blueprintName: doc.getName(),
    identity: { dept: '', number: '', name: '' },
    folder: null,
    folderError: null,
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
  try {
    var t = findDashboardTitle10_(doc);
    if (t) {
      var m = IDENT_RE_10.exec(t.text);
      if (m && m.index <= 40) {
        out.identity.dept   = m[1];
        out.identity.number = m[2];
        var rest = t.text.slice(m.index + m[0].length).replace(/^[\s:\-–—]+/, '').trim();
        out.identity.name = /^(?:Title|Course Title)$/i.test(rest) ? '' : rest;
      }
    }
  } catch (e) { /* prefill is a convenience; never block the sidebar on it */ }
  try {
    out.folder = loadFolderBundle10_(doc, null);
  } catch (e) {
    out.folderError = String(e && e.message || e);
  }
  return out;
}

/** "Pick a different folder": the user pastes a folder URL or ID. */
function loadLinkDocsFolder10(urlOrId) {
  var doc = DocumentApp.getActiveDocument();
  var id = idFromUrl10_(urlOrId);
  if (!id) throw new Error('That does not look like a Drive folder link.');
  return loadFolderBundle10_(doc, id);
}

/**
 * Builds the plan the user reviews. Reads every document, writes nothing.
 * `state` is built by Sidebar10.html (see collectState in that file).
 */
function previewLinkDocs10(state) {
  var plan = buildPlan10_(state, false);
  return pub10_({
    docs: plan.docs,
    catalog: plan.catalog,
    links: plan.links,
    devLinks: plan.devLinks
  });
}

/**
 * Applies what the user ticked. Re-scans first, so a change made between
 * Preview and Apply cannot be applied to the wrong place.
 */
function applyLinkDocs10(state) {
  var plan = buildPlan10_(state, true);
  var report = [];
  var doc = DocumentApp.getActiveDocument();

  // 1. Identity (titles) first, so the Dev-tab headings read right before
  //    the link step looks for them.
  plan.docs.forEach(function (d) {
    if (d.status !== 'ok') return;
    var sel = (state.docs && state.docs[d.key]) || {};
    var did = 0, ok = true;
    try {
      // Others and samples first, titles last (see swapRange10_ note).
      ['other', 'placeholder', 'sample', 'title'].forEach(function (kind) {
        var on = kind === 'title' ? sel.titleOn : (kind === 'sample' ? sel.sampleOn : (kind === 'placeholder' ? sel.phOn : sel.othersOn));
        if (!on) return;
        for (var i = d._items.length - 1; i >= 0; i--) {
          var it = d._items[i];
          if (it.kind !== kind || it.noop) continue;
          it._run();
          did++;
        }
      });
      if (d._doc && !d._isActive) d._doc.saveAndClose();
    } catch (e) {
      ok = false;
      report.push({ ok: false, text: d.label + ': stopped with an error. ' + (e && e.message || e) });
    }
    if (ok) report.push({ ok: true, text: d.label + ': ' + did + ' change' + (did === 1 ? '' : 's') + ' made.' });
  });

  // 2. Catalog description.
  if (plan.catalog && state.catalog && state.catalog.on) {
    try {
      var c = applyCatalog10_(doc, plan.catalog, state.catalog.text, state.catalog.prereqOn ? state.catalog.prereq : '');
      report.push({ ok: c.ok, text: 'Catalog description: ' + c.text });
    } catch (e) {
      report.push({ ok: false, text: 'Catalog description: error. ' + (e && e.message || e) });
    }
  }

  // 3. Links. Re-scanned after the identity step because it renames headings.
  var linkPlan = buildLinkPlan10_(state, doc);
  linkPlan.dash.forEach(function (l) {
    var doIt = l.status === 'add' || (l.status === 'replace' && state.replaceLinks);
    if (!doIt) return;
    try {
      l._el.setLinkUrl(l._s, l._e, l.url);
      report.push({ ok: true, text: 'Dashboard: linked “' + l.label + '”.' });
    } catch (e) {
      report.push({ ok: false, text: 'Dashboard: could not link “' + l.label + '”. ' + (e && e.message || e) });
    }
  });
  if (state.devLinks) {
    linkPlan.dev.forEach(function (l) {
      var doIt = l.status === 'add' || (l.status === 'replace' && state.replaceLinks);
      if (!doIt) return;
      try {
        l._el.setLinkUrl(l._s, l._e, l.url);
        report.push({ ok: true, text: 'Development tab: linked “' + l.label + '”.' });
      } catch (e) {
        report.push({ ok: false, text: 'Development tab: could not link “' + l.label + '”. ' + (e && e.message || e) });
      }
    });
  }
  return { report: report };
}


// ── PLAN BUILDER ─────────────────────────────────────────────

function buildPlan10_(state, forApply) {
  var id = cleanIdentity10_(state.identity);
  var active = DocumentApp.getActiveDocument();
  var docs = [];

  IDENTITY_ORDER_10.forEach(function (key) {
    var label = key === 'blueprint' ? 'Blueprint' : slotByKey10_(key).label;
    var d = { key: key, label: label, status: 'ok', message: '', items: [], _items: [], _doc: null, _isActive: false };
    try {
      var target = null;
      if (key === 'blueprint') {
        target = active; d._isActive = true;
      } else {
        var sel = (state.slots && state.slots[key]) || {};
        if (sel.skip) { d.status = 'skipped'; d.message = 'Skipped.'; docs.push(d); return; }
        var docId = docIdForSlot10_(sel);
        if (!docId) {
          d.status = 'skipped';
          d.message = 'No document chosen. Skipped. Paste a URL in section 4 if there is one.';
          docs.push(d); return;
        }
        // A Word (.docx) file cannot be opened with DocumentApp. Say so plainly.
        var mimeOf = '';
        try { mimeOf = DriveApp.getFileById(docId).getMimeType(); } catch (e0) { /* fall through to the open error below */ }
        if (mimeOf && mimeOf !== MIME_DOC_10) {
          d.status = 'skipped';
          d.message = mimeOf === MIME_DOCX_10
            ? 'This is a Word (.docx) file, which the tool cannot edit. It can still be linked. To fill in the title, open it in Google Docs (File > Save as Google Docs) and run the tool again, or edit it by hand.'
            : 'This file is not a Google Doc, so the tool cannot edit it.';
          docs.push(d); return;
        }
        try {
          target = DocumentApp.openById(docId);
        } catch (e) {
          d.status = 'error';
          d.message = 'Could not open this document. You need edit access to it.';
          docs.push(d); return;
        }
      }
      d._doc = target;
      var scan = scanDoc10_(target, key, id);
      d._items = scan.items;
      d.items = scan.items;
      d.notes = scan.notes;
      if (!scan.items.some(function (i) { return i.kind === 'title'; })) {
        d.status = 'notitle';
        d.message = key === 'eopr'
          ? 'No “COURSE XXX” found in any tab. Nothing to change.'
          : 'No title line recognized. Nothing to change.';
      }
    } catch (e) {
      d.status = 'error';
      d.message = String(e && e.message || e);
    }
    docs.push(d);
  });

  var catalog = null;
  if (state.catalog) catalog = planCatalog10_(active, state.catalog);

  var linkPlan = buildLinkPlan10_(state, active);

  return {
    docs: docs,
    catalog: catalog,
    links: linkPlan.dash,
    devLinks: linkPlan.dev
  };
}

function slotByKey10_(key) {
  for (var i = 0; i < LINKDOCS_SLOTS_10.length; i++) if (LINKDOCS_SLOTS_10[i].key === key) return LINKDOCS_SLOTS_10[i];
  return null;
}

function cleanIdentity10_(raw) {
  raw = raw || {};
  var dept = String(raw.dept || '').trim().toUpperCase();
  var num  = String(raw.number || '').trim().toUpperCase();
  var name = String(raw.name || '').replace(/\s+/g, ' ').trim();
  if (!dept || !num) throw new Error('Enter the department and course number first.');
  return { dept: dept, number: num, name: name, id: dept + ' ' + num };
}

function docIdForSlot10_(sel) {
  if (sel.url && String(sel.url).trim()) return extractDocId(String(sel.url).trim());
  return sel.fileId || null;
}


// ── IDENTITY: SCANNING ───────────────────────────────────────

/**
 * Finds every place in `doc` that the identity step would touch.
 * Returns {items, notes}. Each item:
 *   kind     'title' | 'sample' | 'other'
 *   where    short plain-language location for the preview
 *   before / after   short text snippets for the preview
 *   noop     true when the text already says the right thing
 *   _run     function that applies the change (never sent to the sidebar)
 *
 * DEFAULT (handoff open question 2): in a title line that already has a course
 * code, ONLY the code is replaced. The rest of the line (name, "(Incomplete)",
 * "Syllabus") is left alone. If the name on the line differs from the Course
 * name field the item carries `nameDiffers` and the sidebar warns.
 */
function scanDoc10_(doc, slotKey, id) {
  var tabs = collectTabs(doc);
  var items = [], notes = [], titleSpans = [];   // titleSpans: {ti, ptext, s} to keep out of "others"
  var oldIds = {};

  function addTitleItem(ti, textEl, where) {
    var text = textEl.getText();
    var ed = titleEdit10_(text, id);
    if (!ed) return false;
    if (ed.old && ed.mode === 'replace') oldIds[ed.old] = true;
    titleSpans.push({ ti: ti, ptext: text, s: ed.start });
    items.push({
      kind: 'title', where: where,
      before: snip10_(text), after: snip10_(ed.after),
      mode: ed.mode, old: ed.old, nameDiffers: ed.nameDiffers,
      noop: ed.ops.length === 0,
      _run: function () { ed.ops.forEach(function (op) { swapRange10_(textEl, op.s, op.e, op.repl); }); }
    });
    return true;
  }

  if (slotKey === 'eopr') {
    // DEFAULT: dept + number only, no name. Case-sensitive "COURSE XXX", every tab.
    tabs.forEach(function (t, ti) {
      findAll10_(t.body, 'COURSE XXX').forEach(function (r) {
        var el = r.getElement().asText(), s = r.getStartOffset(), e = r.getEndOffsetInclusive();
        var text = el.getText();
        items.push({
          kind: 'title', where: 'Tab “' + t.title + '”',
          before: snip10_(text), after: snip10_(text.slice(0, s) + id.id + text.slice(e + 1)),
          mode: 'placeholder', noop: false,
          _run: function () { swapRange10_(el, s, e, id.id); }
        });
      });
    });
    return { items: items, notes: notes };
  }

  if (slotKey === 'blueprint') {
    // Dashboard title (Title-styled paragraph).
    var dt = findDashboardTitle10_(doc);
    if (dt) addTitleItem(dt.ti, dt.el, 'Dashboard title');
    else notes.push('No Dashboard title line found.');
    // Development tab: the two Instructor Resources headings.
    tabs.forEach(function (t, ti) {
      if (!/\bdevelopment\b/i.test(t.title)) return;
      topParas10_(t.body).forEach(function (p) {
        if (!/^\s*(?:COURSE|Course|[A-Z]{2,8})\s*(?:###|XXX|\d{3}[A-Z]?)\s+(?:Instructor Guide|Course Suggestions)\b/.test(p.text)) return;
        addTitleItem(ti, p.el, 'Development tab heading');
      });
    });
  } else {
    // Other project documents: the title line is the first short paragraph near
    // the top whose shape matches (placeholder or code-and-number).
    // DEFAULT: the Agenda and Instructor Guide placeholders are unconfirmed
    // (Greg is sending unfilled copies). Until then only the shapes above match;
    // anything else is reported as "No title line recognized", never guessed.
    var first = tabs[0], found = false;
    var paras = topParas10_(first.body).slice(0, 20);
    for (var i = 0; i < paras.length && !found; i++) {
      if (!paras[i].text.trim() || paras[i].text.length > 160) continue;
      found = addTitleItem(0, paras[i].el, slotKey === 'syllabus' ? 'Title' :
        (slotKey === 'suggestions' ? 'Heading' : 'Title line'));
    }
    // Schedule of Deliverables Templates A and C keep the course identity in a
    // Project Information table: empty cells beside "Course Prefix/Number:" and
    // "Course Title:". Template B has a title line (handled above); Template D
    // has neither, so it reports "No title line recognized".
    if (slotKey === 'schedule') scanScheduleTable10_(first.body, id, items);
    if (slotKey === 'guide' && id.name) {
      tabs.forEach(function (t, ti) {
        findAll10_(t.body, 'Welcome to Course Name').forEach(function (r) {
          var el = r.getElement().asText();
          var s = r.getStartOffset() + 'Welcome to '.length, e = r.getEndOffsetInclusive();
          var text = el.getText();
          items.push({
            kind: 'sample', where: 'Sample announcement',
            before: snip10_(text), after: snip10_(text.slice(0, s) + id.name + text.slice(e + 1)),
            noop: false,
            _run: function () { swapRange10_(el, s, e, id.name); }
          });
        });
      });
    }
  }

  // "Other matches": the OLD code (or a leftover placeholder) found anywhere
  // else. Never replaced unless the user ticks the box. DEFAULT: unchecked.
  var olds = Object.keys(oldIds).filter(function (o) { return o !== id.id; });
  var pat = '(?:' + olds.map(escRe10_).concat([PLACEHOLDER_FIND_10]).join('|') + ')';
  tabs.forEach(function (t, ti) {
    findAll10_(t.body, pat).forEach(function (r) {
      var el = r.getElement().asText(), s = r.getStartOffset(), e = r.getEndOffsetInclusive();
      var text = el.getText();
      var isTitle = titleSpans.some(function (ts) { return ts.ti === ti && ts.ptext === text && ts.s === s; });
      if (isTitle) return;
      // DEFAULT: a leftover placeholder ("DEPT ###", "COURSE ###") is NOT an old
      // title, so it gets its own group, checked by default (Schedule Template E
      // has about a dozen "DEPT ### <document>" link labels). Old codes stay
      // unchecked ("other"), per the handoff.
      var isPh = PLACEHOLDER_RE_10.test(text.substring(s, e + 1));
      items.push({
        kind: isPh ? 'placeholder' : 'other', where: 'Tab “' + t.title + '”',
        before: snip10_(text, s), after: id.id, noop: false,
        _run: function () { swapRange10_(el, s, e, id.id); }
      });
    });
  });

  return { items: items, notes: notes };
}

/**
 * Fills the empty cell beside "Course Prefix/Number:" with "DEPT ###" and the
 * one beside "Course Title:" with the course name. A cell that already holds a
 * code gets only the code swapped; a title cell that is already filled is left
 * alone. DEFAULT: credits, people and dates are never touched.
 */
function scanScheduleTable10_(body, id, items) {
  body.getTables().forEach(function (tbl) {
    for (var r = 0; r < tbl.getNumRows(); r++) {
      var row = tbl.getRow(r);
      if (row.getNumCells() < 2) continue;
      var label = row.getCell(0).getText().replace(/\s+/g, ' ').trim();
      var isNum = /^Course Prefix\s*\/\s*Number:?$/i.test(label);
      var isTitle = /^Course Title:?$/i.test(label);
      if (!isNum && !isTitle) continue;
      var want = isNum ? id.id : id.name;
      if (!want) continue;
      var cell = row.getCell(1), cur = cell.getText().trim();
      var where = 'Project Information › ' + (isNum ? 'Course Prefix/Number' : 'Course Title');
      if (!cur) {
        items.push({
          kind: 'title', where: where, before: '(empty)', after: want, mode: 'placeholder', noop: false,
          _run: (function (c, w) { return function () {
            var first = c.getNumChildren() ? c.getChild(0) : null;
            if (first && first.getType() === DocumentApp.ElementType.PARAGRAPH) first.asParagraph().appendText(w);
            else c.appendParagraph(w);
          }; })(cell, want)
        });
      } else if (isNum && cell.getNumChildren() === 1 && cell.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
        var ed = titleEdit10_(cur, id);
        if (ed) {
          var te = cell.getChild(0).asParagraph().editAsText();
          items.push({
            kind: 'title', where: where, before: snip10_(cur), after: snip10_(ed.after), mode: ed.mode, old: ed.old,
            noop: ed.ops.length === 0,
            _run: function () { ed.ops.forEach(function (op) { swapRange10_(te, op.s, op.e, op.repl); }); }
          });
        }
      } else if (isTitle) {
        items.push({ kind: 'title', where: where, before: snip10_(cur), after: snip10_(cur), mode: 'same', noop: true, _run: function () {} });
      }
    }
  });
}

/**
 * Works out how to turn a title line into one that carries the new identity.
 * Returns null when the line does not look like a title. Otherwise:
 *   mode  'placeholder' | 'replace' | 'same'
 *   ops   [{s, e, repl}] in DESCENDING offset order, so applying them one after
 *         another never shifts the next one's position
 *   after the resulting text; start where the code/placeholder begins
 */
function titleEdit10_(text, id) {
  var ops = [], m, mode, old, start;
  m = PLACEHOLDER_RE_10.exec(text);
  if (m) {
    mode = 'placeholder'; old = m[0]; start = m.index;
    // "COURSE ###: Title" / "Course ###: Course Title": the name goes in too.
    var tail = /:\s*(?:\[?Course (?:Title|Name)\]?|Title)\s*$/.exec(text);
    if (tail && id.name && tail.index > m.index) {
      ops.push({ s: tail.index, e: tail.index + tail[0].length - 1, repl: ': ' + id.name });
    }
    ops.push({ s: m.index, e: m.index + m[0].length - 1, repl: id.id });
  } else {
    m = IDENT_RE_10.exec(text);
    if (!m || m.index > 40) return null;
    old = m[0]; start = m.index;
    var same = (m[1] + m[2]).toUpperCase() === (id.dept + id.number).replace(/\s+/g, '');
    if (same) {
      mode = 'same';
    } else {
      mode = 'replace';
      ops.push({ s: m.index, e: m.index + m[0].length - 1, repl: id.id });
    }
  }
  var after = text;
  ops.forEach(function (op) { after = after.slice(0, op.s) + op.repl + after.slice(op.e + 1); });
  var nameDiffers = false;
  if (mode !== 'placeholder' && id.name) {
    nameDiffers = normTitle9_(text).indexOf(normTitle9_(id.name)) === -1;
  }
  return { mode: mode, old: old, ops: ops, after: after, start: start, nameDiffers: nameDiffers };
}

/**
 * Replaces text[s..e] (inclusive) with `repl`, keeping the old text's
 * formatting: the new text goes in AFTER the old run (where it inherits the
 * old run's look), then the old run is deleted.
 */
function swapRange10_(textEl, s, e, repl) {
  if (repl) textEl.insertText(e + 1, repl);
  textEl.deleteText(s, e);
}


// ── DOCUMENT WALKING HELPERS ─────────────────────────────────

/** Top-level paragraphs and list items of a body: [{el (Text), text, idx}]. */
function topParas10_(body) {
  var out = [], n = body.getNumChildren();
  for (var i = 0; i < n; i++) {
    var c = body.getChild(i), ty = c.getType();
    if (ty === DocumentApp.ElementType.PARAGRAPH) {
      out.push({ el: c.asParagraph().editAsText(), text: c.asParagraph().getText(), idx: i, para: c.asParagraph() });
    } else if (ty === DocumentApp.ElementType.LIST_ITEM) {
      out.push({ el: c.asListItem().editAsText(), text: c.asListItem().getText(), idx: i, para: c.asListItem() });
    }
  }
  return out;
}

/** All matches of a RE2 pattern in a body (descends into tables). */
function findAll10_(body, pattern) {
  var out = [], r = body.findText(pattern);
  while (r) {
    if (r.getElement().getType() === DocumentApp.ElementType.TEXT) out.push(r);
    r = body.findText(pattern, r);
  }
  return out;
}

function escRe10_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function snip10_(text, around) {
  text = String(text || '');
  if (around !== undefined) {
    var a = Math.max(0, around - 30);
    return (a > 0 ? '…' : '') + text.slice(a, around + 60) + (text.length > around + 60 ? '…' : '');
  }
  return text.length > 90 ? text.slice(0, 90) + '…' : text;
}

/** Drops every key starting with "_" (closures, Elements) so the result can go to the sidebar. */
function pub10_(x) {
  return JSON.parse(JSON.stringify(x, function (k, v) { return k.charAt(0) === '_' ? undefined : v; }));
}

function idFromUrl10_(urlOrId) {
  var s = String(urlOrId || '').trim();
  var m = s.match(/\/folders\/([a-zA-Z0-9_-]{15,})/) || s.match(/\/d\/([a-zA-Z0-9_-]{15,})/) || s.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
  if (m) return m[1];
  return /^[a-zA-Z0-9_-]{15,}$/.test(s) ? s : null;
}

function sameUrl10_(a, b) {
  if (!a || !b) return false;
  var ia = idFromUrl10_(a), ib = idFromUrl10_(b);
  return ia && ib ? ia === ib : String(a).trim() === String(b).trim();
}


// ── DASHBOARD ────────────────────────────────────────────────

function dashboardTab10_(doc) {
  var tabs = collectTabs(doc);
  for (var i = 0; i < tabs.length; i++) if (/dashboard/i.test(tabs[i].title)) return { ti: i, body: tabs[i].body };
  return { ti: 0, body: tabs[0].body };
}

/** The Dashboard's Title-styled paragraph: {ti, el (Text), text} or null. */
function findDashboardTitle10_(doc) {
  var d = dashboardTab10_(doc);
  var paras = topParas10_(d.body).slice(0, 15), pick = null;
  for (var i = 0; i < paras.length; i++) {
    if (paras[i].para.getHeading && paras[i].para.getHeading() === DocumentApp.ParagraphHeading.TITLE && paras[i].text.trim()) { pick = paras[i]; break; }
  }
  if (!pick) pick = paras.filter(function (p) { return p.text.trim(); })[0];
  return pick ? { ti: d.ti, el: pick.el, text: pick.text } : null;
}

/** Finds a Dashboard label by regex: {_el, _s, _e, existing} or null. */
function findDashLabel10_(dashBody, re) {
  var r = dashBody.findText(re);
  while (r) {
    if (r.getElement().getType() === DocumentApp.ElementType.TEXT) {
      var el = r.getElement().asText(), s = r.getStartOffset(), e = r.getEndOffsetInclusive();
      var seg = el.getText().substring(s, e + 1);
      s += seg.match(/^\s*/)[0].length;
      return { _el: el, _s: s, _e: e, existing: el.getLinkUrl(s) || null };
    }
    r = dashBody.findText(re, r);
  }
  return null;
}


// ── DRIVE FOLDER ─────────────────────────────────────────────

/**
 * The folder, its files, and the best file per slot. `folderId` null means
 * "the blueprint's own parent folder".
 * DEFAULT: Drive scope is the full "drive" scope (Greg chose option b).
 */
function loadFolderBundle10_(doc, folderId) {
  var folder;
  if (folderId) {
    folder = DriveApp.getFolderById(folderId);
  } else {
    var parents = DriveApp.getFileById(doc.getId()).getParents();
    if (!parents.hasNext()) throw new Error('The blueprint is not inside a Drive folder. Use “Pick a different folder”.');
    folder = parents.next();
  }
  var files = [], it = folder.getFiles(), n = 0;
  while (it.hasNext() && n < 300) {
    var f = it.next(); n++;
    var mime = f.getMimeType(), fid = f.getId(), shortcut = false;
    // New project folders hold SHORTCUTS to the Instructor Guide and Course
    // Suggestions (the real documents sit beside the project folder until they
    // are moved to the program folder). Resolve each shortcut to its target, so
    // the tool links and edits the real document, but keep the shortcut's name,
    // which carries the project name.
    if (mime === MIME_SHORTCUT_10) {
      var tid = f.getTargetId();
      if (!tid) continue;
      fid = tid; mime = f.getTargetMimeType(); shortcut = true;
    }
    if (fid === doc.getId()) continue;                     // never offer the bound blueprint itself
    var type = mime === MIME_DOC_10 ? 'doc' : (mime === MIME_SHEET_10 ? 'sheet' : (mime === MIME_DOCX_10 ? 'docx' : 'other'));
    var url = type === 'doc' ? 'https://docs.google.com/document/d/' + fid + '/edit'
            : (type === 'sheet' ? 'https://docs.google.com/spreadsheets/d/' + fid + '/edit' : f.getUrl());
    files.push({ id: fid, name: f.getName(), type: type, shortcut: shortcut, url: url });
  }
  files.sort(function (a, b) { return a.name < b.name ? -1 : 1; });

  var dash = dashboardTab10_(doc);
  var slots = LINKDOCS_SLOTS_10.map(function (s) {
    // Ignore a leading "Copy of " when matching. The type filter is what keeps
    // the generated "Video List - <blueprint title>" sheet from matching.
    // A Word (.docx) file counts as a "doc": it can be linked, but not edited
    // (see buildPlan10_).
    var cands = files.filter(function (f) { return f.type === s.type || (s.type === 'doc' && f.type === 'docx'); });
    var hits = cands.filter(function (f) { return s.nameRe.test(f.name.replace(/^copy of\s+/i, '')); });
    var lab = findDashLabel10_(dash.body, s.dash);
    return {
      key: s.key, label: s.label, short: s.short, type: s.type, optional: s.optional,
      options: cands.map(function (f) { return { id: f.id, name: f.name, note: f.type === 'docx' ? 'Word file' : (f.shortcut ? 'shortcut' : '') }; }),
      guess: hits.length ? hits[0].id : null,
      matches: hits.length,
      hasLabel: !!lab,
      existing: lab ? lab.existing : null
    };
  });
  var folderLab = findDashLabel10_(dash.body, '^\\s*Google Folder');
  return {
    id: folder.getId(), name: folder.getName(), url: folder.getUrl(),
    count: files.length, slots: slots,
    folderHasLabel: !!folderLab, folderExisting: folderLab ? folderLab.existing : null
  };
}


// ── LINKS ────────────────────────────────────────────────────

/**
 * What would be linked where. Returns {dash:[...], dev:[...]}; each entry has
 * label, url, status ('add' | 'same' | 'replace' | 'nolabel' | 'none'), and
 * the private _el/_s/_e that Apply uses.
 */
function buildLinkPlan10_(state, doc) {
  var dash = [], dev = [];
  var dashTab = dashboardTab10_(doc);
  var urls = {};

  LINKDOCS_SLOTS_10.forEach(function (s) {
    var sel = (state.slots && state.slots[s.key]) || {};
    var url = null;
    if (!sel.skip) url = urlForSlot10_(sel);
    urls[s.key] = url;
    if (!url) return;
    var lab = findDashLabel10_(dashTab.body, s.dash);
    dash.push(linkEntry10_(s.short, url, lab));
  });

  // Google Folder: the folder itself (DEFAULT: added automatically).
  if (state.folderUrl && /^https:\/\/drive\.google\.com\//i.test(state.folderUrl)) {
    dash.push(linkEntry10_('Google Folder', state.folderUrl, findDashLabel10_(dashTab.body, '^\\s*Google Folder')));
  }

  // Development tab › Instructor Resources.
  // DEFAULT: if the section already has a linked "Instructor Guide" /
  // "Course Suggestions" phrase in its body text, that is the link spot (the
  // finished-course form). Otherwise the heading text becomes the link (the
  // template form). Never replaces an existing link unless the user asks.
  var devTab = null;
  collectTabs(doc).forEach(function (t) { if (!devTab && /\bdevelopment\b/i.test(t.title)) devTab = t; });
  if (devTab) {
    [['guide', 'Instructor Guide', /Instructor Guide/i],
     ['suggestions', 'Course Suggestions', /Course Suggestions?/i]].forEach(function (x) {
      var url = urls[x[0]];
      if (!url) return;
      var loc = locateInstructorResource10_(devTab.body, x[2]);
      var e = { label: x[1], url: url, status: 'nolabel', existing: null };
      if (loc) {
        e._el = loc.el; e._s = loc.s; e._e = loc.e; e.existing = loc.existing; e.spot = loc.spot;
        e.status = !loc.existing ? 'add' : (sameUrl10_(loc.existing, url) ? 'same' : 'replace');
      }
      dev.push(e);
    });
  }
  return { dash: dash, dev: dev };
}

function linkEntry10_(label, url, lab) {
  var e = { label: label, url: url, status: 'nolabel', existing: null };
  if (lab) {
    e._el = lab._el; e._s = lab._s; e._e = lab._e; e.existing = lab.existing;
    e.status = !lab.existing ? 'add' : (sameUrl10_(lab.existing, url) ? 'same' : 'replace');
  }
  return e;
}

function urlForSlot10_(sel) {
  if (sel.url && String(sel.url).trim()) {
    var u = String(sel.url).trim();
    return /^https:\/\//i.test(u) ? u : null;       // only https links are ever written
  }
  if (sel.fileId) {
    try { return DriveApp.getFileById(sel.fileId).getUrl(); } catch (e) { return null; }
  }
  return null;
}

/**
 * Finds where the link for one Instructor Resources item goes.
 * Returns {el, s, e, existing, spot:'sentence'|'heading'} or null.
 */
function locateInstructorResource10_(body, phraseRe) {
  var paras = topParas10_(body), i, start = -1, end = paras.length;
  for (i = 0; i < paras.length; i++) {
    if (isHeading10_(paras[i].para, 2) && /^\s*Instructor Resources/i.test(paras[i].text)) { start = i; break; }
  }
  if (start < 0) return null;
  for (i = start + 1; i < paras.length; i++) if (isHeading10_(paras[i].para, 2)) { end = i; break; }

  var h = -1;
  for (i = start + 1; i < end; i++) {
    if (isHeading10_(paras[i].para, 0) && phraseRe.test(paras[i].text)) { h = i; break; }
  }
  if (h < 0) return null;

  // Finished-course form: a linked phrase in the body text under the heading.
  for (i = h + 1; i < end && !isHeading10_(paras[i].para, 0); i++) {
    var t = paras[i].el, txt = paras[i].text, m = phraseRe.exec(txt);
    if (m && t.getLinkUrl(m.index)) {
      var url = t.getLinkUrl(m.index), s = m.index, e = m.index + m[0].length - 1;
      while (s > 0 && t.getLinkUrl(s - 1) === url) s--;
      while (e < txt.length - 1 && t.getLinkUrl(e + 1) === url) e++;
      return { el: t, s: s, e: e, existing: url, spot: 'sentence' };
    }
  }
  // Template form: the heading text itself.
  var ht = paras[h].el, htxt = paras[h].text.replace(/\s+$/, '');
  if (!htxt) return null;
  return { el: ht, s: 0, e: htxt.length - 1, existing: ht.getLinkUrl(0) || null, spot: 'heading' };
}

/** level 0 = any heading; 2 = Heading 2. */
function isHeading10_(para, level) {
  if (!para.getHeading) return false;
  var h = para.getHeading(), H = DocumentApp.ParagraphHeading;
  if (level === 2) return h === H.HEADING2;
  return h !== H.NORMAL;
}


// ── CATALOG LOOKUP (Kuali) ───────────────────────────────────
// Found by Greg in the browser's Network tab (2026-11-02): the catalogs run on
// Kuali, and a search request returns course JSON with no sign-in:
//   https://boisestate.kuali.co/api/v1/catalog/search/<catalogId>?q=cs%20321&limit=6
// Each result has: code ("CS321"), subjectCode.name ("CS"), number ("321"),
// title, description, pid, type ("courses" | "programs" | "experiences").
// Needs only script.external_request, which the project already has.
//
// UNVERIFIED: how the catalog IDs are discovered. Catalog IDs change every
// academic year, so the code asks /api/v1/catalog/public/catalogs for the list
// and picks the newest one whose name says undergraduate or graduate. The shape
// of that list is a guess until spikeCatalogSpike10 (Spikes.gs) has been run.
// If discovery fails it falls back to the IDs below. Update the fallbacks yearly
// until the spike confirms discovery works.
var KUALI_BASE_10 = 'https://boisestate.kuali.co/api/v1/catalog/';
var CATALOG_FALLBACK_IDS_10 = {
  undergrad: '68d5535c7e5bba5930769e2d',   // 2026-2027 Undergraduate Catalog (seen in Greg's Network tab)
  grad: ''                                   // not known yet; filled in after the spike
};

/**
 * Looks one course up in the catalog. `level` is 'undergrad' | 'grad', or empty
 * for the automatic choice (course number 500 and above is graduate).
 * Never throws for "not found"; returns {ok:false, message}.
 * DEFAULT: older catalog years are not searched (they are archived PDFs).
 */
function fetchCatalogCourse10(dept, number, level) {
  dept = String(dept || '').trim().toUpperCase();
  number = String(number || '').trim().toUpperCase();
  if (!dept || !number) return { ok: false, message: 'Enter the department and course number first.' };
  var n = parseInt(number.replace(/\D/g, ''), 10);
  var lvl = (level === 'grad' || level === 'undergrad') ? level : (n >= 500 ? 'grad' : 'undergrad');
  var out = { ok: false, level: lvl, levelName: lvl === 'grad' ? 'Graduate' : 'Undergraduate' };
  try {
    var cat = catalogFor10_(lvl);
    if (!cat.id) { out.message = 'The ' + out.levelName.toLowerCase() + ' catalog could not be found automatically. Paste the description instead.'; return out; }
    out.catalogName = cat.name; out.year = cat.year;
    var url = KUALI_BASE_10 + 'search/' + cat.id + '?q=' + encodeURIComponent(dept + ' ' + number) + '&limit=20';
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (resp.getResponseCode() !== 200) { out.message = 'The catalog answered with error ' + resp.getResponseCode() + '. Paste the description instead.'; return out; }
    var list = JSON.parse(resp.getContentText());
    var want = (dept + number).replace(/[\s/-]+/g, '');
    var hit = null;
    for (var i = 0; i < list.length && !hit; i++) {
      var it = list[i];
      if (it.type !== 'courses') continue;
      var code = String((it.subjectCode && it.subjectCode.name || '') + (it.number || '')).toUpperCase().replace(/[\s/-]+/g, '');
      if (code === want || String(it.code || '').toUpperCase().replace(/[\s/-]+/g, '') === want) hit = it;
    }
    if (!hit) { out.message = dept + ' ' + number + ' was not found in the ' + out.levelName.toLowerCase() + ' catalog.'; return out; }
    out.ok = true;
    out.title = String(hit.title || '').trim();
    out.description = cleanCatalogText10_(hit.description);
    out.code = dept + ' ' + number;
    out.prerequisites = '';
    // The search result has no prerequisites. The course-detail request does
    // (seen in Greg's Network tab: fields prerequisites [HTML], credits.value,
    // description). UNVERIFIED URL shape: .../catalog/course/<catalogId>/<pid>.
    // If it fails, the description is still returned and the user types the
    // prerequisite by hand.
    try {
      var dr = UrlFetchApp.fetch(KUALI_BASE_10 + 'course/' + cat.id + '/' + encodeURIComponent(hit.pid),
                                 { muteHttpExceptions: true, followRedirects: true });
      if (dr.getResponseCode() === 200) {
        var d = JSON.parse(dr.getContentText());
        out.prerequisites = cleanCatalogText10_(String(d.prerequisites || '').replace(/<\/li>/gi, '; ')).replace(/[;\s]+$/, '');
        if (d.description) out.description = cleanCatalogText10_(d.description);
        if (d.credits && d.credits.value) out.credits = String(d.credits.value);
      } else {
        out.detailNote = 'Prerequisites could not be fetched (error ' + dr.getResponseCode() + ').';
      }
    } catch (e2) { out.detailNote = 'Prerequisites could not be fetched.'; }
    return out;
  } catch (e) {
    out.message = 'Could not reach the catalog (' + (e && e.message || e) + '). Paste the description instead.';
    return out;
  }
}

/** Newest catalog whose name says undergraduate (or graduate); falls back to the fixed IDs. */
function catalogFor10_(lvl) {
  try {
    var resp = UrlFetchApp.fetch(KUALI_BASE_10 + 'public/catalogs', { muteHttpExceptions: true, followRedirects: true });
    if (resp.getResponseCode() === 200) {
      var list = JSON.parse(resp.getContentText()), best = null;
      (Array.isArray(list) ? list : []).forEach(function (c) {
        var name = String(c.name || c.title || c.fullName || '');
        var isUg = /\bundergraduate\b/i.test(name), isGr = /\bgraduate\b/i.test(name) && !isUg;
        if ((lvl === 'undergrad' && !isUg) || (lvl === 'grad' && !isGr)) return;
        var ym = name.match(/(\d{4})\s*[-–—]\s*(\d{2,4})/);
        var start = ym ? parseInt(ym[1], 10) : 0;
        if (!best || start > best.start) {
          best = { id: c._id || c.id, name: name, start: start,
                   year: ym ? ym[1] + '–' + ym[2].slice(-2) : '' };
        }
      });
      if (best && best.id) return best;
    }
  } catch (e) { /* fall through to the fixed IDs */ }
  var fid = CATALOG_FALLBACK_IDS_10[lvl];
  return { id: fid, name: '', year: lvl === 'undergrad' ? '2026–27' : '' };
}

/** Strips HTML tags and the trailing newline the catalog leaves on descriptions. */
function cleanCatalogText10_(t) {
  return String(t || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ').trim();
}


// ── CATALOG DESCRIPTION ──────────────────────────────────────

/**
 * Where the description would go. Looks for the template's "Add catalog
 * description here" bullet in the Design tab.
 * DEFAULT: if that bullet is not there (a finished course), nothing is added;
 * the sidebar says so. It never writes over a description that is already in.
 * The text comes from the sidebar: either fetched by fetchCatalogCourse10 (the
 * user pressed "Look up in the catalog") or pasted by the user.
 */
function planCatalog10_(doc, cat) {
  var spot = findCatalogSpot10_(doc);
  return { found: !!spot, _spot: spot,
           text: String(cat.text || '').replace(/\s*[\r\n]+\s*/g, ' ').trim() };
}

function findCatalogSpot10_(doc) {
  var tabs = collectTabs(doc), body = null;
  for (var i = 0; i < tabs.length; i++) if (/^\s*design\s*$/i.test(tabs[i].title)) { body = tabs[i].body; break; }
  if (!body) for (var j = 0; j < tabs.length; j++) if (/design/i.test(tabs[j].title)) { body = tabs[j].body; break; }
  if (!body) return null;
  var r = findAll10_(body, '^\\s*Add catalog description here\\s*$')[0];
  if (!r) return null;
  var el = r.getElement().asText();
  return { el: el, len: el.getText().length };
}

function applyCatalog10_(doc, plan, text, prereq) {
  var clean = String(text || '').replace(/\s*[\r\n]+\s*/g, ' ').trim();
  if (!clean) return { ok: false, text: 'nothing to add (the box was empty).' };
  var spot = findCatalogSpot10_(doc);
  if (!spot) return { ok: false, text: 'the “Add catalog description here” bullet was not found, so nothing was added.' };
  swapRange10_(spot.el, 0, spot.len - 1, clean);
  var msg = 'added to the Design tab.';
  var pre = String(prereq || '').replace(/\s+/g, ' ').trim();
  if (pre) {
    // DEFAULT: the prerequisite goes on its own bullet right below the
    // description, as in finished blueprints ("Prerequisite: ...").
    try { insertBulletAfter10_(spot.el, 'Prerequisite: ' + pre); msg += ' Prerequisite added on the next bullet.'; }
    catch (e) { msg += ' The prerequisite bullet could not be added (' + (e && e.message || e) + ').'; }
  }
  return { ok: true, text: msg };
}

/** Inserts a copy of the paragraph/list item holding `textEl`, right after it, with new text. */
function insertBulletAfter10_(textEl, newText) {
  var para = textEl.getParent();
  var parent = para.getParent();
  var idx = parent.getChildIndex(para);
  var copy = para.copy();
  var ct = copy.editAsText();
  ct.setText(newText);
  var kind = parent.getType();
  var isList = para.getType() === DocumentApp.ElementType.LIST_ITEM;
  var host = kind === DocumentApp.ElementType.BODY_SECTION ? parent.asBody()
           : (kind === DocumentApp.ElementType.TABLE_CELL ? parent.asTableCell() : null);
  if (!host) throw new Error('unexpected location');
  if (isList) host.insertListItem(idx + 1, copy.asListItem());
  else host.insertParagraph(idx + 1, copy.asParagraph());
}
