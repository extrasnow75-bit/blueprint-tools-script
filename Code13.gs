// ============================================================
// Blueprint Tools — Code13.gs
// Development Status Check
// Reads the blueprint and reports where each activity stands: review stage
// (untouched, drafted, green, purple, yellow, struck through), settings-table
// links on tool lines, Home / Syllabus / Instructor Resources, module time vs.
// the credit-hour guideline, settings-table tabs marked finalized, and Media
// Deliverables that are Completed but not yet added to the Dev tab.
// ------------------------------------------------------------
// Last updated on 2026-11-08 at 18:00 MST
// ------------------------------------------------------------
//
// Spec: docs/dev-handoff/HANDOFF.md section 3.4. Mockups: Main.png (setup),
// StatusResultsB.png (results). UI: Sidebar13.html.
//
// READ-ONLY. The check never edits the blueprint or the settings table. In
// particular it NEVER ticks "Check when settings are finalized:" or "Link Added
// to Dev Tab (IDC)". The only writes are the optional "save the report" outputs
// (a new tab in the blueprint, or a new Google Doc in the blueprint's folder),
// and only when the user asks for them.
//
// GAS FLAT NAMESPACE. Every name in this file ends in 13 (or 13_ for private
// helpers). Reuses, does NOT redefine:
//   collectTabs, flattenDocsTabs_, TOOL_MARKER, CANVAS_TOOL_OPTIONS (Code.gs)
//   getDevelopmentTabBody, extractDocId                             (Code2.gs)
//   runTimeEstimator                                                (Code3.gs)
//   scanDevStructure9_                                              (Code9.gs)
//   dashboardTab10_, locateInstructorResource10_                    (Code10.gs)
//   getDeployInit11, loadSettingsTable11                            (Code11.gs)
//
// SCOPES. Docs advanced service (already enabled) for the Creative Brief table;
// SpreadsheetApp for the finalized boxes (Sheets scope approved 2026-11-03);
// DriveApp for "save as a new Google Doc in the blueprint's folder" (existing
// drive scope).
//
// FLAGS. Things Greg has not confirmed, or that need a real blueprint to check,
// are marked "DEFAULT:" or "UNVERIFIED:" in a comment. Grep for them. Run
// spikeStatusSpike13 (Spikes.gs) first: it prints the highlight colours, the
// Home / Syllabus regions and the Media Deliverables checkbox data this file
// relies on.
// ============================================================


// ── CONSTANTS ────────────────────────────────────────────────

// Review stages, lowest first. 'revised' (yellow) is handled separately: any
// yellow makes an activity Revised (Greg's rule).
var BP_VER_Code13 = '2026-11-08 18:00';  // keep equal to "Last updated on" above; Code14.gs reads it

var STAGE_ORDER_13 = ['untouched', 'drafted', 'green', 'purple', 'built'];
var STAGE_LABEL_13 = {
  untouched: 'Untouched template', drafted: 'Drafted, not highlighted', green: 'Ready for QA (green)',
  purple: 'Approved by QA (purple)', revised: 'Revised (yellow)', built: 'Built in Canvas'
};
// Greg chose the wording "Ready for the Dev Team" (2026-11-03). Ready = green, purple, yellow or struck through
// (the StatusResultsB mockup's wording). Greg may prefer "Highlighted for QA or
// later"; only STAGE_READY_13 and the sidebar label would change.
var STAGE_READY_13 = { green: true, purple: true, revised: true, built: true };

var PLACEHOLDER_REF_RE_13 = /refer to the template blueprint customization/i;
var FINALIZE_RE_13 = /settings are finalized/i;
var MEDIA_STATUSES_13 = ['Not Started', 'In progress', 'Under review', 'Completed', 'Cancelled'];


// ── MENU ENTRY / SIDEBAR API ─────────────────────────────────

function showStatusSidebar13() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar13')
    .setTitle('Development Status Check')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Setup screen data: credits and course length from the Dashboard (editable in
 * the sidebar), how many numbered modules there are, and the settings table.
 */
function getStatusInit13() {
  var doc = DocumentApp.getActiveDocument();
  var out = { devOk: !!getDevelopmentTabBody(doc), course: {}, modules: 0, table: null, tableError: null, tableUrl: null, hasBrief: false };
  var dash = dashboardTab10_(doc).body;
  out.course = {
    credits: parseNumber13_(dashValue13_(doc, dash, /^\s*credits?\b/i)),
    weeks: parseNumber13_(dashValue13_(doc, dash, /^\s*course\s*length\b/i)),
    modulesPerWeek: 1,
    creditsText: dashValue13_(doc, dash, /^\s*credits?\b/i) || '',
    weeksText: dashValue13_(doc, dash, /^\s*course\s*length\b/i) || ''
  };
  if (out.devOk) {
    try { out.modules = scanDevStructure9_(getDevelopmentTabBody(doc)).length; } catch (e) { /* shown as 0 */ }
  }
  var tabs = collectTabs(doc);
  for (var i = 0; i < tabs.length; i++) if (/^\s*creative brief\s*$/i.test(tabs[i].title)) out.hasBrief = true;
  out.themes = [];
  if (out.devOk) {
    try {
      var mods = scanDevStructure9_(getDevelopmentTabBody(doc)), saved = [];
      try { saved = JSON.parse(PropertiesService.getDocumentProperties().getProperty(THEME_PROP_13) || '[]'); } catch (e2) { saved = []; }
      var all = PRESET_THEMES_13.slice();
      saved.forEach(function (t) { if (all.indexOf(t) === -1) all.push(t); });
      var sugg = suggestThemes13_(mods, all);
      var cand = all.concat(sugg), hits = {};
      mods.forEach(function (m) { m.activities.forEach(function (a) { var t = matchTheme13_(a.title, cand); if (t) hits[t] = (hits[t] || 0) + 1; }); });
      all.forEach(function (t) { out.themes.push({ name: t, hits: hits[t] || 0, checked: saved.length ? saved.indexOf(t) !== -1 : !!hits[t], kind: PRESET_THEMES_13.indexOf(t) !== -1 ? 'common' : 'saved' }); });
      sugg.forEach(function (t) { out.themes.push({ name: t, hits: hits[t] || 0, checked: false, kind: 'suggested' }); });
    } catch (e) { /* themes are optional */ }
  }
  var d = getDeployInit11();
  out.table = d.table; out.tableError = d.error; out.tableUrl = d.url;
  return out;
}

/**
 * Runs the check. Read-only.
 * @param {Object} opts  { credits, weeks, modulesPerWeek, tableId,
 *   tabGids: string[], checks: {stage, links, home, time, finalized, media} }
 * @returns {Object} the report (see buildReport13_ for the shape)
 */
function runStatusCheck13(opts) {
  var doc = DocumentApp.getActiveDocument();
  var checks = opts.checks || {};
  var devBody = getDevelopmentTabBody(doc);
  if (!devBody) throw new Error('Could not find a “Development” tab in this blueprint.');

  var report = {
    when: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a'),
    blueprint: doc.getName(),
    checks: checks,
    modules: [], totals: { activities: 0, ready: 0, stages: { untouched: 0, drafted: 0, green: 0, purple: 0, revised: 0, built: 0 }, linksPresent: 0, linksExpected: 0, mostly: {} },
    home: null, syllabus: null, resources: null, finalized: null, media: null, time: null, notes: []
  };

  // Modules and activities.
  var mods = scanDevStructure9_(devBody);
  mods.forEach(function (m) {
    var mod = { label: m.displayLabel || ('Module ' + m.num), activities: [], counts: { untouched: 0, drafted: 0, green: 0, purple: 0, revised: 0, built: 0 }, ready: 0, mostly: 0, pctSum: 0, pct: 0, time: null };
    m.activities.forEach(function (slot) {
      var tool = readToolLine13_(devBody, slot.h4);
      var a = { title: slot.title, tool: tool.tool, stage: null, link: 'na' };
      if (checks.stage) {
        var sd = stageOfSlot13_(slot, tool);
        a.stage = sd.stage; a.pct = sd.pct; a.module = mod.label; a.moduleShort = (m.num != null ? 'M' + m.num : mod.label);
        mod.pctSum += sd.pct;
        if (sd.mostly) {
          a.mostly = sd.mostly; a.waiting = sd.waiting;
          report.totals.mostly[sd.mostly] = (report.totals.mostly[sd.mostly] || 0) + 1;
          mod.mostly += 1;
        }
        mod.counts[a.stage]++;
        if (STAGE_READY_13[a.stage]) mod.ready++;
        report.totals.stages[a.stage]++;
      }
      if (checks.links && tool.tool && tool.tool !== 'Page') {
        // "Any link attribute counts, not the wording" (Greg): a link anywhere
        // on the tool line is enough; placeholder text without a link is missing.
        a.link = tool.hasLink ? 'ok' : 'missing';
        report.totals.linksExpected++;
        if (tool.hasLink) report.totals.linksPresent++;
      }
      mod.activities.push(a);
    });
    mod.pct = mod.activities.length ? Math.round(mod.pctSum / mod.activities.length) : 0;
    report.totals.activities += mod.activities.length;
    report.totals.ready += mod.ready;
    report.modules.push(mod);
  });

  // Activity themes (Greg, 2026-11-07; named "Activity Themes" 2026-11-08): groups of activities by title keyword.
  var themes = (opts.themes || []).map(function (t) { return String(t || '').trim(); }).filter(String);
  if (checks.stage && themes.length) {
    report.themes = buildThemes13_(report.modules, themes);
    try { PropertiesService.getDocumentProperties().setProperty(THEME_PROP_13, JSON.stringify(themes)); } catch (e) { /* remembering is optional */ }
  }

  // Home, Syllabus, Instructor Resources.
  if (checks.home) {
    var regions = homeSyllabusRegions13_(devBody);
    report.home = describeRegion13_(regions.home, false);
    report.syllabus = describeRegion13_(regions.syllabus, true);
    var g = locateInstructorResource10_(devBody, /Instructor Guide/i);
    var s = locateInstructorResource10_(devBody, /Course Suggestions?/i);
    report.resources = {
      guide: g ? (g.existing ? 'linked' : 'no link') : 'not found',
      suggestions: s ? (s.existing ? 'linked' : 'no link') : 'not found'
    };
  }

  // Module time vs. credit-hour guideline (reuses the Time Estimator).
  if (checks.time) {
    var credits = Number(opts.credits), weeks = Number(opts.weeks), mpw = Number(opts.modulesPerWeek) || 1;
    if (!credits || !weeks) {
      report.notes.push('Module time was skipped: credits and course length are both needed.');
    } else {
      var te = runTimeEstimator(credits, weeks, mpw);
      if (te.error) {
        report.notes.push('Module time was skipped: ' + te.error);
      } else {
        report.time = { min: te.minTarget, max: te.maxTarget, missing: te.missing || [] };
        te.modules.forEach(function (tm, i) {
          var mod = report.modules[i];
          if (mod && String(mod.label).toLowerCase().indexOf(String(tm.name).toLowerCase()) === 0) mod.time = { total: tm.total, status: tm.status };
          else if (mod && !mod.time) mod.time = { total: tm.total, status: tm.status };
        });
      }
    }
  }

  // Settings-table tabs marked finalized (read only).
  if (checks.finalized && opts.tableId) {
    try { report.finalized = finalizedReport13_(opts.tableId, opts.tabGids || []); }
    catch (e) { report.notes.push('Settings table tabs could not be read: ' + (e && e.message || e)); }
  }

  // Media deliverables (Creative Brief tab).
  if (checks.media) {
    try { report.media = mediaReport13_(doc); }
    catch (e) { report.notes.push('Media deliverables could not be read: ' + (e && e.message || e)); }
  }

  return report;
}

/**
 * Optional output. dest: 'tab' (a new tab in this blueprint) or 'doc' (a new
 * Google Doc in the blueprint's folder). Never touches existing content.
 * @returns {{url: string, label: string}}
 */
var REPORT_TITLE_RE_13 = /^Status Check (\d{4}-\d{2}-\d{2} \d{2}:\d{2})( full| plain text short| plain text)$/;

function reportKind13_(format) {
  return format === 'plainshort' ? ' plain text short' : (format === 'plain' ? ' plain text' : ' full');
}

/**
 * Saves one report. replace = true: if an earlier Status Check report of the same
 * kind exists in the same place (a tab in this blueprint, or a Doc in its folder),
 * its contents are rewritten in place instead of making a new tab or Doc. The
 * sidebar asks the user to confirm first (previewStatusReplace13); only reports
 * named "Status Check <date> <kind>" are ever touched.
 */
function saveStatusReport13(report, dest, format, replace) {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var plain = (format === 'plain' || format === 'plainshort');
  var lines = plain ? plainLines13_(report, format === 'plainshort').map(function (t) { return { text: t, h: 0 }; }) : reportLines13_(report);
  var kind = reportKind13_(format), title = 'Status Check ' + stamp + kind;
  if (replace) {
    var doc = DocumentApp.getActiveDocument();
    if (dest === 'tab') {
      var pt = findPriorTab13_(doc, kind);
      if (pt) {
        var body = pt.tab.asDocumentTab().getBody();
        body.clear();
        writeLines13_(body, lines);
        var renamed = false;
        try { pt.tab.setTitle(title); renamed = true; } catch (e) { /* older runtimes: keep the old name */ }
        return { url: doc.getUrl() + '?tab=' + encodeURIComponent(pt.tab.getId()), label: 'Replaced the contents of tab “' + (renamed ? title : pt.name) + '”' };
      }
    } else if (dest === 'doc') {
      var pd = findPriorDoc13_(doc, kind);
      if (pd) {
        var nd = DocumentApp.openById(pd.file.getId());
        nd.getBody().clear();
        writeLines13_(nd.getBody(), lines);
        nd.setName((report.blueprint || doc.getName()) + ' - ' + title);
        nd.saveAndClose();
        return { url: pd.file.getUrl(), label: 'Replaced the contents of Google Doc “' + (report.blueprint || doc.getName()) + ' - ' + title + '” (its version history keeps the old one)' };
      }
    }
  }
  return saveLines13_(lines, title, dest, report.blueprint);
}

/** Tells the sidebar which earlier reports a "replace" save would overwrite (nothing is changed here). */
function previewStatusReplace13(dests, formats) {
  var doc = DocumentApp.getActiveDocument(), out = [];
  (dests || []).forEach(function (d) {
    (formats || []).forEach(function (f) {
      var kind = reportKind13_(f), found = d === 'tab' ? findPriorTab13_(doc, kind) : findPriorDoc13_(doc, kind);
      out.push({ dest: d, format: f, name: found ? found.name : null });
    });
  });
  return out;
}

/** Most recent tab in this blueprint named "Status Check <date> <kind>", or null. */
function findPriorTab13_(doc, kind) {
  var best = null;
  (function walk(tabs) {
    tabs.forEach(function (t) {
      var m = REPORT_TITLE_RE_13.exec(t.getTitle());
      if (m && m[2] === kind && (!best || m[1] > best.stamp)) best = { stamp: m[1], tab: t, name: t.getTitle() };
      walk(t.getChildTabs());
    });
  })(doc.getTabs());
  return best;
}

/** Most recent Doc in the blueprint's folder named "<blueprint> - Status Check <date> <kind>", or null. */
function findPriorDoc13_(doc, kind) {
  var parents = DriveApp.getFileById(doc.getId()).getParents();
  if (!parents.hasNext()) return null;
  var it = parents.next().getFilesByType(MimeType.GOOGLE_DOCS), prefix = doc.getName() + ' - ', best = null;
  while (it.hasNext()) {
    var f = it.next(), n = f.getName();
    if (n.indexOf(prefix) !== 0) continue;
    var m = REPORT_TITLE_RE_13.exec(n.slice(prefix.length));
    if (m && m[2] === kind && (!best || m[1] > best.stamp)) best = { stamp: m[1], file: f, name: n };
  }
  return best;
}

/** Plain text version for the sidebar's copy box (e.g. to paste into QuickBase). */
function statusPlainText13(report, short) {
  return plainLines13_(report, !!short).join('\n');
}

/**
 * Writes a list of {text, h} lines to a new tab in this blueprint (dest 'tab') or
 * a new Google Doc in the blueprint's folder (dest 'doc'). Shared with the
 * Version Check (Code14.gs). Never touches existing content.
 * @returns {{url: string, label: string}}
 */
function saveLines13_(lines, title, dest, blueprintName) {
  var doc = DocumentApp.getActiveDocument();
  if (dest === 'tab') {
    if (typeof doc.addTab !== 'function') throw new Error('This version of Apps Script cannot add a tab from a script. Use “new Google Doc” instead.');
    var tab = doc.addTab(title);
    var body = tab.asDocumentTab().getBody();
    body.clear();
    writeLines13_(body, lines);
    return { url: doc.getUrl(), label: 'New tab “' + title + '” in this blueprint' };
  }
  if (dest === 'doc') {
    var nd = DocumentApp.create((blueprintName || doc.getName()) + ' - ' + title);
    writeLines13_(nd.getBody(), lines);
    nd.saveAndClose();
    try {
      var parents = DriveApp.getFileById(doc.getId()).getParents();
      if (parents.hasNext()) DriveApp.getFileById(nd.getId()).moveTo(parents.next());
    } catch (e) { /* stays in My Drive; the label says where it is */ }
    return { url: nd.getUrl(), label: 'New Google Doc “' + nd.getName() + '”' };
  }
  throw new Error('Unknown save option.');
}

function writeLines13_(body, lines) {
  lines.forEach(function (l) {
    if (typeof l.bar === 'number') { barLine13_(body, l.text, l.bar); return; }
    var p = body.appendParagraph(l.text);
    if (l.h) p.setHeading(l.h === 1 ? DocumentApp.ParagraphHeading.HEADING1 : DocumentApp.ParagraphHeading.HEADING2);
    else p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  });
}

/**
 * A label line followed by a progress bar: a one-row table whose left cells are
 * shaded green for the percent done. If the table cannot be made, a row of block
 * characters is used instead. UNVERIFIED in a real document.
 */
function barLine13_(body, label, pct) {
  var p = body.appendParagraph(label);
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  pct = Math.max(0, Math.min(100, pct));
  try {
    var done = Math.round(pct / 5), cells = [];
    for (var i = 0; i < 20; i++) cells.push('');
    var t = body.appendTable([cells]);
    var row = t.getRow(0);
    for (var c = 0; c < 20; c++) {
      var cell = row.getCell(c);
      cell.setBackgroundColor(c < done ? '#8fd19e' : '#eeeeee');
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(0).setPaddingRight(0);
      cell.setWidth(14);
    }
    t.setBorderWidth(0.5).setBorderColor('#bbbbbb');
  } catch (e) {
    var n = Math.round(pct / 5);
    body.appendParagraph(new Array(n + 1).join('\u2588') + new Array(21 - n).join('\u2591'));
  }
}


// ── THEMES ───────────────────────────────────────────────────

var THEME_PROP_13 = 'BP_STATUS_THEMES_13';
var PRESET_THEMES_13 = ['Quizzes', 'Exams', 'Readings', 'Videos', 'Readings & Multimedia', 'Reflections', 'Assignments', 'Final Project'];
var THEME_STOP_13 = ['the', 'and', 'for', 'with', 'your', 'from', 'this', 'that', 'into', 'about', 'week', 'module', 'unit', 'part', 'activity', 'canvas', 'page', 'discussion', 'discussions'];

/** Lower-case words of a title, '&' read as 'and', plurals trimmed so Quizzes matches Quiz. */
function themeWords13_(s) {
  return String(s || '').toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, ' ').trim().split(' ').filter(String).map(function (w) {
    if (w.length > 4 && /ies$/.test(w)) w = w.slice(0, -3) + 'y';
    else if (w.length > 4 && /(sses|zzes|xes|ches|shes)$/.test(w)) w = w.slice(0, -2);
    else if (w.length > 3 && /[^s]s$/.test(w)) w = w.slice(0, -1);
    return w.replace(/zz$/, 'z');
  });
}

/** The most specific theme whose words all appear, in order, in the title; '' if none. */
function matchTheme13_(title, themes) {
  var tw = ' ' + themeWords13_(title).join(' ') + ' ', best = '', bestN = 0, bestLen = 0;
  themes.forEach(function (t) {
    var kw = themeWords13_(t); if (!kw.length) return;
    if (tw.indexOf(' ' + kw.join(' ') + ' ') === -1) return;
    var len = kw.join(' ').length;
    if (kw.length > bestN || (kw.length === bestN && len > bestLen)) { best = t; bestN = kw.length; bestLen = len; }
  });
  return best;
}

/** Groups activities by theme. Each activity goes to its most specific theme only; unmatched ones go to "Other". */
function buildThemes13_(modules, themes) {
  var groups = {}, order = [], other = [];
  themes.forEach(function (t) { groups[t] = []; order.push(t); });
  modules.forEach(function (m) {
    m.activities.forEach(function (a) {
      var t = matchTheme13_(a.title, themes);
      var row = { module: a.module, moduleShort: a.moduleShort, title: a.title, tool: a.tool, stage: a.stage, pct: a.pct || 0, mostly: a.mostly, waiting: a.waiting };
      (t ? groups[t] : other).push(row);
    });
  });
  function summarize(name, acts) {
    var sum = 0, ready = 0;
    acts.forEach(function (x) { sum += x.pct; if (STAGE_READY_13[x.stage]) ready++; });
    return { name: name, count: acts.length, ready: ready, pct: acts.length ? Math.round(sum / acts.length) : 0, acts: acts };
  }
  var list = order.filter(function (t) { return groups[t].length; }).map(function (t) { return summarize(t, groups[t]); });
  return { list: list, other: other.length ? summarize('Other activities', other) : null };
}

/** Words or word pairs that recur in titles across several modules: candidate themes. */
function suggestThemes13_(modules, known) {
  var uni = {}, bi = {}, forms = {}, nMods = modules.length;
  var knownKeys = {}; known.forEach(function (t) { knownKeys[themeWords13_(t).join(' ')] = true; });
  modules.forEach(function (m, mi) {
    var seenU = {}, seenB = {};
    m.activities.forEach(function (a) {
      var raw = String(a.title || '').toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, ' ').trim().split(' ').filter(String);
      var w = themeWords13_(a.title);
      for (var i = 0; i < w.length; i++) {
        if (w[i].length >= 4 && !/^\d+$/.test(w[i]) && THEME_STOP_13.indexOf(w[i]) === -1 && !seenU[w[i]]) { seenU[w[i]] = true; uni[w[i]] = (uni[w[i]] || 0) + 1; forms[w[i]] = raw[i]; }
        if (i + 1 < w.length && w[i].length >= 3 && w[i + 1].length >= 3 && THEME_STOP_13.indexOf(w[i]) === -1 && THEME_STOP_13.indexOf(w[i + 1]) === -1) {
          var k = w[i] + ' ' + w[i + 1];
          if (!seenB[k]) { seenB[k] = true; bi[k] = (bi[k] || 0) + 1; forms[k] = raw[i] + ' ' + raw[i + 1]; }
        }
      }
    });
  });
  var need = Math.min(3, nMods), out = [];
  Object.keys(bi).forEach(function (k) { if (bi[k] >= need && !knownKeys[k]) out.push({ k: k, n: bi[k] + 0.5 }); });
  Object.keys(uni).forEach(function (k) {
    if (uni[k] < need || knownKeys[k]) return;
    var covered = Object.keys(bi).some(function (b) { return bi[b] >= need && b.split(' ').indexOf(k) !== -1 && bi[b] >= uni[k] * 0.8; });
    if (!covered) out.push({ k: k, n: uni[k] });
  });
  out.sort(function (x, y) { return y.n - x.n; });
  // Two suggestions that match exactly the same activities (e.g. "Reminder" and
  // "Topic Replies" in "Reminder: Discussion [Topic]: - Replies") are one theme
  // to the user: keep the more specific one (more words, then longer).
  var titles = [];
  modules.forEach(function (m) { m.activities.forEach(function (a) { titles.push(' ' + themeWords13_(a.title).join(' ') + ' '); }); });
  function matchSet(k) { var hits = []; titles.forEach(function (t, i) { if (t.indexOf(' ' + k + ' ') !== -1) hits.push(i); }); return hits.join(','); }
  var best = {};
  out.forEach(function (o) {
    var sig = matchSet(o.k), cur = best[sig];
    var w = o.k.split(' ').length;
    if (!cur || w > cur.w || (w === cur.w && o.k.length > cur.o.k.length)) best[sig] = { o: o, w: w };
  });
  out = out.filter(function (o) { var b = best[matchSet(o.k)]; return b && b.o === o; });
  return out.slice(0, 8).map(function (o) { return forms[o.k].replace(/\b[a-z]/g, function (c) { return c.toUpperCase(); }); });
}


// ── ACTIVITY STAGE ───────────────────────────────────────────

/** Tool line under an activity heading: tool name, any link on it, its paragraph. */
function readToolLine13_(body, h4) {
  var out = { tool: '', hasLink: false, el: null };
  var H = DocumentApp.ParagraphHeading;
  var start = body.getChildIndex(h4), n = body.getNumChildren();
  for (var i = start + 1; i < Math.min(start + 8, n); i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para = child.asParagraph(), h = para.getHeading();
    if (h === H.HEADING2 || h === H.HEADING3 || h === H.HEADING4) break;
    var text = para.getText(), trimmed = text.trim();
    if (text.indexOf(TOOL_MARKER) === -1 && text.indexOf('Link to settings tab') === -1 && CANVAS_TOOL_OPTIONS.indexOf(trimmed) === -1) continue;
    var cut = text.search(new RegExp('\\s*(?:' + TOOL_MARKER + '|;)'));
    out.tool = (cut === -1 ? trimmed : text.slice(0, cut)).trim();
    out.el = para.editAsText();
    out.hasLink = anyLink13_(out.el);
    return out;
  }
  return out;
}

function anyLink13_(textEl) {
  var idx = textEl.getTextAttributeIndices();
  for (var i = 0; i < idx.length; i++) if (textEl.getLinkUrl(idx[i])) return true;
  return false;
}

/**
 * Stage of one activity slot.
 * DEFAULT: untouched = no content under the heading (the directions placeholder
 * is still there or nothing is). Otherwise the stage comes from the highlight on
 * the activity's CONTENT text; if the content has no readable text (tables only),
 * the heading and tool line are used. Any yellow = Revised; otherwise the lowest
 * stage found (Greg confirmed yellow; "lowest" is my proposal, not confirmed).
 */
function stageOfSlot13_(slot, tool) {
  if (!slot.content.length) return { stage: 'untouched', pct: 0 };
  var segs = [];
  slot.content.forEach(function (el) { collectSegments13_(el, segs); });
  if (!segs.length) {
    collectSegments13_(slot.h4, segs);
    if (tool.el) collectSegments13_(tool.el, segs);
  }
  return stageDetail13_(segs);
}

/** Pushes {bg, strike} for each non-blank run of text inside an element. */
function collectSegments13_(el, out) {
  var T = DocumentApp.ElementType, type = el.getType();
  var t = null;
  if (type === T.TEXT) t = el.asText();
  else if (type === T.PARAGRAPH) t = el.asParagraph().editAsText();
  else if (type === T.LIST_ITEM) t = el.asListItem().editAsText();
  if (t) {
    var text = t.getText();
    if (!text) return;
    var idx = t.getTextAttributeIndices();
    for (var i = 0; i < idx.length; i++) {
      var end = i + 1 < idx.length ? idx[i + 1] : text.length;
      if (!text.substring(idx[i], end).trim()) continue;
      var piece = text.substring(idx[i], end).trim();
      out.push({ bg: t.getBackgroundColor(idx[i]), strike: !!t.isStrikethrough(idx[i]), fg: t.getForegroundColor(idx[i]), len: piece.length, text: piece });
    }
    return;
  }
  if (type === T.TABLE || type === T.TABLE_ROW || type === T.TABLE_CELL) {
    for (var c = 0; c < el.getNumChildren(); c++) collectSegments13_(el.getChild(c), out);
  }
}

function stageFromSegments13_(segs) { return stageDetail13_(segs).stage; }

// "Mostly" rule (Greg, 2026-11-07): when at least this share of an activity's
// text (red builder instructions excluded) sits at one finished stage and the
// rest is only unhighlighted text, the activity is still Drafted but is tagged
// "Mostly <stage>" and the leftover snippets are listed.
var MOSTLY_SHARE_13 = 0.85;
var MOSTLY_STAGES_13 = ['green', 'purple', 'built'];

/**
 * Bold red text is the instruction to the course developer (builder) and is
 * usually left unhighlighted, so it is ignored when working out the stage.
 * Any red counts, bold or not; the exact red is #ff0000 (RED in Code.gs) but a
 * nearby shade is accepted. UNVERIFIED: spikeStatusSpike13 lists the font colours.
 */
function isInstructionRed13_(hex) {
  var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(hex || ''));
  if (!m) return false;
  var r = parseInt(m[1], 16) / 255, g = parseInt(m[2], 16) / 255, b = parseInt(m[3], 16) / 255;
  return r >= 0.55 && g <= 0.4 && b <= 0.4 && (r - Math.max(g, b)) >= 0.3;
}

/**
 * {stage, pct, mostly, waiting}. pct = share of the activity's text (red
 * instructions excluded) that is highlighted green, purple or yellow, or struck
 * through; unhighlighted text counts as 0. mostly/waiting only for "Mostly" activities.
 */
function stageDetail13_(allSegs) {
  var segs = allSegs.filter(function (s) { return !isInstructionRed13_(s.fg); });
  if (!segs.length) return { stage: 'drafted', pct: 0 };
  var lowest = STAGE_ORDER_13.length, yellow = false, weight = {}, total = 0, hl = 0, drafted = [];
  segs.forEach(function (s) {
    var st = s.strike ? 'built' : stageOfColor13_(s.bg);
    var w = s.len || 1;
    weight[st] = (weight[st] || 0) + w;
    total += w;
    if (st !== 'drafted') hl += w;
    if (st === 'drafted') drafted.push(s.text || '');
    if (st === 'revised') { yellow = true; return; }
    var rank = STAGE_ORDER_13.indexOf(st);
    if (rank < lowest) lowest = rank;
  });
  var pct = total ? Math.round(100 * hl / total) : 0;
  if (yellow) return { stage: 'revised', pct: pct };
  var stage = STAGE_ORDER_13[lowest] || 'drafted';
  if (stage === 'drafted' && total) {
    var best = null;
    MOSTLY_STAGES_13.forEach(function (k) { if ((weight[k] || 0) > (best ? weight[best] : 0)) best = k; });
    // Only if the rest is purely unhighlighted (no other stage mixed in).
    var others = Object.keys(weight).filter(function (k) { return k !== best && k !== 'drafted'; });
    if (best && !others.length && weight[best] / total >= MOSTLY_SHARE_13) {
      var waiting = drafted.map(function (t) { return t.length > 40 ? t.slice(0, 37) + '...' : t; }).slice(0, 3);
      return { stage: 'drafted', pct: pct, mostly: best, waiting: waiting };
    }
  }
  return { stage: stage, pct: pct };
}

/**
 * Highlight colour to stage, by hue rather than exact hex, because the exact
 * hexes QA uses are not written down anywhere. UNVERIFIED: run spikeStatusSpike13
 * and check that every colour in the blueprint is classified the way you expect.
 * Cyan (the Design Map note highlight) and anything unrecognised count as
 * "drafted", not as a review stage.
 */
function stageOfColor13_(hex) {
  var c = hueOf13_(hex);
  if (!c) return 'drafted';
  var h = c.h;
  if (h >= 40 && h < 70) return 'revised';     // yellow
  if (h >= 75 && h < 165) return 'green';
  if (h >= 245 && h < 325) return 'purple';
  return 'drafted';                              // cyan, orange, red, pink, grey ...
}

function hueOf13_(hex) {
  var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(hex || ''));
  if (!m) return null;
  var r = parseInt(m[1], 16) / 255, g = parseInt(m[2], 16) / 255, b = parseInt(m[3], 16) / 255;
  var max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min, l = (max + min) / 2;
  if (d < 0.08 || l > 0.985) return null;        // white / grey: not a highlight
  var h;
  if (max === r) h = ((g - b) / d) % 6; else if (max === g) h = (b - r) / d + 2; else h = (r - g) / d + 4;
  h *= 60; if (h < 0) h += 360;
  return { h: h, l: l };
}


// ── HOME / SYLLABUS ──────────────────────────────────────────

/**
 * Finds the Home and Syllabus sections in the Development tab. Each starts with
 * a single-cell header table ("Home (Course Navigation Menu Link)", "Syllabus
 * (Course Navigation Menu Link)") and runs to the next header table or H2.
 * UNVERIFIED: read from the template copy; the spike prints what it finds.
 */
function homeSyllabusRegions13_(body) {
  var T = DocumentApp.ElementType, H = DocumentApp.ParagraphHeading;
  var out = { home: null, syllabus: null }, cur = null, n = body.getNumChildren();
  for (var i = 0; i < n; i++) {
    var c = body.getChild(i), type = c.getType();
    if (type === T.TABLE) {
      var txt = c.asTable().getText().trim();
      if (/^home\b/i.test(txt))          { cur = out.home = { elements: [] }; continue; }
      if (/^syllabus\b/i.test(txt))      { cur = out.syllabus = { elements: [] }; continue; }
      if (/^modules\b/i.test(txt))       { cur = null; if (out.home || out.syllabus) break; continue; }
    }
    if (type === T.PARAGRAPH && c.asParagraph().getHeading() === H.HEADING2) { if (cur) cur = null; if (out.syllabus) break; continue; }
    if (cur) cur.elements.push(c);
  }
  return out;
}

/** Whether a section has text, a link, and which stage it is at. */
function describeRegion13_(region, isSyllabus) {
  if (!region) return { found: false };
  var text = '', link = false, segs = [];
  region.elements.forEach(function (el) {
    var t = null, ty = el.getType(), T = DocumentApp.ElementType;
    if (ty === T.PARAGRAPH) t = el.asParagraph().editAsText(); else if (ty === T.LIST_ITEM) t = el.asListItem().editAsText();
    if (t) { text += '\n' + t.getText(); if (!link && t.getText() && anyLink13_(t)) link = true; }
    collectSegments13_(el, segs);
  });
  var real = text.split('\n').map(function (s) { return s.trim(); }).filter(function (s) { return s && !/select syllabus format/i.test(s); }).join(' ');
  var untouched = !real || PLACEHOLDER_REF_RE_13.test(real);
  var res = { found: true, untouched: untouched, stage: untouched ? 'untouched' : stageFromSegments13_(segs) };
  if (isSyllabus) res.format = untouched ? '' : (link ? 'linked document' : 'written in place');
  return res;
}

/** "Credits" / "Course Length" value on the Dashboard: the cell beside the label, or text after it. */
/**
 * Dashboard field value. The value is usually a dropdown chip ("Course Length: [7 Weeks]"),
 * which DocumentApp cannot read (it returns empty text), so read it with the Docs API first
 * (found by Greg's spike run, 2026-11-08), then fall back to plain text. "--" and
 * "Select ..." placeholders count as empty.
 */
function dashValue13_(doc, dashBody, labelRe) {
  var v = '';
  try { v = readDashChip13_(doc, labelRe); } catch (e) { v = ''; }
  if (!v) v = readDashField13_(dashBody, labelRe);
  return /^\s*(--|—|select\b.*)\s*$/i.test(v || '') ? '' : (v || '');
}

function readDashChip13_(doc, labelRe) {
  var tab = null;
  flattenDocsTabs_(Docs.Documents.get(doc.getId(), { includeTabsContent: true })).forEach(function (t) {
    if (!tab && /dashboard/i.test((t.tabProperties || {}).title || '')) tab = t;
  });
  if (!tab) return '';
  // Value in a paragraph: the first chip after the label, else the text after it.
  function inParagraph(p) {
    var els = p.elements || [], at = -1, i;
    for (i = 0; i < els.length; i++) if (els[i].textRun && labelRe.test(els[i].textRun.content || '')) { at = i; break; }
    if (at < 0) return null;
    var after = ((els[at].textRun.content || '').replace(/^[^:]*:?\s*/, '')).replace(/\|/g, '').trim();
    if (after) return after;
    for (i = at + 1; i < els.length; i++) {
      if (els[i].dropdown) return String(((els[i].dropdown.dropdownProperties || {}).displayValue) || '').trim();
      if (els[i].textRun) { var t = String(els[i].textRun.content || '').replace(/\|/g, '').trim(); if (t) return t; }
    }
    return '';
  }
  function firstValue(cell) {
    var out = '';
    (cell.content || []).forEach(function (blk) {
      if (out || !blk.paragraph) return;
      (blk.paragraph.elements || []).forEach(function (e) {
        if (out) return;
        if (e.dropdown) out = String(((e.dropdown.dropdownProperties || {}).displayValue) || '').trim();
        else if (e.textRun) out = String(e.textRun.content || '').replace(/\|/g, '').trim();
      });
    });
    return out;
  }
  var result = null;
  (function walk(content) {
    (content || []).forEach(function (el) {
      if (result !== null) return;
      if (el.paragraph) { var v = inParagraph(el.paragraph); if (v !== null) result = v; return; }
      if (el.table) (el.table.tableRows || []).forEach(function (row) {
        var cells = row.tableCells || [];
        for (var c = 0; c < cells.length && result === null; c++) {
          var hit = null;
          (cells[c].content || []).forEach(function (blk) { if (hit === null && blk.paragraph) hit = inParagraph(blk.paragraph); });
          if (hit !== null) { result = hit || (c + 1 < cells.length ? firstValue(cells[c + 1]) : ''); return; }
          walk(cells[c].content);
        }
      });
    });
  })(((tab.documentTab || {}).body || {}).content);
  return result || '';
}

function readDashField13_(dashBody, labelRe) {
  var T = DocumentApp.ElementType, tables = [];
  (function walk(el) {
    for (var i = 0; i < el.getNumChildren(); i++) {
      var c = el.getChild(i);
      if (c.getType() === T.TABLE) tables.push(c.asTable());
      else if (c.getType() === T.TABLE_CELL) walk(c);
    }
  })(dashBody);
  for (var t = 0; t < tables.length; t++) {
    var tb = tables[t];
    for (var r = 0; r < tb.getNumRows(); r++) {
      var row = tb.getRow(r);
      for (var c = 0; c < row.getNumCells(); c++) {
        var txt = row.getCell(c).getText().trim();
        if (!labelRe.test(txt)) continue;
        var rest = txt.replace(/^[^:]*:?\s*/, '').trim();
        if (rest && rest !== txt) return rest;
        if (c + 1 < row.getNumCells()) return row.getCell(c + 1).getText().trim();
        if (r + 1 < tb.getNumRows() && tb.getRow(r + 1).getNumCells() > c) return tb.getRow(r + 1).getCell(c).getText().trim();
      }
    }
  }
  return '';
}

function parseNumber13_(s) {
  var m = String(s || '').match(/\d+(?:\.\d+)?/);
  return m ? Number(m[0]) : null;
}


// ── SETTINGS TABLE: FINALIZED BOXES (read only) ──────────────

/**
 * For each chosen tab: 'ticked', 'unticked', or 'none' (no "settings are
 * finalized" row or no checkbox beside it). READ ONLY: never writes a cell.
 * "none" tabs are listed but not counted as unfinished (Greg's rule).
 */
function finalizedReport13_(tableId, gids) {
  var ss = SpreadsheetApp.openById(extractDocId(tableId) || tableId);
  var want = {}; gids.forEach(function (g) { want[String(g)] = true; });
  var tabs = [], ticked = 0, counted = 0;
  ss.getSheets().forEach(function (sh) {
    if (!want[String(sh.getSheetId())]) return;
    var state = finalizedState13_(sh);
    tabs.push({ name: sh.getName(), state: state, gid: String(sh.getSheetId()) });
    if (state !== 'none') { counted++; if (state === 'ticked') ticked++; }
  });
  return { ticked: ticked, counted: counted, tabs: tabs, url: 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit' };
}

function finalizedState13_(sheet) {
  var vals = sheet.getDataRange().getValues();
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      if (typeof vals[r][c] !== 'string' || !FINALIZE_RE_13.test(vals[r][c])) continue;
      for (var k = c; k < vals[r].length; k++) if (typeof vals[r][k] === 'boolean') return vals[r][k] ? 'ticked' : 'unticked';
      if (r + 1 < vals.length) for (var k2 = c; k2 < vals[r + 1].length; k2++) if (typeof vals[r + 1][k2] === 'boolean') return vals[r + 1][k2] ? 'ticked' : 'unticked';
      return 'none';
    }
  }
  return 'none';
}


// ── MEDIA DELIVERABLES (Creative Brief tab) ──────────────────

/**
 * Reads the Media Deliverables table on the "Creative Brief" tab through the
 * Docs API. Columns are found by header text, not position.
 * UNVERIFIED (main risk, see HANDOFF 3.4): whether a ticked "Link Added to Dev
 * Tab (IDC)" box can be read. This assumes a Docs checklist item, whose ticked
 * state shows as strikethrough on the paragraph's closing newline. If no
 * checklist item is found the row's box is reported as "unknown" and the row is
 * NOT flagged, so a failed read can never produce a false alarm.
 * The box is never written.
 */
function mediaReport13_(doc) {
  var structured = Docs.Documents.get(doc.getId(), { includeTabsContent: true });
  var tab = null;
  flattenDocsTabs_(structured).forEach(function (t) {
    if (!tab && /^\s*creative brief\s*$/i.test((t.tabProperties || {}).title || '')) tab = t;
  });
  if (!tab) return { found: false, reason: 'no tab named “Creative Brief”' };
  var tabUrl = doc.getUrl() + '?tab=' + encodeURIComponent(tab.tabProperties.tabId);
  var content = ((tab.documentTab || {}).body || {}).content || [];

  // The "Media Deliverables" heading is itself a one-cell table in the template (found by
  // Greg's spike run, 2026-11-08), so find the table by its header row instead:
  // the one whose first row has both "Media Type" and "Final Link".
  var table = findMediaTable13_(content);
  if (!table) return { found: false, reason: 'no Media Deliverables table (one whose header row has “Media Type” and “Final Link”)', tabUrl: tabUrl };

  var rows = table.tableRows || [];
  if (!rows.length) return { found: false, reason: 'the Media Deliverables table is empty', tabUrl: tabUrl };
  var head = rows[0].tableCells.map(function (c) { return cellText13_(c); });
  function col(re) { for (var k = 0; k < head.length; k++) if (re.test(head[k])) return k; return -1; }
  var cType = col(/media\s*type/i), cTitle = col(/^\s*title/i), cLink = col(/final\s*link/i), cDev = col(/link\s*added/i), cStat = col(/media\s*status/i);
  if (cLink < 0 || cDev < 0 || cStat < 0) {
    return { found: false, reason: 'could not find the Final Link, Link Added, and Media Status columns by their headers', tabUrl: tabUrl };
  }

  var out = { found: true, tabUrl: tabUrl, items: 0, byStatus: {}, flagged: [], reverse: [], unknownBox: 0 };
  for (var r = 1; r < rows.length; r++) {
    var cells = rows[r].tableCells;
    var title = cTitle >= 0 ? cellText13_(cells[cTitle]) : '';
    var type = cType >= 0 ? cellText13_(cells[cType]) : '';
    if (/^select\s*type$/i.test(type)) type = '';                    // dropdown placeholder
    var hasFinal = cellHasLink13_(cells[cLink]);
    if (!title && !type && !hasFinal) continue;                       // blank row (Media Status defaults to Not Started)
    var status = statusOf13_(cells[cStat]);
    var box = checkboxState13_(cells[cDev]);                         // 'ticked' | 'unticked' | 'unknown'
    out.items++;
    var key = status || '(no status)';
    out.byStatus[key] = (out.byStatus[key] || 0) + 1;
    if (status === 'Cancelled') continue;                             // out of scope (Greg)
    if (box === 'unknown') out.unknownBox++;
    var item = { type: type, title: title, row: r + 1 };
    if (status === 'Completed' && hasFinal && box === 'unticked') out.flagged.push(item);
    if (box === 'ticked' && !hasFinal) out.reverse.push(item);
  }
  return out;
}

/**
 * Table (Docs API) whose first row has "Media Type" and "Final Link", or null.
 * Looks inside other tables too (layout tables), not just at the top level.
 */
function findMediaTable13_(content) {
  for (var i = 0; i < (content || []).length; i++) {
    var t = content[i].table;
    if (!t || !t.tableRows || !t.tableRows.length) continue;
    var head = t.tableRows[0].tableCells.map(function (c) { return cellText13_(c); }).join(' | ');
    if (/media\s*type/i.test(head) && /final\s*link/i.test(head)) return t;
    for (var r = 0; r < t.tableRows.length; r++) {
      var cells = t.tableRows[r].tableCells || [];
      for (var c = 0; c < cells.length; c++) {
        var inner = findMediaTable13_(cells[c].content);
        if (inner) return inner;
      }
    }
  }
  return null;
}

/** For the spike: one line per table (top level and nested) with its first row's text. */
function describeTables13_(content, depth, out) {
  (content || []).forEach(function (el) {
    var t = el.table;
    if (!t || !t.tableRows || !t.tableRows.length) return;
    out.push(new Array((depth || 0) + 1).join('  ') + 'table, ' + t.tableRows.length + ' rows; first row: ' +
      t.tableRows[0].tableCells.map(function (c) { return cellText13_(c).replace(/\s+/g, ' ').slice(0, 30); }).join(' | '));
    t.tableRows.forEach(function (row) { (row.tableCells || []).forEach(function (c) { describeTables13_(c.content, (depth || 0) + 1, out); }); });
  });
  return out;
}

function paraText13_(p) {
  return (p.elements || []).map(function (e) {
    if (e.textRun) return e.textRun.content;
    if (e.dropdown) return String(((e.dropdown.dropdownProperties || {}).displayValue) || '');
    return '';
  }).join('');
}

/** All text in a Docs API table cell. */
function cellText13_(cell) {
  var s = '';
  (cell.content || []).forEach(function (blk) {
    if (blk.paragraph) s += paraText13_(blk.paragraph) + ' ';
  });
  return s.replace(/\s+/g, ' ').trim();
}

/** Media Status is a dropdown chip: match it by its known option names anywhere in the cell. */
function statusOf13_(cell) {
  var hay = cellText13_(cell || {}) || JSON.stringify(cell || {});
  for (var i = 0; i < MEDIA_STATUSES_13.length; i++) {
    if (new RegExp(MEDIA_STATUSES_13[i].replace(/ /g, '\\s*'), 'i').test(hay)) return MEDIA_STATUSES_13[i];
  }
  return '';
}

function cellHasLink13_(cell) {
  var found = false;
  (cell.content || []).forEach(function (blk) {
    if (!blk.paragraph) return;
    (blk.paragraph.elements || []).forEach(function (e) {
      if (e.richLink) found = true;
      if (e.textRun) {
        if ((e.textRun.textStyle || {}).link) found = true;
        if (/https?:\/\//i.test(e.textRun.content || '')) found = true;
      }
    });
  });
  return found;
}

/** Checklist item: ticked shows as strikethrough on the paragraph's closing newline. */
function checkboxState13_(cell) {
  var state = 'unknown';
  (cell.content || []).forEach(function (blk) {
    var p = blk.paragraph;
    if (!p || !p.bullet) return;
    var els = p.elements || [], last = els[els.length - 1];
    var struck = false;
    els.forEach(function (e) { if (e.textRun && (e.textRun.textStyle || {}).strikethrough) struck = true; });
    if (last && last.textRun && (last.textRun.textStyle || {}).strikethrough) struck = true;
    state = struck ? 'ticked' : 'unticked';
  });
  return state;
}


// ── SAVED REPORT TEXT ────────────────────────────────────────

function reportLines13_(r) {
  var L = [], add = function (t, h) { L.push({ text: t, h: h || 0 }); }, bar = function (t, pct) { L.push({ text: t, h: 0, bar: pct }); };
  add('Development Status Check', 1);
  add(r.blueprint + ' · checked ' + r.when + ' · ' + r.totals.activities + ' activities · ' + r.modules.length + ' modules');
  if (r.media && r.media.found && r.media.flagged.length) add('ACTION: ' + r.media.flagged.length + ' completed media item(s) have a Final Link but are not ticked “Link Added to Dev Tab”.');
  if (r.home) {
    add('Home and Syllabus', 2);
    add('Home: ' + regionText13_(r.home, false));
    add('Syllabus: ' + regionText13_(r.syllabus, true));
    add('Instructor Guide: ' + r.resources.guide + ' · Course Suggestions: ' + r.resources.suggestions);
  }
  if (r.checks.stage) {
    add('Course totals', 2);
    add(r.totals.ready + ' of ' + r.totals.activities + ' activities are Ready for the Dev Team.');
    Object.keys(STAGE_LABEL_13).forEach(function (k) { add(STAGE_LABEL_13[k] + ': ' + r.totals.stages[k]); });
    Object.keys(r.totals.mostly || {}).forEach(function (k) { add('Mostly ' + k + ' (counted as drafted, waiting on a few pieces): ' + r.totals.mostly[k]); });
  }
  if (r.checks.links) add('Settings table links: ' + r.totals.linksPresent + ' of ' + r.totals.linksExpected);
  if (r.themes) {
    add('By activity theme', 2);
    themeGroups13_(r).forEach(function (g) {
      bar(g.name + ': ' + g.pct + '% complete · ' + g.ready + ' of ' + g.count + ' ready · ' + g.acts.map(function (x) { return x.moduleShort + ' ' + x.pct + '%'; }).join(', '), g.pct);
    });
  }
  add('By module', 2);
  r.modules.forEach(function (m) {
    if (r.checks.stage) bar(m.label + ': ' + m.pct + '% complete · ' + m.ready + ' of ' + m.activities.length + ' ready', m.pct);
    else add(m.label, 2);
    if (m.time) add('Time: ' + m.time.total + ' (' + (m.time.status.type === 'ok' ? 'within target' : m.time.status.type + ' by ' + m.time.status.diff) + ')' + (r.time ? ' · target ' + r.time.min + ' – ' + r.time.max : ''));
    m.activities.forEach(function (a) {
      add('• ' + a.title + (a.tool ? ' (' + a.tool + ')' : '') + (a.stage ? ' — ' + stageText13_(a) + ' · ' + a.pct + '%' : '') + (a.link === 'missing' ? ' — NO settings table link' : ''));
    });
  });
  if (r.finalized) {
    add('Settings table tabs', 2);
    add(r.finalized.ticked + ' of ' + r.finalized.counted + ' tabs marked finalized.');
    r.finalized.tabs.forEach(function (t) { add('• ' + t.name + ': ' + (t.state === 'none' ? 'no finalize box found' : t.state)); });
  }
  if (r.media && r.media.found) {
    add('Media deliverables', 2);
    add(r.media.flagged.length ? 'Completed with a Final Link but not ticked “Link Added to Dev Tab”:' : 'No completed media is waiting to be added to the Dev tab.');
    r.media.flagged.forEach(function (i) { add('• ' + i.type + ' — ' + i.title); });
    r.media.reverse.forEach(function (i) { add('• Ticked but no Final Link: ' + i.type + ' — ' + i.title); });
    Object.keys(r.media.byStatus).forEach(function (k) { if (k !== 'Cancelled') add(k + ': ' + r.media.byStatus[k]); });
  }
  r.notes.forEach(function (n) { add('Note: ' + n); });
  return L;
}

function stageText13_(a) {
  return a.mostly ? 'Mostly ' + a.mostly + (a.waiting && a.waiting.length ? ' (waiting on: ' + a.waiting.join('; ') + ')' : '') : STAGE_LABEL_13[a.stage];
}

function themeGroups13_(r) {
  if (!r.themes) return [];
  return r.themes.list.concat(r.themes.other ? [r.themes.other] : []);
}

/**
 * Plain text for QuickBase time entries: no headings, bullets, symbols, tables
 * or bars; one fact per line. short = course totals, themes and modules only.
 */
function plainLines13_(r, short) {
  var L = [], head = function (t) { if (L.length) L.push(''); L.push(String(t).toUpperCase()); };
  var mediaWaiting = r.media && r.media.found ? r.media.flagged.length : 0;
  head('Development Status Check');
  L.push(r.blueprint + ', checked ' + r.when + '.');

  head('Course totals');
  if (!r.checks.stage && !r.checks.links) L.push('Not checked.');
  if (r.checks.stage) {
    L.push(r.totals.ready + ' of ' + r.totals.activities + ' activities are Ready for the Dev Team.');
    Object.keys(r.totals.mostly || {}).forEach(function (k) { L.push(r.totals.mostly[k] + ' activities are mostly ' + k + ' and waiting on a few pieces.'); });
  }
  if (r.checks.links) L.push('Settings table links: ' + r.totals.linksPresent + ' of ' + r.totals.linksExpected + '.');
  if (mediaWaiting) L.push(mediaWaiting + ' completed media items are not yet added to the Dev tab.');

  head('By activity theme');
  var groups = themeGroups13_(r);
  if (!r.checks.stage) L.push('Not checked.');
  else if (!groups.length) L.push('No activity themes selected.');
  groups.forEach(function (g) {
    L.push(g.name + ' is ' + g.pct + ' percent complete (' + g.ready + ' of ' + g.count + ' ready): ' + g.acts.map(function (x) { return x.moduleShort + ' ' + x.pct + ' percent'; }).join(', ') + '.');
  });

  head('By module');
  r.modules.forEach(function (m) {
    L.push(m.label + (r.checks.stage ? ' is ' + m.pct + ' percent complete (' + m.ready + ' of ' + m.activities.length + ' ready)' : '') + (m.time ? ', time ' + m.time.total + ' (' + (m.time.status.type === 'ok' ? 'within target' : m.time.status.type + ' by ' + m.time.status.diff) + ')' : '') + '.');
  });
  if (!r.modules.length) L.push('No modules found.');

  if (!short) {
    r.modules.forEach(function (m) {
      head(m.label + ' activities');
      if (!m.activities.length) L.push('None.');
      m.activities.forEach(function (a) {
        L.push(a.title + (a.tool ? ' (' + a.tool + ')' : '') + (a.stage ? ': ' + stageText13_(a) + ', ' + a.pct + ' percent' : '') + (a.link === 'missing' ? ', no settings table link' : '') + '.');
      });
    });

    head('Home and syllabus');
    if (!r.home) L.push('Not checked.');
    else {
      var hs = [['Home', r.home, false], ['Syllabus', r.syllabus, true]];
      hs.forEach(function (x) {
        var bad = !x[1] || !x[1].found || x[1].untouched;
        L.push((bad ? 'Incomplete. ' : '') + x[0] + ': ' + regionText13_(x[1], x[2]).replace(/ \u00b7 /g, ', ') + '.');
      });
      L.push('Instructor Guide: ' + r.resources.guide + '. Course Suggestions: ' + r.resources.suggestions + '.');
    }

    head('Settings table tabs');
    if (!r.finalized) L.push('Not checked.');
    else L.push((r.finalized.ticked === 0 ? 'Incomplete. ' : '') + 'Marked finalized: ' + r.finalized.ticked + ' of ' + r.finalized.counted + '.');

    head('Media deliverables');
    if (!r.media) L.push('Not checked.');
    else if (!r.media.found) L.push(('Not checked. ' + (r.media.reason || '')).trim());
    else if (!r.media.flagged.length) L.push('None waiting.');
    else r.media.flagged.forEach(function (i) { L.push('Not yet added to the Dev tab: ' + i.type + ', ' + i.title + '.'); });

    head('Notes');
    if (!r.notes.length) L.push('None.');
    r.notes.forEach(function (n) { L.push(n); });
  }
  return L;
}

function regionText13_(reg, isSyllabus) {
  if (!reg || !reg.found) return 'section not found';
  return STAGE_LABEL_13[reg.stage] + (isSyllabus && reg.format ? ' · ' + reg.format : '');
}
