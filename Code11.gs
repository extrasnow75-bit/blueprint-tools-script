// ============================================================
// Blueprint Tools — Code11.gs
// Deploy Settings Tables
// Matches each activity in the Development tab to a tab in the project's
// settings table (a Google Sheet) and writes a link to that tab onto the
// activity's tool line (the "Link to settings tab" phrase).
// ------------------------------------------------------------
// Last updated on 2026-11-03 at 08:30 MST
// ------------------------------------------------------------
//
// Spec: docs/dev-handoff/HANDOFF.md section 3.2. Mockup: mockups/DeployB.png
// (by module). UI: Sidebar11.html.
//
// CONTRACT. Nothing is written until the user has pressed "Preview links" and
// then "Add links". Apply re-scans the Development tab (nothing is carried over
// from the preview except the user's picks), links only activities that have a
// tab picked, and never replaces a link that is already there unless "Replace
// links" is ticked. Only the "Link to settings tab" words get the link; no other
// text in the blueprint changes. The settings table itself is only READ.
//
// GAS FLAT NAMESPACE. Every name in this file ends in 11 (or 11_ for private
// helpers). Reuses, does NOT redefine:
//   collectTabs, TOOL_MARKER, CANVAS_TOOL_OPTIONS            (Code.gs)
//   getDevelopmentTabBody, extractDocId                      (Code2.gs)
//   scanDevStructure9_, matchKey9_, normTitle9_              (Code9.gs)
//   dashboardTab10_, findDashLabel10_, LINKDOCS_SLOTS_10     (Code10.gs)
//   pub10_                                                   (Code10.gs)
//
// SCOPES. Reads tab names with SpreadsheetApp.openById, which needs the Sheets
// scope https://www.googleapis.com/auth/spreadsheets. Greg approved it on
// 2026-11-03 (full scope, because Draft Settings Tables needs write access).
//
// FLAGS. Open questions are marked "DEFAULT:" in a comment.
// ============================================================


// ── CONSTANTS ────────────────────────────────────────────────

// The words that get the link. Same phrase Code.gs writes (TOOL_SUFFIX).
var LINK_PHRASE_11 = 'Link to settings tab';

// Tabs that are never matched: the Course Settings tab and any template tab.
var EXCLUDED_TAB_RE_11 = /^\s*course\s*settings\s*$|template/i;

// Ways a settings tab is typically named for each Canvas tool. Compared with
// normTitle9_ output. Used only to SUGGEST a tab, never to match silently.
var TOOL_TAB_NAMES_11 = {
  'Assignment':              ['assignment'],
  'Assignment (Not Graded)': ['assignment', 'ungraded assignment', 'not graded assignment'],
  'Discussion':              ['discussion'],
  'Quiz (Classic)':          ['classic quiz', 'quiz classic', 'quiz'],
  'Quiz (New)':              ['new quiz', 'quiz new', 'quiz']
};


// ── MENU ENTRY ───────────────────────────────────────────────

function showDeploySidebar11() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar11')
    .setTitle('Deploy Settings Tables')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}


// ── SIDEBAR API ──────────────────────────────────────────────

/**
 * On open: looks for the settings table the Dashboard already links. If there
 * is one, loads its tabs. Never throws for a missing link; the sidebar then asks
 * the user to paste one.
 */
function getDeployInit11() {
  var doc = DocumentApp.getActiveDocument();
  var out = { found: false, url: null, table: null, error: null, devOk: !!getDevelopmentTabBody(doc) };
  var url = null;
  try {
    var slot = null;
    for (var i = 0; i < LINKDOCS_SLOTS_10.length; i++) if (LINKDOCS_SLOTS_10[i].key === 'settings') slot = LINKDOCS_SLOTS_10[i];
    var lab = findDashLabel10_(dashboardTab10_(doc).body, slot.dash);
    url = lab && lab.existing ? lab.existing : null;
  } catch (e) { /* fall through to "choose a file" */ }
  if (!url) return out;
  out.url = url;
  try {
    out.table = loadSettingsTable11(url);
    out.found = true;
  } catch (e) {
    out.error = String(e && e.message || e);
  }
  return out;
}

/**
 * Reads the settings table's tabs. Sidebar-callable ("Choose a different file").
 * @param {string} urlOrId
 * @returns {{id, name, url, tabs: Array<{gid, name, hidden, excluded}>}}
 */
function loadSettingsTable11(urlOrId) {
  var id = extractDocId(String(urlOrId || '').trim());
  if (!id) throw new Error('That does not look like a Google Sheets link. Paste the settings table’s address.');
  var ss;
  try { ss = SpreadsheetApp.openById(id); }
  catch (e) { throw new Error('Could not open that file as a Google Sheet. Check that you can open it and that it is a Sheet, not an Excel file. (' + (e && e.message || e) + ')'); }
  var tabs = ss.getSheets().map(function (sh) {
    var name = sh.getName();
    return { gid: String(sh.getSheetId()), name: name, hidden: sh.isSheetHidden(), excluded: EXCLUDED_TAB_RE_11.test(name) };
  });
  return { id: id, name: ss.getName(), url: 'https://docs.google.com/spreadsheets/d/' + id + '/edit', tabs: tabs };
}

/**
 * Preview. Read-only. Scans the Development tab, matches every activity to a
 * settings tab, and reports what Apply would do.
 * @param {{tableId: string}} state
 */
function previewDeploy11(state) {
  var doc = DocumentApp.getActiveDocument();
  var table = loadSettingsTable11(state.tableId);
  var acts = scanActivities11_(doc);
  var usable = table.tabs.filter(function (t) { return !t.excluded && !t.hidden; });
  var keys = usable.map(function (t) { return { tab: t, key: matchKey9_(t.name) }; });
  var usedGids = {};

  var modules = [];
  acts.forEach(function (m) {
    var mod = { label: m.label, activities: [] };
    m.activities.forEach(function (a) {
      var row = { id: a.id, title: a.title, tool: a.tool || '', status: '', gid: '', how: '', note: '', existing: a.existingUrl || '' };
      if (a.isPage) {
        row.status = 'page'; row.note = 'Page activities have no settings tab.';
      } else if (!a.phraseFound) {
        row.status = 'nophrase';
        row.note = 'No “' + LINK_PHRASE_11 + '” words on this tool line, so there is nothing to link.';
      } else {
        var guess = matchTab11_(a, keys);
        row.gid = guess.tab ? guess.tab.gid : '';
        row.how = guess.how;      // 'exact' | 'suggested' | 'none'
        row.note = guess.note;
        row.status = guess.tab ? guess.how : 'none';
        if (guess.tab) usedGids[guess.tab.gid] = true;
      }
      mod.activities.push(row);
    });
    modules.push(mod);
  });

  var unused = usable.filter(function (t) { return !usedGids[t.gid]; }).map(function (t) { return t.name; });
  var existingCount = 0;
  modules.forEach(function (m) { m.activities.forEach(function (r) { if (r.existing) existingCount++; }); });

  return {
    table: { id: table.id, name: table.name, url: table.url, tabs: usable.map(function (t) { return { gid: t.gid, name: t.name }; }), skipped: table.tabs.filter(function (t) { return t.excluded || t.hidden; }).map(function (t) { return t.name; }) },
    modules: modules,
    unusedTabs: unused,
    existingCount: existingCount
  };
}

/**
 * Apply. Re-scans the document, then writes links only for the picks the user
 * confirmed.
 * @param {{tableId: string, picks: Object<string,string>, titles: Object<string,string>, replace: boolean}} state
 *   picks: activity id -> tab gid ('' = leave unlinked). titles: activity id ->
 *   title the preview showed, used to refuse a stale preview.
 */
function applyDeploy11(state) {
  var doc = DocumentApp.getActiveDocument();
  var table = loadSettingsTable11(state.tableId);
  var gidOk = {};
  table.tabs.forEach(function (t) { if (!t.excluded) gidOk[t.gid] = true; });

  var acts = scanActivities11_(doc);
  var res = { added: 0, replaced: 0, same: 0, keptExisting: 0, skipped: 0, problems: [] };

  acts.forEach(function (m) {
    m.activities.forEach(function (a) {
      var gid = state.picks && state.picks[a.id];
      if (!gid) return;
      if (state.titles && state.titles[a.id] !== a.title) {
        res.skipped++; res.problems.push(m.label + ' › ' + a.title + ': the blueprint changed since the preview. Run the preview again.');
        return;
      }
      if (!gidOk[gid]) { res.skipped++; res.problems.push(a.title + ': that tab is no longer in the settings table.'); return; }
      if (a.isPage || !a.phraseFound) { res.skipped++; return; }
      var url = table.url + '#gid=' + gid;
      if (a.existingUrl) {
        if (sameUrl10_(a.existingUrl, url) && String(a.existingUrl).indexOf('gid=' + gid) !== -1) { res.same++; return; }
        if (!state.replace) { res.keptExisting++; return; }
      }
      try {
        a._el.setLinkUrl(a._s, a._e, url);
        if (a.existingUrl) res.replaced++; else res.added++;
      } catch (e) {
        res.skipped++; res.problems.push(a.title + ': ' + (e && e.message || e));
      }
    });
  });
  return res;
}


// ── SCAN THE DEVELOPMENT TAB ─────────────────────────────────

/**
 * Every activity in every numbered module, with its tool and tool-line link.
 * Private fields (_el/_s/_e) are for Apply only. Ids are "<module#>.<slot#>",
 * which is stable between Preview and Apply when the document has not changed.
 */
function scanActivities11_(doc) {
  var devBody = getDevelopmentTabBody(doc);
  if (!devBody) throw new Error('Could not find a “Development” tab in this blueprint.');
  var mods = scanDevStructure9_(devBody), out = [];
  mods.forEach(function (m) {
    var entry = { label: m.displayLabel || ('Module ' + m.num), activities: [] };
    m.activities.forEach(function (slot, si) {
      var info = readToolLine11_(devBody, slot.h4);
      entry.activities.push({
        id: m.num + '.' + (si + 1),
        title: slot.title,
        tool: info.tool,
        isPage: info.tool === 'Page',
        phraseFound: info.found,
        existingUrl: info.existingUrl,
        _el: info.el, _s: info.s, _e: info.e
      });
    });
    out.push(entry);
  });
  return out;
}

/**
 * Finds the tool line under an activity heading (same rules as setNearbyTool in
 * Code.gs: within 8 paragraphs, stopping at the next heading) and the
 * "Link to settings tab" words in it.
 * @returns {{tool: string, found: boolean, existingUrl: string|null, el, s, e}}
 */
function readToolLine11_(body, h4) {
  var out = { tool: '', found: false, existingUrl: null, el: null, s: 0, e: 0 };
  var H = DocumentApp.ParagraphHeading;
  var start = body.getChildIndex(h4), n = body.getNumChildren();
  for (var i = start + 1; i < Math.min(start + 8, n); i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para = child.asParagraph(), h = para.getHeading();
    if (h === H.HEADING2 || h === H.HEADING3 || h === H.HEADING4) break;
    var text = para.getText(), trimmed = text.trim();
    var isToolLine = text.indexOf(TOOL_MARKER) !== -1 || text.indexOf(LINK_PHRASE_11) !== -1 || CANVAS_TOOL_OPTIONS.indexOf(trimmed) !== -1;
    if (!isToolLine) continue;
    // Tool name = text before the marker or the old "; " separator.
    var cut = text.search(new RegExp('\\s*(?:' + TOOL_MARKER + '|;)'));
    out.tool = (cut === -1 ? trimmed : text.slice(0, cut)).trim();
    var at = text.indexOf(LINK_PHRASE_11);
    if (at !== -1) {
      var t = para.editAsText();
      out.found = true; out.el = t; out.s = at; out.e = at + LINK_PHRASE_11.length - 1;
      out.existingUrl = t.getLinkUrl(at) || null;
    }
    return out;
  }
  return out;
}


// ── MATCHING ─────────────────────────────────────────────────

/**
 * Picks a settings tab for one activity. Many activities may share one tab.
 * Strong: the title equals the tab name (ignoring case, punctuation, numbers) or
 * the part of the title before a colon does. Suggested (the user is asked to
 * confirm): a tab name that opens the title, shared keywords, or a tab named for
 * the Canvas tool. No match is a normal outcome and becomes a dropdown.
 * @param {Object} a  activity from scanActivities11_
 * @param {Array<{tab, key}>} keys  usable tabs with matchKey9_ forms
 * @returns {{tab: Object|null, how: string, note: string}}
 */
function matchTab11_(a, keys) {
  var t = matchKey9_(a.title);
  var head = matchKey9_(String(a.title).split(':')[0]);
  var best = null;   // {tab, rank, len, how, note}

  function offer(tab, rank, how, note) {
    var len = tab.name.length;
    if (!best || rank < best.rank || (rank === best.rank && len > best.len)) best = { tab: tab, rank: rank, len: len, how: how, note: note };
  }

  keys.forEach(function (k) {
    var kk = k.key;
    if (kk.raw === t.raw || (kk.norm && kk.norm === t.norm) || (kk.base && kk.base === t.base)) {
      offer(k.tab, 1, 'exact', 'Matched on the activity title.');
    } else if (kk.norm && head.norm && kk.norm === head.norm) {
      offer(k.tab, 2, 'exact', 'Matched on the title before the colon.');
    } else if (kk.base && kk.base.length >= 4 && t.base.indexOf(kk.base + ' ') === 0) {
      offer(k.tab, 3, 'suggested', 'The tab name opens the activity title. Please confirm.');
    } else {
      var score = keywordScore9_(kk, t);
      if (score >= KEYWORD_MIN_SCORE_9) offer(k.tab, 4, 'suggested', 'Matched on a shared keyword. Please confirm.');
    }
  });

  if (!best) {
    var names = TOOL_TAB_NAMES_11[a.tool] || [];
    for (var i = 0; i < keys.length; i++) {
      if (names.indexOf(keys[i].key.norm) !== -1) { offer(keys[i].tab, 5, 'suggested', 'Suggested from the tool name. Please confirm.'); break; }
    }
  }
  return best ? { tab: best.tab, how: best.how, note: best.note } : { tab: null, how: 'none', note: 'No match found. Pick a tab or leave it unlinked.' };
}
