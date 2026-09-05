// ============================================================
// Blueprint Tools — Code9.gs
// Design Map → Dev Tab: copies the Design tab's Course Design Map into the
// Development tab — module titles into the H2 headings, the CLO/MLO row into
// each Module Overview as a real numbered list, and every other row's notes
// under that activity's "Directions go here…" placeholder.
// ------------------------------------------------------------
// Last updated on 2026-09-05 at 00:47 MDT
// ------------------------------------------------------------
//
// Runs AFTER "Add Activity Titles, Tools, Due Date Headers, & Times", which is
// what creates the H2 module headings and the H4 activity slots this tool
// writes into.
//
// Replaces the retired "Add Module Titles & Module Dates (Beta)" tool
// (Code7.gs / Sidebar7.html). Its title half moved here UNCHANGED in behavior —
// same radio-per-module UI, same never-overwrite contract; its dates half became
// the standalone "Specialty Tool: Add Module Dates" (Code8.gs / Sidebar8.html).
// See project memory project_blueprint_designmap_dev_tab.md, sections 3 and 4.
//
// GAS FLAT NAMESPACE. Every .gs file shares one global scope, so a function or
// var defined twice does not error — the later definition silently wins. These
// three came from Code7.gs and are used by Code8.gs as well; this file is their
// ONE home, and Code8.gs must not redefine them:
//   MODULE_PREFIX_RE_7, classifyHeading7_, scanDevelopmentHeadings7_
//
// Relies on shared helpers defined elsewhere in that same namespace — do NOT
// redefine any of these here:
//   collectTabs, zeroIndent_, RIGHT_INDENT                    (Code.gs)
//   getDevelopmentTabBody, stripActivityHeading,
//   DIRECTIONS_PLACEHOLDER_TEXT, restartCopiedListNumbering_,
//   freshListId_                                              (Code2.gs)
//   START_PLACEHOLDER_7, END_PLACEHOLDER_7                    (Code8.gs)
// ============================================================


// ── CONSTANTS ────────────────────────────────────────────────

// The load-bearing contract with the Blueprint template: the tool writes a
// field ONLY where its placeholder is still present, which is what makes
// "skip and report" work without ever clobbering a designer's own text.
var TITLE_PLACEHOLDER_7 = 'Title';

// Matches a module heading prefix in any of the four accepted forms:
//   "Module 1:" / "Week 1:" / "Module 01:" / "Week 01:"
// Group 1 is the label word, group 2 the zero padding, group 3 the number.
// The padding is captured separately so the document's own form can be
// preserved on write — a doc that says "Module 01" must not silently become
// "Module 1".
//
// The colon is optional. A heading a designer has rewritten by hand ("Module 5
// — Vectors & Data") has no colon, and requiring one made classifyHeading7_
// return null for it, so the module vanished from the sidebar entirely instead
// of being listed as untouchable. Matching it lets it be reported.
var MODULE_PREFIX_RE_7 = /^(module|week)\s+(0*)(\d+)\s*:?\s*/i;

// Same idea, but for a Course Design Map's module header cell. This supersedes
// Code7.gs's CDM_MODULE_RE_7 and the archived AI parser's /^module\s+\d+/i:
//   • "Week" and zero padding are accepted — the old AI-tools regex found ZERO
//     modules in a Week-based Blueprint.
//   • The separator is captured (group 4) rather than skipped, because it is
//     what distinguishes a header from an activity — see cdmModuleHeader9_.
// Groups: 1 label, 2 zero padding, 3 number, 4 separator (may be absent),
//         5 everything after it.
var CDM_MODULE_RE_9 =
  /^(module|week)\s+(0*)(\d+)\s*([:–—-])?\s*([\s\S]*)$/i;

// Typed list markers at the head of a plain paragraph in the CLO/MLO cell
// ("1. ", "2) ", "• ", "– "). Real Docs list items never carry one; when the
// designer typed the numbers by hand, the glyph this tool adds would otherwise
// render as "1. 1. Explain …".
var CDM_TYPED_MARKER_RE_9 = /^\s*(?:\d+\s*[.)\]]|[•●▪–—-])\s+/;

// Cyan, the IDC convention for "this is a note to myself, delete before
// hand-off". Applied to pasted ACTIVITY NOTES only — the objectives are real
// course content that stays, so highlighting them would invert the signal.
var NOTE_HIGHLIGHT_9 = '#00ffff';

// Separator for the composite keys the sidebar sends back. "|" cannot appear in
// a module number, and an activity label containing one would only ever collide
// with itself.
var KEY_SEP_9 = '||';


// ── SIDEBAR OPENER ───────────────────────────────────────────

function showDesignMapSidebar9() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar9')
    .setTitle('Design Map → Dev Tab')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}


// ── HEADING STATE ────────────────────────────────────────────
// Moved here from Code7.gs unchanged. Code8.gs calls both of these; this file
// is their only definition.

/**
 * Classifies one Development-tab H2 heading.
 *
 * There is no provenance marker in the document — nothing distinguishes text
 * this tool wrote from text a designer typed — so the shape of the heading is
 * the only signal available. That is sufficient: the tool writes a field only
 * where its placeholder survives, so misreading a hand-written heading as
 * "filled" is safe (it gets skipped and reported, never overwritten).
 *
 * @param {string} text  raw heading text
 * @returns {Object|null}  null when the heading is not a numbered module
 */
function classifyHeading7_(text) {
  var raw   = String(text || '').trim();
  var match = raw.match(MODULE_PREFIX_RE_7);
  if (!match) return null;

  var label   = match[1];
  var padding = match[2];
  var num     = parseInt(match[3], 10);
  var rest    = raw.slice(match[0].length).trim();

  // Split off a trailing parenthetical, which is where the dates live.
  var parenMatch = rest.match(/\(([^()]*)\)\s*$/);
  var datePart   = parenMatch ? parenMatch[1].trim() : null;
  var titlePart  = parenMatch ? rest.slice(0, parenMatch.index).trim() : rest;

  var titleIsPlaceholder = titlePart === TITLE_PLACEHOLDER_7;
  var dateIsPlaceholder  = datePart !== null &&
    new RegExp('^' + START_PLACEHOLDER_7 + '\\s*-\\s*' + END_PLACEHOLDER_7 + '$', 'i').test(datePart);

  var state;
  if (datePart === null)                            state = 'freeform';
  else if (titleIsPlaceholder && dateIsPlaceholder) state = 'pristine';
  else if (dateIsPlaceholder)                       state = 'titleOnly';
  else if (titleIsPlaceholder)                      state = 'datesOnly';
  else                                              state = 'filled';

  return {
    num:                num,
    label:              label,
    padding:            padding,
    // Reassembled exactly as the document has it, for display in the sidebar.
    displayLabel:       label + ' ' + padding + match[3],
    rawHeading:         raw,
    titlePart:          titlePart,
    datePart:           datePart,
    titleIsPlaceholder: titleIsPlaceholder,
    dateIsPlaceholder:  dateIsPlaceholder,
    state:              state
  };
}


/**
 * Walks the Development tab and returns one entry per numbered module heading,
 * in document order. Non-numbered H2s (Course Resources, etc.) are skipped and
 * never touched — matching how the Time Estimator treats them.
 */
function scanDevelopmentHeadings7_(devBody) {
  var found = [];
  var n     = devBody.getNumChildren();
  var H2    = DocumentApp.ParagraphHeading.HEADING2;

  for (var i = 0; i < n; i++) {
    var child = devBody.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    var para = child.asParagraph();
    if (para.getHeading() !== H2) continue;

    var info = classifyHeading7_(para.getText());
    if (!info) continue;

    info.childIndex = i;
    found.push(info);
  }
  return found;
}


// ── COURSE DESIGN MAP PARSER ─────────────────────────────────

/**
 * Decides whether a Course Design Map row's far-left cell is a module header.
 *
 * The far-left column does double duty — it is both the row label and the
 * activity name — so "Module 1 Quiz" as an activity would be read as a header
 * by a bare /^module\s+\d+/ and would silently restart the module block. Three
 * signals are available and any ONE of them is enough to call it a header:
 *   • nothing follows the number            → "Module 1", "Week 07"
 *   • a separator follows it                → "Module 1: Matrices", "Week 5 — Vectors"
 *   • the Notes column beside it is empty   → a merged/blue header band
 * Only the combination "no separator + trailing words + notes of its own"
 * is rejected, which is exactly the activity-row shape.
 *
 * @param {string} firstCell  text of the far-left cell
 * @param {string} valueText  text of the Notes cell beside it ('' when absent)
 * @returns {Object|null}
 */
function cdmModuleHeader9_(firstCell, valueText) {
  var match = String(firstCell || '').trim().match(CDM_MODULE_RE_9);
  if (!match) return null;

  var separator = !!match[4];
  var rest      = String(match[5] || '').trim();

  if (!separator && rest && valueText) return null;

  return {
    num:          parseInt(match[3], 10),
    // Padding preserved, same reasoning as classifyHeading7_'s displayLabel:
    // a map that says "Module 01" should read back as "Module 01".
    displayLabel: match[1] + ' ' + match[2] + match[3],
    title:        rest,
    separator:    separator
  };
}


/**
 * Reads the Design tab's Course Design Map into one record per module.
 *
 * Merges what the retired tools read separately: Code7.gs's
 * parseDesignMapTitles7_ (both title sources) and the archived AI tools'
 * parseCourseDesignMap (CLO/MLO plus the activity rows).
 *
 * Each module block is a two-column run of rows. The far-left cell is the row
 * label AND the activity name; the right-hand cell is the Notes column. Rows
 * are classified by their label:
 *   "Title"          → the module title (second of two title sources)
 *   "CLOs, MLOs"     → objectives, destined for the Module Overview
 *   anything else    → an activity, whose notes go under its Directions line
 *
 * Cells are held as CELL REFERENCES, never as getText(). Both objectives and
 * notes carry hyperlinks and multi-paragraph structure that text would drop.
 *
 * @param {GoogleAppsScript.Document.Body} designBody
 * @returns {{byNumber: Object, order: number[]}}
 */
function parseDesignMap9_(designBody) {
  var byNumber = {};
  var order    = [];
  var tables   = designBody.getTables();

  for (var t = 0; t < tables.length; t++) {
    var table   = tables[t];
    var numRows = table.getNumRows();
    if (numRows < 2) continue;

    // Identify the Course Design Map by a "CLO"/"MLO" label in the first
    // several rows. Cell reads are server round-trips, so the detection scan is
    // bounded and its text is cached for the parse loop below.
    var rowCache  = {};
    var isCDM     = false;
    var scanLimit = Math.min(numRows, 15);

    for (var ri = 0; ri < scanLimit && !isCDM; ri++) {
      var scanRow = table.getRow(ri);
      var scanN   = scanRow.getNumCells();
      var rowText = [];
      for (var ci = 0; ci < scanN && ci < 2; ci++) {
        var cellStr = scanRow.getCell(ci).getText();
        rowText.push(cellStr);
        var lc = cellStr.toLowerCase();
        if (lc.indexOf('clo') !== -1 || lc.indexOf('mlo') !== -1) isCDM = true;
      }
      rowCache[ri] = rowText;
    }
    if (!isCDM) continue;

    var current = null;

    for (var r = 0; r < numRows; r++) {
      var row   = table.getRow(r);
      var cells = rowCache[r];
      if (!cells) {
        var dataN = row.getNumCells();
        cells = [];
        for (var c = 0; c < dataN && c < 2; c++) cells.push(row.getCell(c).getText());
      }
      if (cells.length === 0) continue;

      var firstCell = String(cells[0] || '').trim();
      var valueText = String(cells[1] || '').trim();
      var hasValue  = cells.length > 1;

      var header = cdmModuleHeader9_(firstCell, valueText);
      if (header) {
        current = byNumber[header.num];
        if (!current) {
          current = byNumber[header.num] = {
            num:            header.num,
            displayLabel:   header.displayLabel,
            headerTitle:    '',
            rowTitle:       '',
            objectivesCell: null,
            activities:     []
          };
          order.push(header.num);
        }
        // The title normally sits in the header cell itself. A map that splits
        // it across the two columns ("Module 2:" | "Matrices as Datasets") is a
        // real layout too, and the separator is what makes it unambiguous.
        if (header.title)                                current.headerTitle = header.title;
        else if (header.separator && hasValue && valueText) current.headerTitle = valueText;
        continue;
      }

      // Every remaining row needs a label AND a value column to be useful.
      if (!current || !firstCell || !hasValue) continue;

      var label = firstCell.toLowerCase();

      if (label === 'title') {
        current.rowTitle = valueText;
        continue;
      }

      if (label.indexOf('clo') !== -1 || label.indexOf('mlo') !== -1) {
        if (valueText) current.objectivesCell = row.getCell(1);
        continue;
      }

      current.activities.push({
        label:    firstCell,
        cell:     valueText ? row.getCell(1) : null,
        hasNotes: !!valueText
      });
    }

    if (order.length > 0) break; // found the CDM — stop searching
  }

  order.sort(function (a, b) { return a - b; });
  return { byNumber: byNumber, order: order };
}


// ── DEVELOPMENT TAB STRUCTURE ────────────────────────────────

/**
 * One linear pass over the Development tab that returns everything this tool
 * needs to write: each module's Overview region and each of its activity slots,
 * with the anchor element and any content already sitting in the destination.
 *
 * A pass rather than per-slot lookups on purpose. Calling
 * findDirectionsPlaceholder (Code2.gs:933) for every slot costs ~11 document
 * round trips each; at 15 modules × 7 activities that is a four-figure call
 * count before the sidebar can even render.
 *
 * ELEMENT REFERENCES, not child indices, are what get stored. Indices go stale
 * the moment anything is inserted; a reference stays valid and its current
 * index can be read back with body.getChildIndex() immediately before a write.
 *
 * @param {GoogleAppsScript.Document.Body} devBody
 * @returns {Array<Object>} one entry per numbered module, in document order
 */
function scanDevStructure9_(devBody) {
  var H2     = DocumentApp.ParagraphHeading.HEADING2;
  var H3     = DocumentApp.ParagraphHeading.HEADING3;
  var H4     = DocumentApp.ParagraphHeading.HEADING4;
  var PARA   = DocumentApp.ElementType.PARAGRAPH;

  var modules  = [];
  var current  = null;   // module being walked
  var slot     = null;   // activity slot being walked
  var overview = null;   // Module Overview region being walked
  var n        = devBody.getNumChildren();

  for (var i = 0; i < n; i++) {
    var child = devBody.getChild(i);

    if (child.getType() !== PARA) {
      // Tables and list items are always real content in whichever region we
      // are standing in — this is how already-pasted notes get detected. Not
      // gated on having seen the placeholder: a slot whose directions were
      // already deployed has no placeholder left, and its content is very often
      // list items, which would otherwise read as an empty slot.
      if      (slot)     slot.content.push(child);
      else if (overview) overview.content.push(child);
      continue;
    }

    var para    = child.asParagraph();
    var heading = para.getHeading();
    var text    = para.getText();
    var trimmed = text.trim();

    if (heading === H2) {
      slot = null; overview = null;
      var info = classifyHeading7_(trimmed);
      if (info) {
        current = {
          num:          info.num,
          displayLabel: info.displayLabel,
          info:         info,
          overview:     null,
          activities:   []
        };
        modules.push(current);
      } else {
        current = null;   // Course Resources and friends — never touched
      }
      continue;
    }

    if (!current) continue;

    if (heading === H4) {
      overview = null;
      slot = {
        rawTitle:    trimmed,
        title:       stripActivityHeading(trimmed),
        h4:          para,
        placeholder: null,
        content:     []
      };
      current.activities.push(slot);
      continue;
    }

    if (heading === H3) {
      slot = null;
      // "Module 3 Overview" opens the overview region; every other H3 in a
      // module is a due-by header, which closes it.
      if (/overview\s*$/i.test(trimmed)) {
        overview = { h3: para, refPara: null, anchor: para, content: [] };
        current.overview = overview;
      } else {
        overview = null;
      }
      continue;
    }

    // ── Normal paragraph ──
    if (slot) {
      // Slot preamble, never content. Same two tests readModuleContent_
      // (Code2.gs) uses, and for the same reason: "Estimated time:" precedes
      // the tool line and matches neither a blank nor a heading.
      if (/^estimated time/i.test(trimmed))       continue;
      if (/link to settings tab$/i.test(trimmed)) continue;
      if (trimmed === '')                         continue;

      if (!slot.placeholder &&
          (text === DIRECTIONS_PLACEHOLDER_TEXT || text === 'Directions go here...')) {
        slot.placeholder = para;
        continue;
      }
      slot.content.push(para);
      continue;
    }

    if (overview) {
      if (!overview.refPara && /^\[?\s*refer to the template/i.test(trimmed)) {
        overview.refPara = para;
        // Objectives go after the refer-to line when it is there. When a
        // designer has deleted it, the Overview heading is the next best
        // anchor — better than refusing to write the module at all.
        overview.anchor  = para;
        continue;
      }
      if (trimmed === '') continue;
      overview.content.push(para);
      continue;
    }
  }

  return modules;
}


// ── ACTIVITY MATCHING ────────────────────────────────────────

/**
 * Normalizes a title for comparison: case, curly quotes, "&", and every run of
 * punctuation collapse away, so "Readings & Multimedia" and "Readings and
 * multimedia" land on the same string.
 */
function normTitle9_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * normTitle9_ with a leading or trailing bare index removed, which is how the
 * Course Design Map distinguishes repeats of the same kind of row:
 * "Readings 1" / "Readings 2" both reduce to "readings".
 */
function normBase9_(s) {
  return normTitle9_(s)
    .replace(/^\d+(?:\s+|$)/, '')
    .replace(/(?:\s+|^)\d+$/, '')
    .trim();
}

/** Pre-computes every comparison form of one title, so the tiers stay cheap. */
function matchKey9_(title) {
  var raw = String(title || '').trim();
  return { raw: raw, lower: raw.toLowerCase(), norm: normTitle9_(raw), base: normBase9_(raw) };
}

// Ordered strongest → weakest. Each tier is tried across ALL still-unmatched
// design-map rows before the next one is considered, so a weak match can never
// steal a slot that a stronger one wants.
var MATCH_TIERS_9 = [
  { name: 'exact title',        fn: function (a, b) { return a.raw  === b.raw;  } },
  { name: 'ignoring case',      fn: function (a, b) { return a.lower === b.lower; } },
  { name: 'ignoring wording',   fn: function (a, b) { return !!a.norm && a.norm === b.norm; } },
  { name: 'ignoring numbering', fn: function (a, b) { return !!a.base && a.base === b.base; } },
  { name: 'by leading words',   fn: function (a, b) {
      // Word-boundaried on purpose: "read" must not claim "readings", and a
      // three-letter stem is too weak to be evidence of anything.
      if (!a.base || !b.base || a.base.length < 4 || b.base.length < 4) return false;
      return a.base.indexOf(b.base + ' ') === 0 || b.base.indexOf(a.base + ' ') === 0;
    } }
];

/**
 * Pairs one module's Course Design Map rows with its Development tab activity
 * slots. Deterministic and side-effect free, so the sidebar's preview and the
 * apply pass compute the same answer from the same document.
 *
 * A slot is claimed by at most one row. Rows that claim nothing come back as
 * -1 and are handed to the user as a dropdown.
 *
 * @param {Array<{label: string}>} cdmRows
 * @param {Array<{title: string}>} devActivities
 * @returns {{ pairs: number[], tiers: string[] }} index-aligned with cdmRows
 */
function matchActivities9_(cdmRows, devActivities) {
  var cdm = [], dev = [], pairs = [], tiers = [];
  var d, c, t;

  for (c = 0; c < cdmRows.length; c++) {
    cdm.push(matchKey9_(cdmRows[c].label));
    pairs.push(-1);
    tiers.push('');
  }
  for (d = 0; d < devActivities.length; d++) dev.push(matchKey9_(devActivities[d].title));

  var taken = {};

  for (t = 0; t < MATCH_TIERS_9.length; t++) {
    var tier = MATCH_TIERS_9[t];
    for (c = 0; c < cdm.length; c++) {
      if (pairs[c] !== -1) continue;
      for (d = 0; d < dev.length; d++) {
        if (taken[d]) continue;
        if (!tier.fn(cdm[c], dev[d])) continue;
        pairs[c] = d;
        tiers[c] = tier.name;
        taken[d] = true;
        break;
      }
    }
  }

  return { pairs: pairs, tiers: tiers };
}


// ── SIDEBAR DATA ─────────────────────────────────────────────

/**
 * Sidebar-callable. One read-only pass that produces everything the picker
 * needs: the title half's per-module options, the design-map payload per
 * module, which destinations already have content, and the match the apply
 * pass will make so the user can override it before anything is written.
 *
 * The Development tab drives the module list, since that is where every write
 * lands. Course Design Map modules with no matching heading are reported but
 * do not block the run.
 */
function getDesignMapSidebarData9() {
  var result = {
    modules:         [],   // title half — one per Development tab module heading
    conflicts:       [],
    unmatchedDesign: [],
    mapModules:      [],   // design-map payload, aligned with .modules by num
    devActivities:   [],   // union of activity titles, for the activity picker
    unmatchedRows:   0,
    error:           ''
  };

  try {
    var doc  = DocumentApp.getActiveDocument();
    var tabs = collectTabs(doc);

    var designTab = null;
    for (var i = 0; i < tabs.length; i++) {
      if (/\bdesign\b/i.test(tabs[i].title)) { designTab = tabs[i]; break; }
    }

    var devBody = getDevelopmentTabBody(doc);
    if (!devBody) {
      result.error = 'Could not find a "Development" tab in this document.';
      return result;
    }
    if (!designTab) {
      result.error = 'Could not find a "Design" tab in this document.';
      return result;
    }

    var devModules = scanDevStructure9_(devBody);
    if (devModules.length === 0) {
      result.error = 'No numbered module headings (Module 1, Week 1, …) were found ' +
                     'in the Development tab. Run "Add Activity Titles, Tools, Due Date ' +
                     'Headers, & Times" first.';
      return result;
    }

    var map  = parseDesignMap9_(designTab.body);
    var seen = {};
    var devTitleSeen = {};

    for (var m = 0; m < devModules.length; m++) {
      var devMod = devModules[m];
      var info   = devMod.info;
      var design = map.byNumber[devMod.num] ||
                   { headerTitle: '', rowTitle: '', objectivesCell: null, activities: [] };
      seen[devMod.num] = true;

      // ── Title half (ported unchanged from getModuleTitlesSidebarData7) ──
      // Match across tabs on the parsed integer, never the label string — a
      // "Module 1" Design Map row must pair with a "Week 01" Dev heading.
      var options = [];
      if (design.headerTitle) options.push({ origin: 'Course Design Map header', text: design.headerTitle });
      if (design.rowTitle && design.rowTitle !== design.headerTitle) {
        options.push({ origin: 'Course Design Map "Title" row', text: design.rowTitle });
      }
      if (info.state !== 'pristine' && info.titlePart && !info.titleIsPlaceholder) {
        options.push({ origin: 'currently in document', text: info.titlePart });
      }

      var isConflict = !!(design.headerTitle && design.rowTitle &&
                          design.headerTitle !== design.rowTitle);
      if (isConflict) {
        result.conflicts.push({
          module: info.displayLabel,
          header: design.headerTitle,
          row:    design.rowTitle
        });
      }

      result.modules.push({
        num:           info.num,
        displayLabel:  info.displayLabel,
        state:         info.state,
        rawHeading:    info.rawHeading,
        currentTitle:  info.titleIsPlaceholder ? '' : info.titlePart,
        canWriteTitle: info.titleIsPlaceholder,
        options:       options,
        conflict:      isConflict
      });

      // ── Design-map payload ──
      var devTitles = [];
      for (var a = 0; a < devMod.activities.length; a++) {
        var slotTitle = devMod.activities[a].title;
        devTitles.push(slotTitle);
        if (slotTitle && !devTitleSeen[slotTitle.toLowerCase()]) {
          devTitleSeen[slotTitle.toLowerCase()] = true;
          result.devActivities.push(slotTitle);
        }
      }

      var noteRows = [];
      for (var q = 0; q < design.activities.length; q++) {
        if (design.activities[q].hasNotes) noteRows.push(design.activities[q]);
      }

      var matched = matchActivities9_(noteRows, devMod.activities);
      var rows    = [];

      for (var k = 0; k < noteRows.length; k++) {
        var target = matched.pairs[k] === -1 ? null : devMod.activities[matched.pairs[k]];
        if (!target) result.unmatchedRows++;
        rows.push({
          label:         noteRows[k].label,
          key:           devMod.num + KEY_SEP_9 + noteRows[k].label,
          matchedTitle:  target ? target.title : '',
          // Which tier claimed it. The sidebar leaves confident matches alone
          // and only asks about the weakest one — surfacing a dozen certain
          // matches for confirmation would bury the few that need a decision.
          matchedHow:    matched.tiers[k],
          occupied:      !!target && target.content.length > 0
        });
      }

      var ov = devMod.overview;
      result.mapModules.push({
        num:          devMod.num,
        displayLabel: devMod.displayLabel,
        devTitles:    devTitles,
        objectives: {
          available: !!design.objectivesCell,
          hasAnchor: !!(ov && ov.anchor),
          noRefLine: !!(ov && !ov.refPara),
          occupied:  !!(ov && ov.content.length > 0)
        },
        rows: rows
      });
    }

    for (var key in map.byNumber) {
      if (!seen[key]) {
        var d = map.byNumber[key];
        result.unmatchedDesign.push({
          module: d.displayLabel || ('Module ' + key),
          title:  d.headerTitle || d.rowTitle || '(untitled)'
        });
      }
    }

    Logger.log('getDesignMapSidebarData9: %s module(s), %s conflict(s), %s unmatched design module(s), ' +
               '%s unmatched note row(s).',
               result.modules.length, result.conflicts.length,
               result.unmatchedDesign.length, result.unmatchedRows);

  } catch (e) {
    Logger.log('getDesignMapSidebarData9 error: ' + e.message);
    result.error = e.message;
  }

  return result;
}


// ── COPYING ──────────────────────────────────────────────────

// The character-level attributes worth carrying across from the Course Design
// Map. Deliberately NOT the whole map getAttributes() returns:
//   • paragraph keys (HEADING, INDENT_START, SPACING_*, LINE_SPACING) would drag
//     the map cell's own indents into a list whose indents this tool is at pains
//     to zero;
//   • FONT_FAMILY / FONT_SIZE / BACKGROUND_COLOR would import the Design tab's
//     look into the Development tab, where the objectives should read as native
//     body text.
// What is left is the formatting that carries meaning rather than styling.
var TEXT_ATTRS_9 = ['BOLD', 'ITALIC', 'UNDERLINE', 'STRIKETHROUGH',
                    'FOREGROUND_COLOR', 'LINK_URL'];

/** Narrows a getAttributes() map to the keys in TEXT_ATTRS_9. */
function textAttrs9_(attrs) {
  var out = {};
  if (!attrs) return out;
  for (var i = 0; i < TEXT_ATTRS_9.length; i++) {
    var key = DocumentApp.Attribute[TEXT_ATTRS_9[i]];
    if (attrs[key] !== undefined && attrs[key] !== null) out[key] = attrs[key];
  }
  return out;
}

/**
 * Rebuilds one source Text element inside a destination paragraph or list item,
 * one formatting run at a time.
 *
 * appendText alone drops hyperlinks — the destination inherits whatever the
 * preceding text had — and Course Design Map objectives routinely carry them.
 * getTextAttributeIndices() gives the run boundaries; each run is appended,
 * then its attributes and link are re-applied to exactly the range it occupies.
 *
 * @param {GoogleAppsScript.Document.Text} srcText
 * @param {GoogleAppsScript.Document.Paragraph|GoogleAppsScript.Document.ListItem} destEl
 */
function copyRuns9_(srcText, destEl) {
  var s = srcText.getText();
  if (!s) return;

  var dest = destEl.editAsText();
  var idxs = srcText.getTextAttributeIndices();
  if (idxs.length === 0 || idxs[0] !== 0) idxs = [0].concat(idxs);

  // Tracked locally rather than re-reading dest.getText().length each run —
  // that read is a server round trip, and this runs once per objective.
  var written = 0;

  for (var k = 0; k < idxs.length; k++) {
    var start = idxs[k];
    var stop  = (k + 1 < idxs.length) ? idxs[k + 1] : s.length;
    if (stop <= start) continue;

    var chunk = s.substring(start, stop);
    dest.appendText(chunk);

    var from = written;
    var to   = written + chunk.length - 1;
    written += chunk.length;

    dest.setAttributes(from, to, textAttrs9_(srcText.getAttributes(start)));

    // Belt and braces. LINK_URL is in the map above, but setAttributes has been
    // seen to no-op on it when the run is the first thing in an empty element.
    var url = srcText.getLinkUrl(start);
    if (url) dest.setLinkUrl(from, to, url);
  }
}


/**
 * Reads the CLO/MLO cell into one entry per objective.
 *
 * Both shapes occur in the wild: a real Docs list, and plain paragraphs with
 * the numbers typed by hand. `typed` records which, because only the second
 * needs its "1. " stripped before a real glyph is added.
 */
function objectiveSources9_(cell) {
  var out = [];
  var n   = cell.getNumChildren();

  for (var i = 0; i < n; i++) {
    var el   = cell.getChild(i);
    var type = el.getType();

    if (type === DocumentApp.ElementType.LIST_ITEM) {
      var li = el.asListItem();
      if (!li.getText().trim()) continue;
      out.push({ text: li.editAsText(), level: li.getNestingLevel(), typed: false });
    } else if (type === DocumentApp.ElementType.PARAGRAPH) {
      var p = el.asParagraph();
      if (!p.getText().trim()) continue;
      out.push({ text: p.editAsText(), level: 0, typed: true });
    }
  }
  return out;
}


/**
 * Reads a Notes cell into the elements to copy, with blank paragraphs trimmed
 * off both ends. Interior blanks are kept — they are the designer's own
 * paragraph breaks.
 */
function noteSources9_(cell) {
  var all = [];
  var n   = cell.getNumChildren();

  for (var i = 0; i < n; i++) {
    var el   = cell.getChild(i);
    var type = el.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH ||
        type === DocumentApp.ElementType.LIST_ITEM ||
        type === DocumentApp.ElementType.TABLE) {
      all.push(el);
    }
  }

  function isBlank(el) {
    var t = el.getType();
    if (t === DocumentApp.ElementType.PARAGRAPH) return !el.asParagraph().getText().trim();
    if (t === DocumentApp.ElementType.LIST_ITEM) return !el.asListItem().getText().trim();
    return false;
  }

  var start = 0, end = all.length - 1;
  while (start <= end && isBlank(all[start])) start++;
  while (end >= start && isBlank(all[end]))   end--;

  return all.slice(start, end + 1);
}


/** Paints the IDC self-note cyan over one inserted element. */
function highlightElement9_(el) {
  var type = el.getType();

  if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM) {
    var textEl = (type === DocumentApp.ElementType.PARAGRAPH)
      ? el.asParagraph().editAsText()
      : el.asListItem().editAsText();
    // setBackgroundColor throws on an empty range, and a copied cell can carry
    // an empty interior paragraph.
    if (textEl.getText().length > 0) textEl.setBackgroundColor(NOTE_HIGHLIGHT_9);
    return;
  }

  if (type === DocumentApp.ElementType.TABLE) {
    var table = el.asTable();
    for (var r = 0; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      for (var c = 0; c < row.getNumCells(); c++) {
        var cell = row.getCell(c);
        for (var i = 0; i < cell.getNumChildren(); i++) highlightElement9_(cell.getChild(i));
      }
    }
  }
}


// ── WRITING ──────────────────────────────────────────────────

/**
 * Writes one module's objectives into its Module Overview as a REAL Docs
 * numbered list, immediately after the anchor: one blank paragraph, then the
 * list. (Decided 2026-09-04: real lists, not typed "1)" text — they match the
 * Course Design Map exactly and the mechanism already exists.)
 *
 * ORDER OF OPERATIONS IS LOAD-BEARING. Adopting a list ID — and setGlyphType /
 * setNestingLevel with it — pulls in the list preset's 18pt/36pt indents, so
 * the indents have to be zeroed AFTER setListId. Zeroing first is silently
 * undone and the whole list comes out indented.
 *
 * @returns {number} objectives written
 */
function writeObjectives9_(devBody, anchorEl, cell) {
  var sources = objectiveSources9_(cell);
  if (sources.length === 0) return 0;

  var at    = devBody.getChildIndex(anchorEl);
  var items = [];

  // Inserted in REVERSE at a fixed index so the forward order comes out right —
  // the same trick replaceWithCopiedElements (Code2.gs:970) uses.
  for (var k = sources.length - 1; k >= 0; k--) {
    var item = devBody.insertListItem(at + 1, '');
    copyRuns9_(sources[k].text, item);

    if (sources[k].typed) {
      var textEl = item.editAsText();
      var full   = textEl.getText();
      var marker = full.match(CDM_TYPED_MARKER_RE_9);
      // Never delete the whole line: a cell holding only "1." is malformed, but
      // an empty list item is worse than a redundant numeral.
      if (marker && marker[0].length < full.length) textEl.deleteText(0, marker[0].length - 1);
    }
    items.unshift(item);
  }

  // "…objectives go in the SECOND, leaving one blank line" — the blank sits
  // between the refer-to line and the list.
  var blank = devBody.insertParagraph(at + 1, '');
  blank.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  zeroIndent_(blank);

  for (var g = 0; g < items.length; g++) {
    items[g].setGlyphType(DocumentApp.GlyphType.NUMBER);
    items[g].setNestingLevel(sources[g].level);
  }

  // One fresh list ID per module, so numbering restarts at 1 in each Overview.
  // A blank paragraph does NOT end a list — continuity is by list ID, not
  // adjacency — which is why this is required rather than cosmetic.
  freshListId_(devBody, items);

  for (var z = 0; z < items.length; z++) {
    if (items[z].getNestingLevel() === 0) {
      items[z].setIndentFirstLine(0);
      items[z].setIndentStart(0);
    }
    // Level 2+ keeps the preset indent on purpose: zeroing it would flatten the
    // hierarchy into an unreadable single column.
  }

  return items.length;
}


/**
 * Copies one Notes cell into the Development tab immediately after anchorEl.
 *
 * INSERTS AFTER the "Directions go here…" placeholder — never replaces it. The
 * deploy tools find their target by that exact string, so removing it would
 * quietly break "Deploy Activity Directions" for every slot this tool touched.
 *
 * @returns {number} elements written
 */
function insertNotes9_(devBody, anchorEl, cell, highlight) {
  var sources = noteSources9_(cell);
  if (sources.length === 0) return 0;

  var NORMAL   = DocumentApp.ParagraphHeading.NORMAL;
  var at       = devBody.getChildIndex(anchorEl);
  var inserted = [];

  for (var k = sources.length - 1; k >= 0; k--) {
    var el    = sources[k];
    var type  = el.getType();
    var added = null;

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      added = devBody.insertParagraph(at + 1, el.asParagraph().copy());
      // Forced to Normal: a Course Design Map cell styled as a heading would
      // otherwise inject an H2/H3 mid-module and break every other tool's idea
      // of where this module ends. Only when it actually is one — applying the
      // named style needlessly would restyle notes that were already fine.
      if (added.getHeading() !== NORMAL) added.setHeading(NORMAL);
      added.setIndentStart(0);
      added.setIndentFirstLine(0);
      added.setIndentEnd(RIGHT_INDENT);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      var src = el.asListItem();
      added   = devBody.insertListItem(at + 1, src.copy());
      if (added.getHeading() !== NORMAL) added.setHeading(NORMAL);
      // Copying a ListItem drops its bullet, because the source list ID does
      // not exist here — re-apply both, exactly as replaceWithCopiedElements does.
      added.setGlyphType(src.getGlyphType());
      added.setNestingLevel(src.getNestingLevel());
    } else if (type === DocumentApp.ElementType.TABLE) {
      added = devBody.insertTable(at + 1, el.asTable().copy());
    } else {
      continue;
    }

    inserted.unshift(added);
  }

  // Give each contiguous run of copied list items its own list ID so the
  // numbering restarts per activity instead of continuing the previous one's.
  restartCopiedListNumbering_(devBody, at + 1, inserted.length);

  if (highlight) {
    for (var h = 0; h < inserted.length; h++) highlightElement9_(inserted[h]);
  }

  return inserted.length;
}


/**
 * Opens a gap after content that is already in a destination, so "paste anyway"
 * appends rather than crowding the designer's own last line.
 *
 * @returns {GoogleAppsScript.Document.Paragraph} the new anchor
 */
function appendSpacer9_(devBody, lastEl) {
  var blank = devBody.insertParagraph(devBody.getChildIndex(lastEl) + 1, '');
  blank.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  zeroIndent_(blank);
  return blank;
}


// ── APPLY ────────────────────────────────────────────────────

/**
 * Sidebar-callable. Writes the chosen titles, objectives and activity notes
 * into the Development tab.
 *
 * Nothing here trusts the sidebar's view of the document. The parse, the scan
 * and the match are all redone from scratch; the parameters carry only the
 * user's DECISIONS (which modules, which activities, which title, how to
 * resolve an ambiguous row, whether to paste into an occupied slot). Matching
 * is deterministic, so the preview and the write agree — and a document edited
 * between the two is re-read rather than written from stale indices.
 *
 * Never overwrites. A destination that already has content is skipped and
 * reported unless the user explicitly ticked "paste anyway" for it, and the
 * "Directions go here…" placeholder is always preserved.
 *
 * @param {Object} params
 *   .titles      {Object}  module number → chosen title text
 *   .modules     {number[]} module numbers to process
 *   .activities  {string[]} Development-tab activity titles to include; [] means
 *                          none, and omitting the field means no filter at all
 *   .resolutions {Object}  "num||cdmLabel" → chosen Dev activity title, '' to skip
 *   .pasteAnyway {Object}  "ov||num" / "act||num||cdmLabel" → true
 *   .highlight   {boolean} cyan-highlight the pasted notes (default true)
 * @returns {string} plain-text summary for the sidebar
 */
function applyDesignMapToDevTab9(params) {
  params = params || {};
  var titles      = params.titles      || {};
  var resolutions = params.resolutions || {};
  var pasteAnyway = params.pasteAnyway || {};
  var highlight   = params.highlight !== false;

  var doc = DocumentApp.getActiveDocument();

  var devBody = getDevelopmentTabBody(doc);
  if (!devBody) throw new Error('Could not find a "Development" tab in this document.');

  var tabs = collectTabs(doc);
  var designTab = null;
  for (var i = 0; i < tabs.length; i++) {
    if (/\bdesign\b/i.test(tabs[i].title)) { designTab = tabs[i]; break; }
  }
  if (!designTab) throw new Error('Could not find a "Design" tab in this document.');

  // ── 1. Titles ──────────────────────────────────────────────
  // First, because replaceText does not change the child count — every
  // structural insert below would invalidate scanDevelopmentHeadings7_'s
  // childIndex values if it ran the other way round.
  var titlesWritten = 0;
  var titlesSkipped = [];
  var H2 = DocumentApp.ParagraphHeading.HEADING2;

  var headings = scanDevelopmentHeadings7_(devBody);
  if (headings.length === 0) {
    throw new Error('No numbered module headings found in the Development tab.');
  }

  for (var h = 0; h < headings.length; h++) {
    var info = headings[h];
    var chosenTitle = titles[info.num];
    if (!chosenTitle) continue;

    var para = devBody.getChild(info.childIndex).asParagraph();
    if (para.getHeading() !== H2) continue;   // document shifted since the scan

    if (info.titleIsPlaceholder) {
      // Replacing the placeholder substring rather than calling setText keeps
      // the replacement surgical: the parentheses, the hyphen and the heading's
      // Arial 17 bold all survive untouched. Word-bounded so it cannot match
      // inside a longer word.
      para.replaceText('\\b' + TITLE_PLACEHOLDER_7 + '\\b', chosenTitle);
      titlesWritten++;
    } else {
      titlesSkipped.push(info.displayLabel + ' — already reads "' + info.titlePart + '"');
    }
  }

  // ── 2. Objectives and notes ────────────────────────────────
  var wanted = {};
  var moduleNums = params.modules || [];
  for (var w = 0; w < moduleNums.length; w++) wanted[parseInt(moduleNums[w], 10)] = true;

  // An EMPTY array means "no activities chosen", not "all of them". The sidebar
  // always sends the list it rendered (every box ticked by default), so a user
  // who deliberately unticked all of them gets what they asked for rather than
  // the exact opposite. Omitting the field entirely is what disables the filter.
  var actFilter = null;
  if (Object.prototype.toString.call(params.activities) === '[object Array]') {
    actFilter = {};
    for (var f = 0; f < params.activities.length; f++) {
      actFilter[String(params.activities[f]).toLowerCase()] = true;
    }
  }

  var map        = parseDesignMap9_(designTab.body);
  var devModules = scanDevStructure9_(devBody);

  var objectivesWritten = 0;
  var notesWritten      = 0;
  var objectivesSkipped = [];
  var notesSkipped      = [];
  var unmatched         = [];

  for (var m = 0; m < devModules.length; m++) {
    var devMod = devModules[m];
    if (!wanted[devMod.num]) continue;

    var design = map.byNumber[devMod.num];
    if (!design) continue;

    // ── Objectives → Module Overview ──
    if (design.objectivesCell) {
      var ov = devMod.overview;
      if (!ov || !ov.anchor) {
        objectivesSkipped.push(devMod.displayLabel + ' — no "Module ' + devMod.num +
                               ' Overview" section was found');
      } else if (ov.content.length > 0 && !pasteAnyway['ov' + KEY_SEP_9 + devMod.num]) {
        objectivesSkipped.push(devMod.displayLabel + ' — the Module Overview already ' +
                               'has content, so nothing was overwritten');
      } else {
        // No appendSpacer9_ here: writeObjectives9_ already opens with a blank
        // paragraph of its own, and two in a row would just look like a gap.
        var ovAnchor = ov.content.length > 0
          ? ov.content[ov.content.length - 1]
          : ov.anchor;
        var wroteObjectives = writeObjectives9_(devBody, ovAnchor, design.objectivesCell);
        if (wroteObjectives > 0) objectivesWritten++;
      }
    }

    // ── Notes → under each activity's Directions placeholder ──
    var noteRows = [];
    for (var q = 0; q < design.activities.length; q++) {
      if (design.activities[q].hasNotes) noteRows.push(design.activities[q]);
    }
    if (noteRows.length === 0) continue;

    var matched    = matchActivities9_(noteRows, devMod.activities);
    var targets    = [];
    var skipRow    = {};
    var overridden = {};
    var reserved   = {};
    var k, s;

    // The user's dropdown picks are resolved FIRST and reserve their slot. A
    // pick is an explicit decision, so it must never lose a slot to the
    // matcher's guess for some row that happens to come earlier in the table.
    for (k = 0; k < noteRows.length; k++) {
      targets.push(matched.pairs[k]);

      var pick = resolutions[devMod.num + KEY_SEP_9 + noteRows[k].label];
      if (pick === undefined) continue;

      overridden[k] = true;
      targets[k]    = -1;
      if (String(pick) === '') { skipRow[k] = true; continue; }   // "Skip this row"

      for (s = 0; s < devMod.activities.length; s++) {
        if (reserved[s]) continue;
        if (devMod.activities[s].title.toLowerCase() === String(pick).toLowerCase()) {
          targets[k]  = s;
          reserved[s] = true;
          break;
        }
      }
    }

    // The computed matches then take whatever the overrides did not claim. One
    // slot, one row: a second claimant is reported rather than pasted twice.
    for (k = 0; k < targets.length; k++) {
      if (overridden[k] || targets[k] === -1) continue;
      if (reserved[targets[k]]) { targets[k] = -1; continue; }
      reserved[targets[k]] = true;
    }

    for (k = 0; k < noteRows.length; k++) {
      if (skipRow[k]) continue;

      var row     = noteRows[k];
      var rowName = devMod.displayLabel + ' / "' + row.label + '"';
      var slotIdx = targets[k];

      if (slotIdx === -1) {
        unmatched.push(rowName + ' — no matching activity in the Development tab');
        continue;
      }

      var slot = devMod.activities[slotIdx];

      if (actFilter && !actFilter[slot.title.toLowerCase()]) continue;  // not chosen

      if (!slot.placeholder && slot.content.length === 0) {
        notesSkipped.push(rowName + ' → "' + slot.title +
                          '" — no "Directions go here…" line to paste under');
        continue;
      }

      var anchor;
      if (slot.content.length > 0) {
        // Keyed by the Course Design Map row, not by the destination activity:
        // the row is what the sidebar lists, and a resolution override can move
        // it to a different slot after the user has ticked the box.
        if (!pasteAnyway['act' + KEY_SEP_9 + devMod.num + KEY_SEP_9 + row.label]) {
          notesSkipped.push(rowName + ' → "' + slot.title +
                            '" — that slot already has directions, so nothing was overwritten');
          continue;
        }
        anchor = appendSpacer9_(devBody, slot.content[slot.content.length - 1]);
      } else {
        anchor = slot.placeholder;
      }

      if (insertNotes9_(devBody, anchor, row.cell, highlight) > 0) notesWritten++;
    }
  }

  Logger.log('applyDesignMapToDevTab9: %s title(s), %s overview(s), %s note block(s) written; ' +
             '%s title(s) / %s overview(s) / %s note(s) skipped, %s row(s) unmatched.',
             titlesWritten, objectivesWritten, notesWritten,
             titlesSkipped.length, objectivesSkipped.length, notesSkipped.length,
             unmatched.length);

  // ── 3. Summary ─────────────────────────────────────────────
  var lines = ['✅ Development tab updated.', ''];
  lines.push('Module titles written: ' + titlesWritten);
  lines.push('Module Overviews filled: ' + objectivesWritten);
  lines.push('Activity note blocks pasted: ' + notesWritten);

  if (highlight && notesWritten > 0) {
    lines.push('', 'Pasted notes are highlighted cyan. Delete them before the ' +
                   'Blueprint goes to the SME.');
  }

  if (titlesSkipped.length > 0) {
    lines.push('', 'Titles left alone (' + titlesSkipped.length + ') — these headings ' +
                   'no longer have the "Title" placeholder:');
    for (var t1 = 0; t1 < titlesSkipped.length; t1++) lines.push('  • ' + titlesSkipped[t1]);
  }

  if (objectivesSkipped.length > 0) {
    lines.push('', 'Module Overviews left alone (' + objectivesSkipped.length + '):');
    for (var t2 = 0; t2 < objectivesSkipped.length; t2++) lines.push('  • ' + objectivesSkipped[t2]);
  }

  if (notesSkipped.length > 0) {
    lines.push('', 'Activity notes left alone (' + notesSkipped.length + '):');
    for (var t3 = 0; t3 < notesSkipped.length; t3++) lines.push('  • ' + notesSkipped[t3]);
  }

  if (unmatched.length > 0) {
    lines.push('', 'Course Design Map rows with nowhere to go (' + unmatched.length + '):');
    for (var t4 = 0; t4 < unmatched.length; t4++) lines.push('  • ' + unmatched[t4]);
  }

  return lines.join('\n');
}
