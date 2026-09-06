// ============================================================
// Blueprint Tools — Code9.gs
// Design Map → Dev Tab: copies the Design tab's Course Design Map into the
// Development tab — module titles into the H2 headings, the CLO/MLO row into
// each Module Overview as a real numbered list, and every other row's notes
// under that activity's "Directions go here…" placeholder.
// ------------------------------------------------------------
// Last updated on 2026-09-05 at 23:20 MDT
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
//   collectTabs, zeroIndent_, _fmt-era constants
//     FONT / BLACK / RED, RIGHT_INDENT                        (Code.gs)
//   getDevelopmentTabBody, stripActivityHeading,
//   DIRECTIONS_PLACEHOLDER_TEXT, restartCopiedListNumbering_  (Code2.gs)
//   START_PLACEHOLDER_7, END_PLACEHOLDER_7                    (Code8.gs)
//
// freshListId_ (Code2.gs) is deliberately NOT used any more. Objectives were
// real Docs numbered lists until 2026-09-05; they are unnumbered paragraphs
// now, so this file mints no list IDs. The helper stays in Code2.gs for the
// deploy tools, which still need it.
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

// Typed list markers at the head of a line in the CLO/MLO cell ("1. ", "2) ",
// "• ", "– "). Objectives are written UNNUMBERED (user, 2026-09-05), so a
// hand-typed numeral would be the only number on the page and would read as a
// leftover. Real Docs list items carry their number as a glyph, not as text,
// so they arrive here already clean.
var CDM_TYPED_MARKER_RE_9 = /^\s*(?:\d+\s*[.)\]]|[•●▪–—-])\s+/;

// The objectives block. Rather than hardcoding one header, the cell's OWN
// header lines are detected and reformatted — a Course Design Map cell holds
// MLOs, or CLOs, or both, and mirroring what it says keeps all three correct.
// DEFAULT_OBJECTIVES_HEADER_9 is the fallback for a cell that is a bare list of
// objectives with no header at all.
var OBJECTIVES_HEADER_RE_9      = /^\s*(?:course|module)\s+learning\s+objectives\b/i;
var DEFAULT_OBJECTIVES_HEADER_9 = 'Module Learning Objectives (MLOs)';

// The Blueprint's convention for "make this an H2 in Canvas": the marker is
// bold red text on an otherwise Normal paragraph, NOT a real Docs heading —
// same as "Required Readings (H2)" in the deployed directions.
var H2_MARKER_9 = '(H2)';

// A heading marker the cell already carries, so "… (MLOs) (H2)" does not come
// out as "… (MLOs) (H2) (H2)".
var TRAILING_MARKER_RE_9 = /\s*\((?:H[1-6])\)\s*:?\s*$/i;

// Cyan blue, the IDC convention for "this is a note to myself, delete before
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
        // Last of the slot's preamble lines (H4 → "Estimated time:" → tool
        // line). Notes anchor here when the placeholder is gone, which is how
        // they land ABOVE directions that were already deployed.
        lastPreamble: para,
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
      // the tool line and matches neither a blank nor a heading. Tracked rather
      // than merely skipped, because the last of them is the notes anchor in a
      // slot whose placeholder has already been consumed by Deploy.
      if (/^estimated time/i.test(trimmed))       { slot.lastPreamble = para; continue; }
      if (/link to settings tab$/i.test(trimmed)) { slot.lastPreamble = para; continue; }
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

// Words that carry no signal about WHICH activity a row is. Deliberately short:
// every word dropped here is a word the keyword pass can no longer match on,
// and Blueprint activity names are only two or three words long to begin with.
var STOPWORDS_9 = {
  a: 1, an: 1, and: 1, the: 1, or: 1, of: 1, to: 1, in: 1, on: 1, at: 1,
  for: 1, with: 1, from: 1, by: 1, this: 1, that: 1, is: 1, are: 1, as: 1,
  into: 1, related: 1
};

/**
 * Crude singulariser, enough for activity names: "Videos" and "Video",
 * "Readings" and "Reading", "Replies" and "Reply" have to land on one token or
 * the keyword pass misses the obvious matches it exists to catch.
 */
function stem9_(w) {
  if (w.length > 4 && w.slice(-3) === 'ies') return w.slice(0, -3) + 'y';
  if (w.length > 3 && w.slice(-1) === 's' && w.slice(-2) !== 'ss') return w.slice(0, -1);
  return w;
}

/**
 * Content words of a normalized title. Bare numbers are dropped along with the
 * stopwords: the "1" in "Readings 1" is an index, not a keyword, and keeping it
 * would let "Readings 1" match "Discussion 1".
 */
function words9_(norm) {
  var parts = String(norm || '').split(' ');
  var out   = [];
  for (var i = 0; i < parts.length; i++) {
    var w = parts[i];
    if (!w || STOPWORDS_9[w] || /^\d+$/.test(w)) continue;
    out.push(stem9_(w));
  }
  return out;
}

/**
 * How strongly two titles share content words, 0 to 1.
 *
 * Divided by the SHORTER word list, not the union: a Course Design Map row is
 * often a fragment of the activity it belongs to ("Videos 1" → "Watch Videos"),
 * and a Jaccard score would punish that for the words the longer title adds.
 */
function keywordScore9_(a, b) {
  if (!a.words.length || !b.words.length) return 0;

  var seen = {};
  for (var i = 0; i < a.words.length; i++) seen[a.words[i]] = true;

  var shared = 0;
  for (var j = 0; j < b.words.length; j++) {
    if (seen[b.words[j]]) { shared++; seen[b.words[j]] = false; }  // count each word once
  }
  if (shared === 0) return 0;

  return shared / Math.min(a.words.length, b.words.length);
}

// One whole content word in common with the shorter title. Below this the guess
// is noise — and an unmatched row is a fine outcome, since it still reaches the
// user as a dropdown either way.
var KEYWORD_MIN_SCORE_9 = 0.5;

/** Pre-computes every comparison form of one title, so the tiers stay cheap. */
function matchKey9_(title) {
  var raw  = String(title || '').trim();
  var norm = normTitle9_(raw);
  return {
    raw:   raw,
    lower: raw.toLowerCase(),
    norm:  norm,
    base:  normBase9_(raw),
    words: words9_(norm)
  };
}

// Ordered strongest → weakest. Each tier is tried across ALL still-unmatched
// design-map rows before the next one is considered, so a weak match can never
// steal a slot that a stronger one wants.
// `weak` marks a tier whose result is a GUESS rather than a match. Weak results
// are still applied, but the sidebar surfaces them for confirmation with the
// guess pre-selected, so the user changes them by exception instead of picking
// every row by hand.
var MATCH_TIERS_9 = [
  { name: 'exact title',        fn: function (a, b) { return a.raw  === b.raw;  } },
  { name: 'ignoring case',      fn: function (a, b) { return a.lower === b.lower; } },
  { name: 'ignoring wording',   fn: function (a, b) { return !!a.norm && a.norm === b.norm; } },
  { name: 'ignoring numbering', fn: function (a, b) { return !!a.base && a.base === b.base; } },
  { name: 'by leading words',   weak: true, note: 'matched on the opening words only',
    fn: function (a, b) {
      // Word-boundaried on purpose: "read" must not claim "readings", and a
      // three-letter stem is too weak to be evidence of anything.
      if (!a.base || !b.base || a.base.length < 4 || b.base.length < 4) return false;
      return a.base.indexOf(b.base + ' ') === 0 || b.base.indexOf(a.base + ' ') === 0;
    } }
];

var KEYWORD_NOTE_9 = 'matched on a shared keyword';

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
 * @returns {{pairs: number[], tiers: string[], weak: boolean[], notes: string[]}}
 *          all index-aligned with cdmRows
 */
function matchActivities9_(cdmRows, devActivities) {
  var cdm = [], dev = [], pairs = [], tiers = [], weak = [], notes = [];
  var d, c, t;

  for (c = 0; c < cdmRows.length; c++) {
    cdm.push(matchKey9_(cdmRows[c].label));
    pairs.push(-1);
    tiers.push('');
    weak.push(false);
    notes.push('');
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
        weak[c]  = !!tier.weak;
        notes[c] = tier.note || '';
        taken[d] = true;
        break;
      }
    }
  }

  // ── Keyword pass ───────────────────────────────────────────
  // Last resort for the rows the string tiers could not place: score every
  // remaining pair on shared content words and take the best ones. "Videos 1"
  // finds "Watch Videos" here — they share no prefix, so nothing above could
  // see it, but the one word they do share is the whole of the shorter title.
  //
  // Scored globally and assigned best-first, NOT row by row. Walking the table
  // in order would let an early mediocre pairing consume a slot that a later
  // row matches outright.
  var candidates = [];
  for (c = 0; c < cdm.length; c++) {
    if (pairs[c] !== -1) continue;
    for (d = 0; d < dev.length; d++) {
      if (taken[d]) continue;
      var score = keywordScore9_(cdm[c], dev[d]);
      if (score >= KEYWORD_MIN_SCORE_9) candidates.push({ c: c, d: d, score: score });
    }
  }

  // Ties break by table order, so the same document always produces the same
  // answer — which is what lets the apply pass re-derive the sidebar's preview.
  candidates.sort(function (x, y) {
    return (y.score - x.score) || (x.c - y.c) || (x.d - y.d);
  });

  for (var k = 0; k < candidates.length; k++) {
    var cand = candidates[k];
    if (pairs[cand.c] !== -1 || taken[cand.d]) continue;
    pairs[cand.c]  = cand.d;
    tiers[cand.c]  = 'by shared keywords';
    weak[cand.c]   = true;
    notes[cand.c]  = KEYWORD_NOTE_9;
    taken[cand.d]  = true;
  }

  // ── Shared-slot pass ───────────────────────────────────────
  // Everything above is one row per slot. That is right as a default — it makes
  // genuinely distinct rows spread across distinct activities — but it strands
  // the common case where a Course Design Map lists "Readings 1" AND "Readings
  // 2" against a single "Readings" activity: the first claims it and the second
  // reports no match at all, in every module.
  //
  // So rows still unplaced get one more look, this time at slots that are
  // already spoken for. Always a guess, never silent: the sidebar surfaces it
  // pre-selected, and two rows landing in one slot stack in table order.
  for (c = 0; c < cdm.length; c++) {
    if (pairs[c] !== -1) continue;

    var bestD = -1;
    var bestScore = 0;
    for (d = 0; d < dev.length; d++) {
      if (!taken[d]) continue;                 // free slots were already tried
      var s = keywordScore9_(cdm[c], dev[d]);
      if (s > bestScore) { bestScore = s; bestD = d; }
    }

    if (bestD !== -1 && bestScore >= KEYWORD_MIN_SCORE_9) {
      pairs[c] = bestD;
      tiers[c] = 'sharing an activity';
      weak[c]  = true;
      notes[c] = 'no activity of its own — this one already has notes from another row';
    }
  }

  return { pairs: pairs, tiers: tiers, weak: weak, notes: notes };
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
      // devTitles feeds the match-resolution dropdowns only. It is NOT the
      // picker: the Development tab's activities all come from the one Course
      // Pattern Table, so they are identical in every module, whereas the
      // Course Design Map's rows vary module by module — and the map rows are
      // what actually gets copied, so they are what the user chooses from.
      var devTitles = [];
      for (var a = 0; a < devMod.activities.length; a++) {
        devTitles.push(devMod.activities[a].title);
      }

      // Rows with an empty Notes cell are reported but not offered: there is
      // nothing to copy, and silently omitting them would read as the tool
      // having missed a row that is plainly in the map.
      var noteRows  = [];
      var emptyRows = [];
      for (var q = 0; q < design.activities.length; q++) {
        if (design.activities[q].hasNotes) noteRows.push(design.activities[q]);
        else                               emptyRows.push(design.activities[q].label);
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
          matchedHow:    matched.tiers[k],
          matchedNote:   matched.notes[k],
          // Confident matches are not surfaced: a dozen certain pairings
          // awaiting confirmation would bury the few that need a decision.
          // A guess IS surfaced, pre-selected, so the user corrects by
          // exception rather than choosing every row by hand.
          needsDecision: !target || matched.weak[k],
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
        rows:      rows,
        emptyRows: emptyRows
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
 * Reads the CLO/MLO cell into one entry per line, flagging which lines are
 * section headers ("Module Learning Objectives (MLOs)", "Course Learning
 * Objectives (CLOs)") rather than objectives.
 *
 * Both cell shapes occur in the wild — a real Docs list, and plain paragraphs —
 * and both are read the same way now that objectives are written unnumbered.
 * A real list item carries its number as a glyph, so its text is already clean;
 * a hand-typed "1. " is stripped on write.
 */
function objectiveSources9_(cell) {
  var out = [];
  var n   = cell.getNumChildren();

  for (var i = 0; i < n; i++) {
    var el   = cell.getChild(i);
    var type = el.getType();
    var textEl;

    if      (type === DocumentApp.ElementType.LIST_ITEM) textEl = el.asListItem().editAsText();
    else if (type === DocumentApp.ElementType.PARAGRAPH) textEl = el.asParagraph().editAsText();
    else continue;

    var raw = textEl.getText();
    if (!raw.trim()) continue;

    out.push({
      text:     textEl,
      isHeader: OBJECTIVES_HEADER_RE_9.test(raw),
      // The header's own wording is kept — a cell that says "Course Learning
      // Objectives (CLOs)" must not be relabelled as MLOs — minus any heading
      // marker it already carries.
      label:    raw.trim().replace(TRAILING_MARKER_RE_9, '').trim()
    });
  }
  return out;
}


/**
 * Writes one objectives section header into an existing empty paragraph:
 * the label in bold black, then the "(H2)" marker in bold red.
 *
 * The paragraph stays Normal style. "(H2)" is the Blueprint's instruction to
 * whoever builds the Canvas page, not a Docs heading — making it a real
 * Heading 2 would render it at 17pt and, worse, would look like a module
 * boundary to every other tool in the suite that walks headings.
 */
function writeObjectivesHeader9_(para, label) {
  var text = para.editAsText();
  text.appendText(label + ' ' + H2_MARKER_9);

  var redStart = label.length + 1;
  var end      = redStart + H2_MARKER_9.length - 1;

  text.setFontFamily(0, end, FONT);
  text.setFontSize(0, end, 11);
  text.setBold(0, end, true);
  text.setForegroundColor(0, redStart - 1, BLACK);
  text.setForegroundColor(redStart, end, RED);
  return para;
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
 * Writes one module's objectives into its Module Overview, immediately after
 * the anchor: one blank paragraph, then a bold section header, then one plain
 * paragraph per objective.
 *
 * PLAIN PARAGRAPHS, NOT A DOCS LIST (user, 2026-09-05 — reversing the
 * 2026-09-04 "keep real numbered lists" call, and the "MLO 3.1:" prefix scheme
 * that briefly replaced it). Objectives carry no numbering at all now. That
 * retires freshListId_, setGlyphType/setNestingLevel and the whole
 * zero-the-indents-AFTER-setListId trap from this path.
 *
 * Section headers come from the cell itself, so a cell holding CLOs, or MLOs,
 * or both gets the right header on each block. Only a cell with no header of
 * its own gets DEFAULT_OBJECTIVES_HEADER_9 written for it.
 *
 * @returns {number} paragraphs written, header included
 */
function writeObjectives9_(devBody, anchorEl, cell) {
  var sources = objectiveSources9_(cell);
  if (sources.length === 0) return 0;

  var NORMAL  = DocumentApp.ParagraphHeading.NORMAL;
  var at      = devBody.getChildIndex(anchorEl);
  var written = 0;
  var hasOwnHeader = false;

  for (var s = 0; s < sources.length; s++) {
    if (sources[s].isHeader) { hasOwnHeader = true; break; }
  }

  // Inserted in REVERSE at a fixed index so the forward order comes out right —
  // the same trick replaceWithCopiedElements (Code2.gs:970) uses.
  for (var k = sources.length - 1; k >= 0; k--) {
    var para = devBody.insertParagraph(at + 1, '');
    para.setHeading(NORMAL);
    zeroIndent_(para);

    if (sources[k].isHeader) {
      writeObjectivesHeader9_(para, sources[k].label);
    } else {
      copyRuns9_(sources[k].text, para);

      var textEl = para.editAsText();
      var full   = textEl.getText();
      var marker = full.match(CDM_TYPED_MARKER_RE_9);
      // Never delete the whole line: a cell holding only "1." is malformed, but
      // an empty paragraph is worse than a stray numeral.
      if (marker && marker[0].length < full.length) textEl.deleteText(0, marker[0].length - 1);
    }
    written++;
  }

  if (!hasOwnHeader) {
    var head = devBody.insertParagraph(at + 1, '');
    head.setHeading(NORMAL);
    zeroIndent_(head);
    writeObjectivesHeader9_(head, DEFAULT_OBJECTIVES_HEADER_9);
    written++;
  }

  // "…objectives go in the SECOND, leaving one blank line" — the blank sits
  // between the refer-to line and the header.
  var blank = devBody.insertParagraph(at + 1, '');
  blank.setHeading(NORMAL);
  zeroIndent_(blank);

  return written;
}


/**
 * Copies one Notes cell into the Development tab immediately after anchorEl.
 *
 * INSERTS AFTER the "Directions go here…" placeholder — never replaces it. The
 * deploy tools find their target by that exact string, so removing it would
 * quietly break "Deploy Activity Directions" for every slot this tool touched.
 *
 * @returns {Array} the inserted elements, in document order. The caller needs
 *   the last one as the anchor when a second Design Map row lands in the same
 *   slot.
 */
function insertNotes9_(devBody, anchorEl, cell, highlight) {
  var sources = noteSources9_(cell);
  if (sources.length === 0) return [];

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

  return inserted;
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

// The run is split into a titles call plus one call per chunk of modules, so
// the sidebar can show a progress fraction that actually advances — a single
// blocking google.script.run call can report nothing until it returns.
//
// CHUNKED, not one call per module. Every call re-reads the document (parse the
// Course Design Map, scan the Development tab), and that fixed cost is the bulk
// of the work; paying it fifteen times instead of three would make the run
// slower than the version with no progress bar at all.
//
// Nothing in either function trusts the sidebar's view of the document. The
// parse, the scan and the match are redone from scratch; the parameters carry
// only the user's DECISIONS (which modules, which activities, which title, how
// to resolve an ambiguous row, whether to paste into an occupied slot).
// Matching is deterministic, so the preview and the write agree — and a
// document edited between the two is re-read rather than written from stale
// indices.

/**
 * Sidebar-callable, step 1 of the run. Writes the chosen module titles.
 *
 * Separate from the module pass and always run FIRST, because replaceText does
 * not change the child count — any structural insert would invalidate
 * scanDevelopmentHeadings7_'s childIndex values if it ran the other way round.
 *
 * Never overwrites: a heading whose "Title" placeholder is already gone is
 * skipped and reported.
 *
 * @param {Object} params  .titles {Object} module number → chosen title text
 * @returns {{written: number, skipped: string[]}}
 */
function applyDesignMapTitles9(params) {
  params = params || {};
  var titles = params.titles || {};

  var devBody = getDevelopmentTabBody(DocumentApp.getActiveDocument());
  if (!devBody) throw new Error('Could not find a "Development" tab in this document.');

  var written = 0;
  var skipped = [];
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
      written++;
    } else {
      skipped.push(info.displayLabel + ' — already reads "' + info.titlePart + '"');
    }
  }

  Logger.log('applyDesignMapTitles9: %s written, %s skipped.', written, skipped.length);
  return { written: written, skipped: skipped };
}


/**
 * Sidebar-callable, step 2 of the run, called once per chunk of modules.
 * Writes each chosen module's objectives into its Module Overview and its
 * Course Design Map notes into the matching activity slots.
 *
 * Never overwrites. A destination that already has content is skipped and
 * reported unless the user explicitly ticked "paste anyway" for it, and the
 * "Directions go here…" placeholder is always preserved.
 *
 * @param {Object} params
 *   .modules     {number[]} module numbers in THIS chunk
 *   .rows        {Object}  "num||cdmLabel" → true for each Design Map note row
 *                          to copy; anything absent is left alone
 *   .objectives  {Object}  module number → true to write that Module Overview
 *   .resolutions {Object}  "num||cdmLabel" → chosen Dev activity title, '' to skip
 *   .pasteAnyway {Object}  "ov||num" / "act||num||cdmLabel" → true
 *   .highlight   {boolean} cyan-blue-highlight the pasted notes (default true)
 * @returns {{objectives: number, notes: number, objectivesSkipped: string[],
 *            notesSkipped: string[], unmatched: string[]}}
 */
function applyDesignMapModules9(params) {
  params = params || {};
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

  var wanted = {};
  var moduleNums = params.modules || [];
  for (var w = 0; w < moduleNums.length; w++) wanted[parseInt(moduleNums[w], 10)] = true;

  // Selection is per Course Design Map ROW, keyed "num||label", not per
  // Development tab activity title. The dev tab's activities are one repeated
  // Course Pattern Table, identical in every module; the map's rows differ
  // module by module, and they are what gets copied. An absent key means
  // unchecked — the sidebar always sends everything it offered.
  var wantRow        = params.rows       || {};
  var wantObjectives = params.objectives || {};

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
    if (design.objectivesCell && wantObjectives[devMod.num]) {
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

    var matched = matchActivities9_(noteRows, devMod.activities);
    var targets = [];
    var skipRow = {};
    var k, s;

    // matchActivities9_ has already settled who gets which slot, including the
    // rare case of two rows sharing one. All that is left is to honour the
    // user's dropdown picks, which override the computed answer outright — an
    // explicit choice is never second-guessed, and never blocked because some
    // other row got there first.
    for (k = 0; k < noteRows.length; k++) {
      targets.push(matched.pairs[k]);

      var pick = resolutions[devMod.num + KEY_SEP_9 + noteRows[k].label];
      if (pick === undefined) continue;

      targets[k] = -1;
      if (String(pick) === '') { skipRow[k] = true; continue; }   // "Skip this row"

      for (s = 0; s < devMod.activities.length; s++) {
        if (devMod.activities[s].title.toLowerCase() === String(pick).toLowerCase()) {
          targets[k] = s;
          break;
        }
      }
    }

    // Where a slot receives more than one row, each block is anchored after the
    // one before it so they stack in table order instead of the later row
    // landing on top of the earlier one.
    var slotAnchor = {};

    for (k = 0; k < noteRows.length; k++) {
      if (skipRow[k]) continue;

      var row     = noteRows[k];
      var rowKey  = devMod.num + KEY_SEP_9 + row.label;
      if (!wantRow[rowKey]) continue;                    // not ticked in the picker

      var rowName = devMod.displayLabel + ' / "' + row.label + '"';
      var slotIdx = targets[k];

      if (slotIdx === -1) {
        unmatched.push(rowName + ' — no matching activity in the Development tab');
        continue;
      }

      var slot = devMod.activities[slotIdx];

      if (slot.content.length > 0 &&
          // Keyed by the Course Design Map row, not by the destination activity:
          // the row is what the sidebar lists, and a resolution override can move
          // it to a different slot after the user has ticked the box.
          !pasteAnyway['act' + KEY_SEP_9 + devMod.num + KEY_SEP_9 + row.label]) {
        notesSkipped.push(rowName + ' → "' + slot.title +
                          '" — that slot already has directions, so nothing was overwritten');
        continue;
      }

      // TOP of the slot, above any directions already there (user, 2026-09-05).
      //
      // The anchor is the placeholder when one survives, NOT the tool line, and
      // that is a hard constraint rather than a preference: findDirectionsPlaceholder
      // (Code2.gs:933) only scans 11 children past the H4, so a note block
      // inserted ABOVE the placeholder can push it out of that window — after
      // which "Deploy Activity Directions" silently stops finding the slot.
      // Once Deploy has consumed the placeholder there is nothing left to
      // protect, so the tool line becomes the anchor and the notes land above
      // the directions.
      var anchor = slotAnchor[slotIdx] || slot.placeholder || slot.lastPreamble;
      if (!anchor) {
        notesSkipped.push(rowName + ' → "' + slot.title +
                          '" — could not find a place to paste inside that slot');
        continue;
      }

      // One blank line between the anchor and the notes — "two hard returns",
      // which is one empty paragraph.
      var spacer   = appendSpacer9_(devBody, anchor);
      var inserted = insertNotes9_(devBody, spacer, row.cell, highlight);

      if (inserted.length > 0) {
        notesWritten++;
        slotAnchor[slotIdx] = inserted[inserted.length - 1];
      }
    }
  }

  Logger.log('applyDesignMapModules9: %s module(s) in chunk; %s overview(s), %s note block(s) ' +
             'written; %s overview(s) / %s note(s) skipped, %s row(s) unmatched.',
             moduleNums.length, objectivesWritten, notesWritten,
             objectivesSkipped.length, notesSkipped.length, unmatched.length);

  // Structured counts rather than prose: the sidebar calls this once per chunk
  // and accumulates across the whole run before rendering one summary.
  return {
    objectives:        objectivesWritten,
    notes:             notesWritten,
    objectivesSkipped: objectivesSkipped,
    notesSkipped:      notesSkipped,
    unmatched:         unmatched
  };
}
