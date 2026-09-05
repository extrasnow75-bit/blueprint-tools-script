// ============================================================
// Blueprint Tools — Code7.gs
// Add Module Titles & Module Dates (Beta): copies module titles from the
// Design tab's Course Design Map into the Development tab's H2
// headings, and optionally fills in module start/end dates from
// the Boise State registrar's academic calendar.
// ------------------------------------------------------------
// Last updated on 2026-09-04 at 23:18 MDT
// ------------------------------------------------------------
//
// Runs AFTER "Add Activity Titles, Tools, Due Date Headers, & Times", which is what
// creates the H2 headings this tool writes into.
//
// Relies on shared helpers in Code.gs / Code2.gs (same GAS namespace):
//   collectTabs, getDevelopmentTabBody
//
// The dates half of this tool has been split out into the standalone
// "Add Module Dates" tool (Code8.gs / Sidebar8.html) — see project memory
// project_blueprint_designmap_dev_tab.md, Phase 2. This file is kept in place,
// unchanged in behavior, because its title half still has to be absorbed into
// the upcoming Design Map → Dev Tab tool (Phase 3) before it can be retired.
// The academic-calendar helpers and the START_PLACEHOLDER_7/END_PLACEHOLDER_7/
// MONTHS_7/MONTH_LOOKUP_7/CALENDAR_BASE_7/FIVE_YEAR_URL_7 constants now live in
// Code8.gs — do NOT redefine them here (flat GAS namespace: a duplicate
// definition would not error, it would just silently win or lose).
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

// Same shape, but for Course Design Map header cells, where the colon and the
// title after it are optional ("Module 4:" with no title is a real case).
var CDM_MODULE_RE_7 = /^(module|week)\s+0*(\d+)\s*:?\s*(.*)$/i;


// ── SIDEBAR OPENER ───────────────────────────────────────────

function showModuleTitlesSidebar7() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar7')
    .setTitle('Add Module Titles & Module Dates')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}


// ── HEADING STATE ────────────────────────────────────────────

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


// ── COURSE DESIGN MAP: TITLES ONLY ───────────────────────────

/**
 * Reads module titles out of the Design tab's Course Design Map.
 *
 * Deliberately narrow. The AI tools' parseCourseDesignMap pulled CLOs,
 * readings, activity descriptions and hyperlink targets to feed Gemini
 * prompts; none of that is wanted here, and its link-scraping helpers went
 * with the AI files. This reads the two title sources and nothing else.
 *
 * Two sources per module, either of which may be absent:
 *   • the merged blue header cell — "Module 2: Matrices as Datasets…"
 *   • the "Title" row's value cell — "Using Matrices to Represent Data"
 *
 * @returns {Object} map of module number → {headerTitle, rowTitle, displayLabel}
 */
function parseDesignMapTitles7_(designBody) {
  var byNumber = {};
  var tables   = designBody.getTables();

  for (var t = 0; t < tables.length; t++) {
    var table   = tables[t];
    var numRows = table.getNumRows();
    if (numRows < 2) continue;

    // Identify the Course Design Map by a "CLO"/"MLO" label anywhere in the
    // first several rows. Cell reads are server round-trips, so the scan is
    // bounded and its results are reused by the parse loop below.
    var rowCache  = {};
    var isCDM     = false;
    var scanLimit = Math.min(numRows, 15);

    for (var ri = 0; ri < scanLimit && !isCDM; ri++) {
      var scanRow  = table.getRow(ri);
      var scanN    = scanRow.getNumCells();
      var rowText  = [];
      for (var ci = 0; ci < scanN && ci < 2; ci++) {
        var cellStr = scanRow.getCell(ci).getText();
        rowText.push(cellStr);
        var lc = cellStr.toLowerCase();
        if (lc.indexOf('clo') !== -1 || lc.indexOf('mlo') !== -1) isCDM = true;
      }
      rowCache[ri] = rowText;
    }
    if (!isCDM) continue;

    var currentNum = null;

    for (var r = 0; r < numRows; r++) {
      var cells = rowCache[r];
      if (!cells) {
        var dataRow = table.getRow(r);
        var dataN   = dataRow.getNumCells();
        cells = [];
        for (var c = 0; c < dataN && c < 2; c++) cells.push(dataRow.getCell(c).getText());
      }
      if (cells.length === 0) continue;

      var firstCell = String(cells[0]).trim();
      var headerHit = firstCell.match(CDM_MODULE_RE_7);

      if (headerHit) {
        currentNum = parseInt(headerHit[2], 10);
        if (!byNumber[currentNum]) {
          byNumber[currentNum] = { headerTitle: '', rowTitle: '', displayLabel: '' };
        }
        // Everything after "Module N:" is the header title. Often empty.
        byNumber[currentNum].headerTitle  = String(headerHit[3] || '').trim();
        byNumber[currentNum].displayLabel = headerHit[1] + ' ' + headerHit[2];
        continue;
      }

      if (currentNum === null || cells.length < 2) continue;

      if (firstCell.toLowerCase() === 'title') {
        byNumber[currentNum].rowTitle = String(cells[1]).trim();
      }
    }

    if (Object.keys(byNumber).length > 0) break; // found the CDM — stop searching
  }

  return byNumber;
}


// ── SIDEBAR DATA ─────────────────────────────────────────────

/**
 * Sidebar-callable. Builds the per-module panel data: current heading state,
 * both candidate titles, and any conflicts or mismatches worth surfacing.
 *
 * The Development tab drives the module list, since that is where the writes
 * land. Course Design Map modules with no matching heading are reported but do
 * not block the run.
 */
function getModuleTitlesSidebarData7() {
  var result = {
    modules:          [],
    unmatchedDesign:  [],
    conflicts:        [],
    error:            ''
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

    var headings = scanDevelopmentHeadings7_(devBody);
    if (headings.length === 0) {
      result.error = 'No numbered module headings (Module 1, Week 1, …) were found ' +
                     'in the Development tab. Run "Add Activity Titles, Tools, Due Date ' +
                     'Headers, & Times" first.';
      return result;
    }

    var designTitles = designTab ? parseDesignMapTitles7_(designTab.body) : {};
    if (!designTab) {
      result.error = 'Could not find a "Design" tab in this document.';
      return result;
    }

    var seen = {};
    for (var h = 0; h < headings.length; h++) {
      var info = headings[h];
      // Match across tabs on the parsed integer, never the label string — a
      // "Module 1" Design Map row must pair with a "Week 01" Dev heading.
      var design = designTitles[info.num] || { headerTitle: '', rowTitle: '' };
      seen[info.num] = true;

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
        num:          info.num,
        displayLabel: info.displayLabel,
        state:        info.state,
        rawHeading:   info.rawHeading,
        currentTitle: info.titleIsPlaceholder ? '' : info.titlePart,
        currentDates: info.dateIsPlaceholder  ? '' : (info.datePart || ''),
        canWriteTitle: info.titleIsPlaceholder,
        canWriteDates: info.dateIsPlaceholder,
        options:      options,
        conflict:     isConflict
      });
    }

    for (var key in designTitles) {
      if (!seen[key]) {
        var d = designTitles[key];
        result.unmatchedDesign.push({
          module: d.displayLabel || ('Module ' + key),
          title:  d.headerTitle || d.rowTitle || '(untitled)'
        });
      }
    }

    Logger.log('getModuleTitlesSidebarData7: %s heading(s), %s conflict(s), %s unmatched.',
               result.modules.length, result.conflicts.length, result.unmatchedDesign.length);

  } catch (e) {
    Logger.log('getModuleTitlesSidebarData7 error: ' + e.message);
    result.error = e.message;
  }

  return result;
}


// ── APPLY ────────────────────────────────────────────────────

/**
 * Sidebar-callable. Writes the chosen titles and dates into the Development
 * tab's module headings.
 *
 * Writes by replacing the placeholder substrings rather than calling setText
 * on the paragraph. That keeps the replacement surgical: the parentheses, the
 * hyphen and the heading's Arial 17 bold all survive untouched, and it sidesteps
 * the question of whether that bold lives in the Heading 2 named style or as a
 * manual run override.
 *
 * Never overwrites. A field whose placeholder is already gone is skipped and
 * reported, so a re-run cannot destroy a designer's own edits.
 *
 * @param {Object} params
 *   .titles {Object}  module number → chosen title text
 *   .dates  {Object}  module number → {start: ms, end: ms}  (omit for titles-only)
 * @returns {string} plain-text summary for the sidebar
 */
function applyModuleTitlesAndDates7(params) {
  var titles = (params && params.titles) || {};
  var dates  = (params && params.dates)  || {};

  var doc     = DocumentApp.getActiveDocument();
  var devBody = getDevelopmentTabBody(doc);
  if (!devBody) throw new Error('Could not find a "Development" tab in this document.');

  var headings = scanDevelopmentHeadings7_(devBody);
  if (headings.length === 0) throw new Error('No numbered module headings found in the Development tab.');

  var titlesWritten  = 0;
  var datesWritten   = 0;
  var titlesSkipped  = [];
  var datesSkipped   = [];

  var H2 = DocumentApp.ParagraphHeading.HEADING2;

  for (var i = 0; i < headings.length; i++) {
    var info = headings[i];
    var para = devBody.getChild(info.childIndex).asParagraph();

    // Guard against the document having shifted since the scan.
    if (para.getHeading() !== H2) continue;

    var chosenTitle = titles[info.num];
    if (chosenTitle) {
      if (info.titleIsPlaceholder) {
        // Word-bounded so it cannot match inside a longer word.
        para.replaceText('\\b' + TITLE_PLACEHOLDER_7 + '\\b', chosenTitle);
        titlesWritten++;
      } else {
        titlesSkipped.push(info.displayLabel + ' — already reads "' + info.titlePart + '"');
      }
    }

    var moduleDates = dates[info.num];
    if (moduleDates && moduleDates.start && moduleDates.end) {
      if (info.dateIsPlaceholder) {
        para.replaceText('\\b' + START_PLACEHOLDER_7 + '\\b', formatModuleDate7(moduleDates.start));
        para.replaceText('\\b' + END_PLACEHOLDER_7   + '\\b', formatModuleDate7(moduleDates.end));
        datesWritten++;
      } else {
        datesSkipped.push(info.displayLabel + ' — already reads "(' + info.datePart + ')"');
      }
    }
  }

  Logger.log('applyModuleTitlesAndDates7: %s title(s), %s date range(s) written; %s / %s skipped.',
             titlesWritten, datesWritten, titlesSkipped.length, datesSkipped.length);

  var lines = ['✅ Development tab updated.', ''];
  lines.push('Titles written: ' + titlesWritten);
  lines.push('Date ranges written: ' + datesWritten);

  if (titlesSkipped.length > 0) {
    lines.push('', 'Titles left alone (' + titlesSkipped.length + ') — these headings ' +
                   'no longer have the "Title" placeholder, so nothing was overwritten:');
    for (var s = 0; s < titlesSkipped.length; s++) lines.push('  • ' + titlesSkipped[s]);
  }

  if (datesSkipped.length > 0) {
    lines.push('', 'Dates left alone (' + datesSkipped.length + ') — these headings already ' +
                   'have dates, so nothing was overwritten:');
    for (var d2 = 0; d2 < datesSkipped.length; d2++) lines.push('  • ' + datesSkipped[d2]);
  }

  if (Object.keys(dates).length === 0) {
    lines.push('', 'Dates were not added, so "(start date - end date)" is still in each ' +
                   'heading. Run this tool again later to fill them in, or remove the ' +
                   'text by hand before sharing the Blueprint.');
  }

  return lines.join('\n');
}
