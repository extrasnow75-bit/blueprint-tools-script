// ============================================================
// Blueprint Tools — Code7.gs
// Add Module Titles & Dates: copies module titles from the
// Design tab's Course Design Map into the Development tab's H2
// headings, and optionally fills in module start/end dates from
// the Boise State registrar's academic calendar.
// ------------------------------------------------------------
// Last updated on 2026-08-30 at 20:53 MDT
// ------------------------------------------------------------
//
// Runs AFTER "Add Activity Titles, Tools, & Times", which is what
// creates the H2 headings this tool writes into.
//
// Relies on shared helpers in Code.gs / Code2.gs (same GAS namespace):
//   collectTabs, getDevelopmentTabBody
// ============================================================


// ── CONSTANTS ────────────────────────────────────────────────

// The three placeholder substrings this tool replaces. They are the
// load-bearing contract with the Blueprint template: the tool writes a
// field ONLY where its placeholder is still present, which is what makes
// "skip and report" work without ever clobbering a designer's own text.
var TITLE_PLACEHOLDER_7 = 'Title';
var START_PLACEHOLDER_7 = 'start date';
var END_PLACEHOLDER_7   = 'end date';

// Month names as the user specified them: AP style, NOT uniform three-letter.
// March through July are spelled out and September truncates to four letters.
var MONTHS_7 = ['Jan', 'Feb', 'March', 'April', 'May', 'June',
                'July', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];

var MONTH_LOOKUP_7 = {
  january: 0, jan: 0, february: 1, feb: 1, march: 2, mar: 2,
  april: 3, apr: 3, may: 4, june: 5, jun: 5, july: 6, jul: 6,
  august: 7, aug: 7, september: 8, sept: 8, sep: 8, october: 9, oct: 9,
  november: 10, nov: 10, december: 11, dec: 11
};

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

var CALENDAR_BASE_7 = 'https://www.boisestate.edu/registrar/';
var FIVE_YEAR_URL_7 = CALENDAR_BASE_7 + 'boise-state-academic-calendars/5-year-academic-calendar/';


// ── SIDEBAR OPENER ───────────────────────────────────────────

function showModuleTitlesSidebar7() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar7')
    .setTitle('Add Module Titles & Dates')
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
                     'in the Development tab. Run "Add Activity Titles, Tools, & Times" first.';
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


// ── ACADEMIC CALENDAR ────────────────────────────────────────

/** Strips tags and decodes the handful of entities the registrar pages use. */
function stripHtml7_(html) {
  return String(html)
    // Superscripts serve two different jobs on these pages and must be treated
    // differently. Header cells carry footnote markers ("Start Date<sup>2</sup>")
    // which have to go, or they pollute the column matching. Session names carry
    // ordinals ("1<sup>st</sup> 5-week") which are part of the name — stripping
    // those turned six of the eight Fall 2026 sessions into "1 5-week",
    // "2 5-week" and so on. Numeric content is a marker; alphabetic is an ordinal.
    .replace(/<sup[^>]*>\s*[\d\s.,*†‡]+\s*<\/sup>/gi, '')
    .replace(/<sup[^>]*>([\s\S]*?)<\/sup>/gi, '$1')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&#8211;|&ndash;/gi, '-')
    .replace(/&#8212;|&mdash;/gi, '-')
    .replace(/&quot;/gi, '"')
    .replace(/&#0?39;|&rsquo;|&lsquo;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}


/**
 * Parses "August 24" / "November 23, 2026" into a Date in the given year.
 * Returns null when no month/day can be found.
 */
function parseCalendarDate7_(text, year) {
  var m = String(text).match(/([A-Za-z]+)\.?\s+(\d{1,2})/);
  if (!m) return null;
  var month = MONTH_LOOKUP_7[m[1].toLowerCase()];
  if (month === undefined) return null;

  // An explicit year in the string wins — Spring terms print December dates
  // from the prior calendar year in a few places.
  var explicit = String(text).match(/\b(20\d{2})\b/);
  var useYear  = explicit ? parseInt(explicit[1], 10) : year;
  return new Date(useYear, month, parseInt(m[2], 10));
}


/**
 * Pulls the session deadlines table out of a registrar calendar page.
 * @returns {Array<{session, start: Date, end: Date}>}
 */
function parseSessionTable7_(html, year) {
  var sessions = [];
  var tables   = html.match(/<table[\s\S]*?<\/table>/gi) || [];

  for (var t = 0; t < tables.length; t++) {
    var tableHtml = tables[t];
    if (!/last date of course instruction/i.test(tableHtml)) continue;

    var rows = tableHtml.match(/<tr[\s\S]*?<\/tr>/gi) || [];
    if (rows.length < 2) continue;

    var headerCells = (rows[0].match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || []).map(stripHtml7_);
    var sessionCol = -1, startCol = -1, endCol = -1;

    for (var c = 0; c < headerCells.length; c++) {
      var h = headerCells[c].toLowerCase();
      if      (sessionCol < 0 && h.indexOf('session') !== -1)     sessionCol = c;
      else if (startCol   < 0 && h.indexOf('start date') !== -1)  startCol   = c;
      else if (endCol     < 0 && h.indexOf('last date of course instruction') !== -1) endCol = c;
    }
    if (sessionCol < 0 || startCol < 0 || endCol < 0) continue;

    for (var r = 1; r < rows.length; r++) {
      var cells = (rows[r].match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || []).map(stripHtml7_);
      if (cells.length <= endCol) continue;

      var name  = cells[sessionCol];
      var start = parseCalendarDate7_(cells[startCol], year);
      var end   = parseCalendarDate7_(cells[endCol],   year);
      if (!name || !start || !end) continue;

      sessions.push({ session: name, start: start.getTime(), end: end.getTime() });
    }

    if (sessions.length > 0) break;
  }

  return sessions;
}


/**
 * Recovers a holiday's name from the text immediately preceding its
 * "(no classes)" marker.
 *
 * The tricky part is that the previous entry's dates butt right up against
 * this entry's name once tags are stripped ("…December 22 Labor Day"), so the
 * trailing words have to be taken and then any leading date expression peeled
 * back off. Names containing digits are real — "America250" — so a blanket
 * "drop tokens with numbers" rule would break them.
 */
function extractHolidayName7_(before) {
  // Work backwards from the end rather than peeling from the front. The name
  // is whatever follows the LAST date expression, and anchoring a peel at the
  // front of a fixed-size candidate window does not line up with wherever that
  // date actually falls ("…Grades due December 22 Labor Day" kept the date).
  //
  // "America250" survives because the date pattern requires whitespace between
  // the month and the day, which a name with trailing digits does not have.
  var dateExpr = /[A-Za-z]+\.?\s+\d{1,2}(?:,\s*\d{4})?/g;
  var lastEnd  = 0;
  var d;
  while ((d = dateExpr.exec(before)) !== null) {
    lastEnd = d.index + d[0].length;
  }

  var candidate = before.slice(lastEnd).replace(/^[\s,;.:–—-]+/, '').trim();
  var words     = candidate.split(/\s+/).filter(function (w) { return w.length > 0; });

  // No date preceded it (the entry opens the section) — fall back to the
  // trailing words, which is all the context there is.
  if (words.length === 0) {
    words = before.trim().split(/\s+/).filter(function (w) { return w.length > 0; });
  }

  return words.slice(-4).join(' ');
}


/**
 * Takes the leading date expression from the text following a "(no classes)"
 * marker, stopping before the next entry's name. Recognises a single date, a
 * range, and an "and"-joined pair.
 */
function extractHolidayDates7_(after) {
  var DAY   = '[A-Za-z]+\\.?\\s+\\d{1,2}(?:,\\s*\\d{4})?';
  var RANGE = '(?:\\s*[-–—]\\s*(?:[A-Za-z]+\\.?\\s+)?\\d{1,2}(?:,\\s*\\d{4})?)?';
  var AND   = '(?:\\s+and\\s+' + DAY + ')*';

  var m = after.match(new RegExp('^\\s*(' + DAY + RANGE + AND + ')'));
  return m ? m[1].trim() : '';
}


/**
 * Pulls no-class days out of the "<Semester> <Year> at a Glance" section.
 *
 * Handles the three shapes the registrar uses:
 *   "Labor Day (no classes): September 7, 2026"
 *   "Thanksgiving Holiday (no classes): November 23-November 29"
 *   "America250 (no classes): July 2, 2026 and July 6, 2026"
 *
 * @returns {Array<{name, start: number, end: number, fullWeek: boolean}>}
 */
function parseHolidays7_(html, year) {
  var text     = stripHtml7_(html);
  var holidays = [];

  // Locate every "(no classes)" marker first, then slice the name and value
  // around it. A single regex with a bounded value group does not work here:
  // the "at a Glance" entries run together into one line after tag stripping,
  // so a greedy value swallowed the NEXT holiday's name and dates too, fusing
  // Labor Day and Thanksgiving into one bogus full-week break.
  var marker  = /\(no\s*classes\)\s*:?\s*/gi;
  var markers = [];
  var mk;
  while ((mk = marker.exec(text)) !== null) {
    markers.push({ start: mk.index, end: mk.index + mk[0].length });
  }

  for (var i = 0; i < markers.length; i++) {
    var before = text.slice(0, markers[i].start);
    var after  = text.slice(markers[i].end,
                            i + 1 < markers.length ? markers[i + 1].start : text.length);

    var name  = extractHolidayName7_(before);
    var value = extractHolidayDates7_(after);
    if (!name || !value) continue;

    // "July 2, 2026 and July 6, 2026" — two separate single days.
    var parts = value.split(/\s+and\s+/i);
    for (var p = 0; p < parts.length; p++) {
      var part = parts[p].trim();

      // A range: "November 23-November 29" or "March 16-20".
      var range = part.match(/^(.+?)\s*[-–—]\s*(.+)$/);
      if (range) {
        var rStart = parseCalendarDate7_(range[1], year);
        var rEndRaw = range[2].trim();
        // A bare day number inherits the month from the range's start.
        if (/^\d{1,2}\b/.test(rEndRaw) && rStart) {
          rEndRaw = MONTHS_7[rStart.getMonth()] + ' ' + rEndRaw;
        }
        var rEnd = parseCalendarDate7_(rEndRaw, year);
        if (rStart && rEnd && rEnd >= rStart) {
          holidays.push({
            name:  name,
            start: rStart.getTime(),
            end:   rEnd.getTime(),
            fullWeek: coversFullWeek7_(rStart, rEnd)
          });
          continue;
        }
      }

      var single = parseCalendarDate7_(part, year);
      if (single) {
        holidays.push({
          name: name, start: single.getTime(), end: single.getTime(), fullWeek: false
        });
      }
    }
  }

  return holidays;
}


/**
 * True when a break covers a whole instructional week — Monday through at
 * least Friday. Thanksgiving (Mon–Sun) and Spring Break (Mon–Fri) both
 * qualify; Labor Day does not. Only full weeks shift the module schedule.
 */
function coversFullWeek7_(start, end) {
  var monday = mondayOf7_(start);
  var friday = new Date(monday.getTime());
  friday.setDate(friday.getDate() + 4);
  return start.getTime() <= monday.getTime() && end.getTime() >= friday.getTime();
}


/** Monday of the week containing d (weeks run Monday–Sunday). */
function mondayOf7_(d) {
  var out   = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  var day   = out.getDay();               // 0 = Sunday
  var shift = (day === 0) ? -6 : (1 - day);
  out.setDate(out.getDate() + shift);
  return out;
}


/**
 * Sidebar-callable. Fetches the academic calendar for a semester and year.
 *
 * Three tiers, because the registrar publishes only about a year ahead and
 * Blueprints are routinely built further out than that:
 *   1. the per-semester page — full session table plus break dates
 *   2. the 5-year calendar   — regular semesters only, no breaks
 *   3. nothing               — caller falls back to manual date pickers
 *
 * @returns {Object} {tier, sessions[], holidays[], warnings[], error}
 */
function fetchAcademicCalendar7(semester, year) {
  var out = { tier: 0, sessions: [], holidays: [], warnings: [], error: '', url: '' };

  var sem = String(semester || '').toLowerCase();
  if (['spring', 'summer', 'fall'].indexOf(sem) === -1) {
    out.error = 'Please choose Spring, Summer, or Fall.';
    return out;
  }

  if (!/^\d{4}$/.test(String(year))) {
    out.error = 'Enter a four-digit year, e.g. 2027.';
    return out;
  }
  var yr = parseInt(year, 10);
  if (yr < 2000 || yr > 2100) {
    out.error = 'Enter a year between 2000 and 2100.';
    return out;
  }

  // The registrar uses two URL shapes and is not consistent about which — the
  // summer calendars sit at the short form while fall sits at the long one, so
  // a single-pattern lookup reports real pages as missing.
  var candidates = [
    CALENDAR_BASE_7 + 'boise-state-academic-calendars/' + sem + '-' + yr + '-academic-calendar/',
    CALENDAR_BASE_7 + sem + '-' + yr + '-academic-calendar/'
  ];

  for (var i = 0; i < candidates.length; i++) {
    var html = fetchPage7_(candidates[i]);
    if (!html) continue;

    var sessions = parseSessionTable7_(html, yr);
    if (sessions.length === 0) continue;

    out.tier     = 1;
    out.url      = candidates[i];
    out.sessions = sessions;
    out.holidays = parseHolidays7_(html, yr);

    if (out.holidays.length === 0) {
      out.warnings.push('No "no classes" dates were found on the calendar page. ' +
                        'Check breaks by hand before running.');
    }
    Logger.log('fetchAcademicCalendar7: tier 1, %s session(s), %s holiday(s) from %s',
               sessions.length, out.holidays.length, candidates[i]);
    return out;
  }

  // Tier 2 — the 5-year calendar. Regular semesters only, and it carries no
  // usable break dates for years past the current one.
  if (sem === 'summer') {
    out.tier  = 3;
    out.error = 'No Summer ' + yr + ' calendar has been published yet, and the 5-year ' +
                'calendar lists summer only as a single block with no session dates. ' +
                'Enter the first and last day of the course below.';
    return out;
  }

  var fiveYear = fetchPage7_(FIVE_YEAR_URL_7);
  if (fiveYear && new RegExp(sem + '\\s+' + yr, 'i').test(stripHtml7_(fiveYear))) {
    out.tier = 2;
    out.url  = FIVE_YEAR_URL_7;
    out.warnings.push('No ' + sem + ' ' + yr + ' calendar page has been published yet. ' +
                      'Using the 5-year calendar, which lists the regular semester only — ' +
                      'no 5-, 7-, or 10-week sessions.');
    out.warnings.push('The 5-year calendar does not list break dates, so Thanksgiving ' +
                      'and Spring Break will NOT be skipped automatically. Adjust the ' +
                      'dates by hand below.');
    return out;
  }

  out.tier  = 3;
  out.error = 'No academic calendar could be found for ' + sem + ' ' + yr + '. ' +
              'Enter the first and last day of the course below.';
  return out;
}


/** Fetches a page, returning its HTML or '' on any non-200 or transport error. */
function fetchPage7_(url) {
  try {
    var resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects:    true
    });
    if (resp.getResponseCode() !== 200) {
      Logger.log('fetchPage7_: %s returned %s', url, resp.getResponseCode());
      return '';
    }
    return resp.getContentText();
  } catch (e) {
    Logger.log('fetchPage7_ error on %s: %s', url, e.message);
    return '';
  }
}


// ── MODULE DATE GENERATION ───────────────────────────────────

/**
 * Sidebar-callable. Lays module date ranges over a session's instructional
 * weeks, skipping any week wholly consumed by a break.
 *
 * Uneven division leaves the remainder UNASSIGNED rather than padding a
 * module: 7 modules across 15 weeks get 2 weeks each and the course dates
 * stop a week short, which the caller reports.
 *
 * @returns {Object} {modules[], weeksPerModule, skippedWeeks[], notes[], warnings[]}
 */
function buildModuleDates7(startMs, endMs, moduleCount, holidays) {
  var out = {
    modules: [], weeksPerModule: 0, skippedWeeks: [],
    notes: [], warnings: [], error: ''
  };

  var start = new Date(startMs);
  var end   = new Date(endMs);
  var count = parseInt(moduleCount, 10);

  if (isNaN(start.getTime()) || isNaN(end.getTime()) || end <= start) {
    out.error = 'The course end date must fall after the start date.';
    return out;
  }
  if (!count || count < 1) {
    out.error = 'No modules to date.';
    return out;
  }

  holidays = holidays || [];

  // Build the pool of Monday–Sunday weeks the session spans, dropping any week
  // a full-week break consumes and noting single-day holidays in place.
  var weeks  = [];
  var cursor = mondayOf7_(start);

  while (cursor.getTime() <= end.getTime()) {
    var weekStart = new Date(cursor.getTime());
    var weekEnd   = new Date(cursor.getTime());
    weekEnd.setDate(weekEnd.getDate() + 6);

    var fullWeekBreak = null;
    var dayOffNames   = [];

    for (var h = 0; h < holidays.length; h++) {
      var hol   = holidays[h];
      var hs    = new Date(hol.start);
      var he    = new Date(hol.end);
      var overlaps = hs.getTime() <= weekEnd.getTime() && he.getTime() >= weekStart.getTime();
      if (!overlaps) continue;

      if (hol.fullWeek) fullWeekBreak = hol.name;
      else              dayOffNames.push(hol.name);
    }

    if (fullWeekBreak) {
      out.skippedWeeks.push({
        name:  fullWeekBreak,
        start: weekStart.getTime(),
        end:   weekEnd.getTime()
      });
    } else {
      weeks.push({ start: weekStart, end: weekEnd, daysOff: dayOffNames });
    }

    cursor.setDate(cursor.getDate() + 7);
  }

  if (weeks.length < count) {
    out.warnings.push('This session has only ' + weeks.length + ' instructional week(s) ' +
                      'but the Development tab has ' + count + ' modules. Each module has ' +
                      'been given one week and the last ' + (count - weeks.length) +
                      ' have no dates — adjust them by hand.');
  }

  var perModule = Math.max(1, Math.floor(weeks.length / count));
  out.weeksPerModule = perModule;

  for (var m = 0; m < count; m++) {
    var firstWeek = weeks[m * perModule];
    var lastWeek  = weeks[Math.min(m * perModule + perModule - 1, weeks.length - 1)];

    if (!firstWeek) {
      out.modules.push({ index: m, start: null, end: null });
      continue;
    }

    var daysOff = [];
    for (var w = m * perModule; w <= Math.min(m * perModule + perModule - 1, weeks.length - 1); w++) {
      if (weeks[w]) daysOff = daysOff.concat(weeks[w].daysOff);
    }
    if (daysOff.length > 0) {
      out.notes.push({
        module: m + 1,
        text:   daysOff.length + ' no-class day(s) fall in this module: ' + daysOff.join(', ') + '.'
      });
    }

    out.modules.push({
      index: m,
      start: firstWeek.start.getTime(),
      end:   lastWeek.end.getTime()
    });
  }

  var used = perModule * count;
  if (weeks.length > used) {
    out.warnings.push((weeks.length - used) + ' instructional week(s) at the end of the ' +
                      'session are not covered by any module. Adjust the last module\'s ' +
                      'end date if it should run longer.');
  }

  return out;
}


/** Formats a date as "Aug 24", using the AP-style month names above. */
function formatModuleDate7(ms) {
  var d = new Date(ms);
  return MONTHS_7[d.getMonth()] + ' ' + d.getDate();
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
