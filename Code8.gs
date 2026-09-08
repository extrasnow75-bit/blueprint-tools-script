// ============================================================
// Blueprint Tools — Code8.gs
// Specialty Tool: Add Module Dates &/or Holiday Modules. Fills in module
// start/end dates on the Development tab's H2 headings from the Boise State
// registrar's academic calendar, and optionally inserts unnumbered Spring
// Break / Thanksgiving Break modules.
// ------------------------------------------------------------
// Last updated on 2026-09-07 at 19:51 MDT
// ------------------------------------------------------------
//
// Split out of the "Add Module Titles & Module Dates (Beta)" tool — see project
// memory project_blueprint_designmap_dev_tab.md, Phase 2. That tool is now
// retired: its title half moved into "Design Map → Dev Tab" (Code9.gs /
// Sidebar9.html) and this is its standalone dates-only replacement.
//
// Runs AFTER "Add Activity Titles, Tools, Due Date Headers, & Times", which is
// what creates the H2 headings this tool writes into.
//
// Relies on shared helpers defined elsewhere in the same flat GAS namespace —
// do NOT redefine any of these here:
//   getDevelopmentTabBody                 (Code2.gs)
//   classifyHeading7_,
//   scanDevelopmentHeadings7_,
//   MODULE_PREFIX_RE_7                    (Code9.gs — they used to live in
//     Code7.gs and moved with the title half when that file was deleted)
// The START_PLACEHOLDER_7 / END_PLACEHOLDER_7 constants are defined below, in
// this file, but their consumer classifyHeading7_ is the Code9.gs one.
// ============================================================


// ── CONSTANTS ────────────────────────────────────────────────

// The two placeholder substrings this tool replaces. Load-bearing contract
// with the Blueprint template, same as Code9.gs's TITLE_PLACEHOLDER_7: the
// tool writes a field ONLY where its placeholder is still present, which is
// what makes "skip and report" work without ever clobbering a designer's own
// text. classifyHeading7_ (Code9.gs) reads these to compute dateIsPlaceholder.
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

var CALENDAR_BASE_7 = 'https://www.boisestate.edu/registrar/';
var FIVE_YEAR_URL_7 = CALENDAR_BASE_7 + 'boise-state-academic-calendars/5-year-academic-calendar/';

// The separator inside a module heading's "(start – end)". An EN DASH, because
// it spans a range. The Blueprint template ships a hyphen; every heading this
// tool writes normalises to the en dash, placeholder branch included, so a
// document does not end up with both.
var DATE_RANGE_SEP_8 = '–';


// ── HOLIDAY MODULES ──────────────────────────────────────────
// Optional, unnumbered modules for a full-week break. They are NOT numbered on
// purpose: an unnumbered H2 is invisible to MODULE_PREFIX_RE_7, to Code.gs's
// countExistingModules/deleteModule, and to this file's own scan — so adding
// one cannot disturb any module counting elsewhere in the suite, and a later
// re-run of the first tool will not try to fill it with activity slots.
//
// buildModuleDates7 already DROPS a full-week break from the week pool before
// dividing weeks among modules, so a holiday module only makes visible a week
// the schedule was already skipping. It does not shift any module's dates.
var HOLIDAY_SPECS_8 = {
  spring: {
    name:  'Spring Break',
    // U+1F60E SMILING FACE WITH SUNGLASSES, hard against the word, exactly as
    // it appears in the Blueprint template.
    emoji: '😎',
    // Matches the scraped break's name, which the registrar spells various ways.
    match: /spring\s*break/i,
    // Where it goes when no calendar dates are available: the middle of the
    // course. ceil(n/2) gives "after Module 4" in a 7-module course and
    // "after Module 8" in a 15-module one.
    fallback: function (n) { return Math.ceil(n / 2); }
  },
  thanksgiving: {
    name:  'Thanksgiving Break',
    emoji: '😎',
    match: /thanksgiving/i,
    // A 15-week fall course puts Thanksgiving near the end; a second-7-week
    // fall session starts mid-October and hits it two modules in. Module count
    // is the only signal available here, so it is the one used.
    fallback: function (n) { return n >= 12 ? 11 : 2; }
  }
};

// The developer instruction that rides under a holiday module's heading. The
// lead is bold red — it is a note TO the course developer, not course content —
// and the quoted sentence is the text they are being asked to place.
var HOLIDAY_NOTE_LEAD_8 = 'Add a text header that reads,';
var HOLIDAY_NOTE_BODY_8 = ' “Nothing is due this week. Enjoy your time off!”';

// What a holiday heading carries when the user asked for no dates, or none
// could be found. Deliberately the template's own placeholder wording, so a
// later run of this tool with a calendar in hand fills it in like any other.
var HOLIDAY_DATE_PLACEHOLDER_8 =
  START_PLACEHOLDER_7 + ' ' + DATE_RANGE_SEP_8 + ' ' + END_PLACEHOLDER_7;


// ── SIDEBAR OPENER ───────────────────────────────────────────

function showModuleDatesSidebar8() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar8')
    .setTitle('Specialty Tool: Add Module Dates &/or Holiday Modules')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}


// ── SIDEBAR DATA ─────────────────────────────────────────────

/**
 * Sidebar-callable. Builds the per-module panel data this tool needs: just the
 * Development tab's numbered module headings and whether each already has
 * dates. Unlike getDesignMapSidebarData9 (Code9.gs), this does NOT touch
 * the Design tab — a dates-only tool has no reason to require one, and
 * requiring it would block a course whose Design tab is missing or renamed
 * from adding dates at all.
 */
function getModuleDatesSidebarData8() {
  var result = { modules: [], existingHolidays: [], error: '' };

  try {
    var doc     = DocumentApp.getActiveDocument();
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

    for (var i = 0; i < headings.length; i++) {
      var info = headings[i];
      result.modules.push({
        num:              info.num,
        displayLabel:     info.displayLabel,
        canWriteDates:    info.dateIsPlaceholder,
        // True only when there is a real "(...)" to overwrite — a heading a
        // designer rewrote by hand with no trailing parenthetical at all
        // (info.datePart === null) has nowhere to put a date range, no matter
        // what the user chooses, so it is neither "writable" nor "overwritable".
        hasExistingDates: info.datePart !== null && !info.dateIsPlaceholder,
        currentDates:     info.dateIsPlaceholder ? '' : (info.datePart || '')
      });
    }

    // Which holiday modules are already in the document. The sidebar uses this
    // to pre-tick and disable those boxes, so the user is told up front rather
    // than finding out from the summary that the run skipped them.
    result.existingHolidays = [];
    for (var key in HOLIDAY_SPECS_8) {
      if (!Object.prototype.hasOwnProperty.call(HOLIDAY_SPECS_8, key)) continue;
      if (hasHolidayModule8_(devBody, HOLIDAY_SPECS_8[key])) result.existingHolidays.push(key);
    }

    Logger.log('getModuleDatesSidebarData8: %s heading(s), %s holiday module(s) already present.',
               result.modules.length, result.existingHolidays.length);

  } catch (e) {
    Logger.log('getModuleDatesSidebarData8 error: ' + e.message);
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
    // Bounded, and without the surrounding \s* the class already covers. The
    // old /\s*[\d\s…]+\s*/ had three overlapping whitespace quantifiers, so an
    // unterminated <sup> followed by a long run of spaces backtracked cubically
    // — 3,200 spaces took 7.7s, and a page of them would never finish. Only
    // parseHolidays7_ passes the whole fetched page through here, so it is
    // reachable by anything that can control those bytes.
    .replace(/<sup[^>]*>[\d\s.,*†‡]{1,20}<\/sup>/gi, '')
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
 * Reads one labelled field out of a 5-year calendar cell.
 *
 * A cell is a run of "<label>: <value>" pairs separated by <br>, which tag
 * stripping flattens onto one line, so a value runs until the next known label.
 *
 * Commas are dropped unless a four-digit year follows. The day-of-week prefix
 * ("Mon, Aug 23") and one typo in the 2029-2030 row ("Mon, Aug, 20") both put a
 * comma where parseCalendarDate7_ expects whitespace between month and day;
 * "November 23, 2026" keeps its comma so the explicit-year path still fires.
 */
function fiveYearField7_(cellText, label) {
  var lab = label.replace(/\s+/g, '\\s+');
  var re  = new RegExp('\\b' + lab + '\\s*:\\s*([\\s\\S]*?)(?=\\b(?:start|end|' +
                       'thanksgiving|spring\\s+break|commencement)\\s*:|$)', 'i');
  var m = cellText.match(re);
  if (!m) return '';
  return m[1].replace(/,(?!\s*\d{4}\b)/g, ' ').replace(/\s+/g, ' ').trim();
}


/**
 * Turns a 5-year break field ("Nov 22-26", "March 16-20") into a holiday record
 * shaped exactly like the ones parseHolidays7_ builds, so buildModuleDates7 can
 * treat both sources alike. A bare end day inherits the start's month, as on
 * the per-semester pages.
 */
function fiveYearBreak7_(name, text, yr) {
  var range = String(text).match(/^(.+?)\s*[-\u2013\u2014]\s*(.+)$/);
  if (!range) return null;

  var start = parseCalendarDate7_(range[1], yr);
  if (!start) return null;

  var endRaw = range[2].trim();
  if (/^\d{1,2}\b/.test(endRaw)) endRaw = MONTHS_7[start.getMonth()] + ' ' + endRaw;

  var end = parseCalendarDate7_(endRaw, yr);
  if (!end || end.getTime() < start.getTime()) return null;

  return {
    name:     name,
    start:    start.getTime(),
    end:      end.getTime(),
    fullWeek: coversFullWeek7_(start, end)
  };
}


/**
 * Pulls one semester's dates out of the 5-year calendar.
 *
 * That table is organised by ACADEMIC year ("2027-2028") with a column per
 * semester, so neither "Fall 2027" nor "Spring 2028" appears anywhere on the
 * page. Tier 2 used to gate on finding exactly those strings, which meant it
 * could never fire: every unpublished term fell through to the manual pickers
 * while reporting that no calendar existed.
 *
 * Fall of "Y-(Y+1)" is calendar year Y and Spring of "(Y-1)-Y" is calendar
 * year Y, so the requested year is correct for every date in the chosen column.
 *
 * @returns {?{start: Date, end: Date, holiday: ?Object, tbdBreak: string}}
 */
function parseFiveYearRow7_(html, sem, yr) {
  var tables = html.match(/<table[\s\S]*?<\/table>/gi) || [];

  for (var t = 0; t < tables.length; t++) {
    var rows = tables[t].match(/<tr[\s\S]*?<\/tr>/gi) || [];
    if (rows.length < 2) continue;

    // Fall and Spring head their columns "<Semester> Semester"; summer is
    // "Summer Sessions", plural, because it holds several overlapping ones.
    var heading = (sem === 'summer') ? 'summer sessions' : sem + ' semester';
    var header  = (rows[0].match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || []).map(stripHtml7_);
    var col = -1;
    for (var c = 0; c < header.length; c++) {
      if (header[c].toLowerCase().indexOf(heading) !== -1) { col = c; break; }
    }
    if (col < 0) continue;

    // A row spans two calendar years: only Fall falls in the first of them.
    // Spring 2028 and Summer 2028 both live in the 2027-2028 row.
    var first = (sem === 'fall') ? yr : yr - 1;
    var label = new RegExp('^\\s*' + first + '\\s*[-\u2013\u2014]\\s*(?:' + (first + 1) +
                           '|' + String(first + 1).slice(2) + ')\\s*$');

    for (var r = 1; r < rows.length; r++) {
      var cells = (rows[r].match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || []).map(stripHtml7_);
      if (cells.length <= col || !label.test(cells[0])) continue;

      var cell  = cells[col];
      var start = parseCalendarDate7_(fiveYearField7_(cell, 'start'), yr);
      var end   = parseCalendarDate7_(fiveYearField7_(cell, 'end'),   yr);
      if (!start || !end) return null;

      // Fall lists Thanksgiving; Spring lists Spring Break, which the registrar
      // marks TBD for every year after the current one. Summer lists no break
      // at all, so it must not be reported as a missing one.
      var breakName = (sem === 'fall') ? 'Thanksgiving'
                    : (sem === 'spring') ? 'Spring Break' : '';
      var breakText = breakName ? fiveYearField7_(cell, breakName) : '';
      var holiday   = breakText && !/tbd/i.test(breakText)
                        ? fiveYearBreak7_(breakName, breakText, yr)
                        : null;

      return {
        start:    start,
        end:      end,
        holiday:  holiday,
        tbdBreak: (breakName && !holiday) ? breakName : ''
      };
    }
  }

  return null;
}


/**
 * Sidebar-callable. Fetches the academic calendar for a semester and year.
 *
 * Three tiers, because the registrar publishes only about a year ahead and
 * Blueprints are routinely built further out than that:
 *   1. the per-semester page — full session table, every no-class day
 *   2. the 5-year calendar   — regular semester only, plus its one big break
 *   3. nothing               — caller falls back to manual date pickers,
 *                              pre-filled from suggestedStart/End when summer's
 *                              block was readable
 *
 * @returns {Object} {tier, sessions[], holidays[], warnings[], error,
 *                    suggestedStart, suggestedEnd}
 */
function fetchAcademicCalendar7(semester, year) {
  var out = {
    tier: 0, sessions: [], holidays: [], warnings: [], error: '', url: '',
    // Tier 3 only: dates to pre-fill the manual pickers with, when something
    // usable was found even though no session table could be built.
    suggestedStart: 0, suggestedEnd: 0
  };

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
  var fiveYear = fetchPage7_(FIVE_YEAR_URL_7);

  // Summer never reaches tier 2: the 5-year table holds one undifferentiated
  // block, so there is nothing to put in a session dropdown. It does carry the
  // block's own first and last day, though, which is worth pre-filling into the
  // manual pickers rather than leaving the user to look it up.
  if (sem === 'summer') {
    var summer = fiveYear ? parseFiveYearRow7_(fiveYear, 'summer', yr) : null;

    out.tier  = 3;
    out.error = 'No Summer ' + yr + ' calendar has been published yet, and the 5-year ' +
                'calendar lists summer as a single block with no 3-, 5-, 7-, or 10-week ' +
                'session dates. ' +
                (summer
                  ? 'The dates below are that whole block, from the 5-year calendar — ' +
                    'shorten them to your session before generating.'
                  : 'Enter the first and last day of the course below.');

    if (summer) {
      out.suggestedStart = summer.start.getTime();
      out.suggestedEnd   = summer.end.getTime();
      out.url            = FIVE_YEAR_URL_7;
    }
    return out;
  }

  var row = fiveYear ? parseFiveYearRow7_(fiveYear, sem, yr) : null;
  if (row) {
    out.tier     = 2;
    out.url      = FIVE_YEAR_URL_7;
    out.sessions = [{
      session: 'Regular semester',
      start:   row.start.getTime(),
      end:     row.end.getTime()
    }];
    if (row.holiday) out.holidays = [row.holiday];

    out.warnings.push('No ' + sem + ' ' + yr + ' calendar page has been published yet. ' +
                      'Using the 5-year calendar, which lists the regular semester only — ' +
                      'no 5-, 7-, or 10-week sessions.');
    // Tier 1 reports every "(no classes)" day; the 5-year table lists only the
    // one long break, so silence here is not evidence there are no others.
    out.warnings.push('The 5-year calendar does not list single days off, so Labor Day, ' +
                      'MLK Day and similar are NOT reported below. Check them by hand.');
    if (row.tbdBreak) {
      out.warnings.push('The 5-year calendar shows ' + row.tbdBreak + ' as TBD for this ' +
                        'year, so that week will NOT be skipped automatically. Adjust the ' +
                        'dates by hand once the registrar publishes it.');
    }

    Logger.log('fetchAcademicCalendar7: tier 2, %s - %s, break: %s',
               row.start, row.end, row.holiday ? row.holiday.name : (row.tbdBreak || 'none'));
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
 * Sidebar-callable. Writes the generated start/end dates into the Development
 * tab's module headings.
 *
 * Dates-only extract of the retired Beta tool's applyModuleTitlesAndDates7.
 * Its title half went the other way, into applyDesignMapToDevTab9 (Code9.gs);
 * this is the standalone dates version.
 *
 * Three outcomes per heading:
 *   - blank placeholder ("start date - end date")  → always filled in. Writes
 *     by replacing the placeholder substrings rather than calling setText on
 *     the paragraph, which keeps the parentheses, hyphen and heading's Arial
 *     17 bold untouched.
 *   - already has a real date range                → filled in ONLY when the
 *     sidebar marked this module for overwrite (params.overwrite[num] truthy).
 *     There is no fixed placeholder word left to target here, so the whole
 *     trailing "(...)" is replaced as a unit instead.
 *   - no trailing "(...)" at all (a heading rewritten by hand)  → never
 *     written, regardless of params — there is nowhere to put a date range.
 *
 * The user, not this function, decides overwrite vs. preserve per module in
 * the sidebar; params.overwrite carries that choice. Defaulting to preserve
 * there (not here) is what keeps a re-run from silently destroying a
 * designer's own edit unless the user explicitly asks to replace it.
 *
 * @param {Object} params
 *   .dates     {Object}  module number → {start: ms, end: ms}
 *   .overwrite {Object}  module number → true if an existing date range
 *                        should be replaced (ignored for blank placeholders,
 *                        which are always written; ignored for modules with
 *                        no trailing "(...)" at all, which can never be)
 * @returns {string} plain-text summary for the sidebar
 */
/** "Aug 24 – Sept 6" from a {start, end} pair. One place, so the separator
 *  cannot drift between the placeholder and overwrite branches again. */
function moduleRange8_(d) {
  return formatModuleDate7(d.start) + ' ' + DATE_RANGE_SEP_8 + ' ' + formatModuleDate7(d.end);
}


/**
 * True when the Development tab already carries a holiday module for `spec`.
 *
 * Checked against H2 headings only. A "Spring Break" mentioned in someone's
 * directions paragraph is not a module and must not suppress the insert.
 */
function hasHolidayModule8_(devBody, spec) {
  return findHolidayModule8_(devBody, spec) !== null;
}


/**
 * The H2 paragraph of an existing holiday module for `spec`, or null.
 *
 * Split out of hasHolidayModule8_ so a second run can write dates INTO a
 * holiday module added by an earlier one — a master course gets its breaks
 * with the "(start date – end date)" placeholder left in, and the dates only
 * arrive once the course is scheduled and the calendar is looked up.
 */
function findHolidayModule8_(devBody, spec) {
  var H2    = DocumentApp.ParagraphHeading.HEADING2;
  var paras = devBody.getParagraphs();
  for (var i = 0; i < paras.length; i++) {
    if (paras[i].getHeading() !== H2) continue;
    if (spec.match.test(paras[i].getText())) return paras[i];
  }
  return null;
}


/**
 * True when this heading's trailing "(...)" is still the untouched placeholder.
 *
 * Matches on START_PLACEHOLDER_7 rather than on the whole placeholder string so
 * that a heading someone pasted from the template — whose separator is the
 * template's hyphen, not this tool's en dash — is still recognised.
 */
function holidayDateIsPlaceholder8_(text) {
  var m = String(text).match(/\(([^()]*)\)\s*$/);
  return !!m && m[1].toLowerCase().indexOf(START_PLACEHOLDER_7) !== -1;
}


/**
 * Inserts the requested holiday modules into the Development tab.
 *
 * @param {Body}  devBody
 * @param {Array} headings   scanDevelopmentHeadings7_ output, captured BEFORE
 *                           any insert — used only for its element references
 *                           and module numbers, never its childIndex values,
 *                           which go stale the moment anything is inserted.
 * @param {Array} requests   [{key, start, end, afterModule}] from the sidebar.
 *                           start/end may be null, meaning "leave the
 *                           (start date – end date) placeholder in".
 * @param {boolean} overwriting  true when the run is overwriting the numbered
 *                           modules' dates. A holiday module whose dates are
 *                           already real is only rewritten on such a run;
 *                           one still showing the placeholder is always filled.
 * @returns {{added: string[], updated: string[], skipped: string[]}}
 */
function insertHolidayModules8_(devBody, headings, requests, overwriting) {
  var out = { added: [], updated: [], skipped: [] };
  if (!requests || !requests.length) return out;

  var H2     = DocumentApp.ParagraphHeading.HEADING2;
  var NORMAL = DocumentApp.ParagraphHeading.NORMAL;

  // Element references, not indices: these survive the inserts below, whereas
  // every childIndex captured now is invalidated by the first one.
  var anchors = {};   // module number → the H2 element of that module
  var refPara = null; // any module heading, used as the formatting model
  for (var h = 0; h < headings.length; h++) {
    var el = devBody.getChild(headings[h].childIndex);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    anchors[headings[h].num] = el.asParagraph();
    if (!refPara) refPara = el.asParagraph();
  }
  if (!refPara) return out;

  var maxModule = 0;
  for (var m = 0; m < headings.length; m++) {
    if (headings[m].num > maxModule) maxModule = headings[m].num;
  }

  // Copied wholesale from a real module heading rather than hardcoded: the
  // template's H2 look (font, size, colour, spacing) is not documented
  // anywhere, and a hand-built approximation would be visibly off beside it.
  var refAttrs = refPara.getAttributes();

  for (var r = 0; r < requests.length; r++) {
    var req  = requests[r] || {};
    var spec = HOLIDAY_SPECS_8[req.key];
    if (!spec) continue;

    var hasDates = !!(req.start && req.end);
    var range = hasDates
      ? moduleRange8_({ start: req.start, end: req.end })
      : HOLIDAY_DATE_PLACEHOLDER_8;

    // Already in the document? Then this run either writes dates into it or
    // leaves it alone — it never adds a second one.
    var existing = findHolidayModule8_(devBody, spec);
    if (existing) {
      if (!hasDates) {
        out.skipped.push(spec.name + ' — no dates to add');
      } else if (holidayDateIsPlaceholder8_(existing.getText())) {
        // The case this whole branch exists for: added dateless on an earlier
        // run, dated now that the academic calendar has been looked up.
        existing.replaceText('\\([^()]*\\)\\s*$', '(' + range + ')');
        out.updated.push(spec.name + ' — dates filled in (' + range + ')');
      } else if (overwriting) {
        existing.replaceText('\\([^()]*\\)\\s*$', '(' + range + ')');
        out.updated.push(spec.name + ' — dates overwritten (' + range + ')');
      } else {
        out.skipped.push(spec.name + ' — already dated, kept as-is');
      }
      Logger.log('insertHolidayModules8_: %s already present (hasDates=%s).', spec.name, hasDates);
      continue;
    }

    // Where it goes. The sidebar computes afterModule from the real calendar
    // when it has one; the spec's fallback covers a run with no dates.
    var after = parseInt(req.afterModule, 10);
    if (!after || after < 1 || after > maxModule) after = spec.fallback(maxModule);
    if (after > maxModule) after = maxModule;

    // Insert BEFORE the following module's heading. The last module has no
    // follower, so the block goes at the end of the tab.
    var next = anchors[after + 1];
    var at   = next ? devBody.getChildIndex(next) : devBody.getNumChildren();

    // Built last-first at a FIXED index so they land in reading order: heading,
    // note, then a blank paragraph separating this block from what follows.
    var spacer = devBody.insertParagraph(at, '');
    spacer.setHeading(NORMAL);
    zeroIndent_(spacer);

    var note = devBody.insertParagraph(at, HOLIDAY_NOTE_LEAD_8 + HOLIDAY_NOTE_BODY_8);
    note.setHeading(NORMAL);
    zeroIndent_(note);
    var noteText = note.editAsText();
    noteText.setFontFamily(FONT);
    noteText.setFontSize(11);
    noteText.setBold(false);
    noteText.setItalic(false);
    noteText.setForegroundColor(BLACK);
    // The lead is an instruction to the course developer, not course content —
    // same bold-red convention the tool suite uses everywhere else for that.
    noteText.setBold(0, HOLIDAY_NOTE_LEAD_8.length - 1, true);
    noteText.setForegroundColor(0, HOLIDAY_NOTE_LEAD_8.length - 1, RED);

    var head = devBody.insertParagraph(at, spec.name + spec.emoji + ' (' + range + ')');
    // setAttributes carries HEADING across with everything else, but it is set
    // explicitly first so the paragraph is a real H2 even if the reference
    // heading somehow is not.
    head.setHeading(H2);
    head.setAttributes(refAttrs);
    // Deliberately NOT zeroIndent_'d, unlike the two lines above it. The
    // indents just copied from a real module heading ARE the template's, and
    // forcing them to zero would leave this heading sitting differently from
    // every numbered module around it.

    out.added.push(spec.name + ' (' + range + ') — after Module ' + after);
    Logger.log('insertHolidayModules8_: added %s after module %s at index %s.',
               spec.name, after, at);
  }

  return out;
}


function applyModuleDates8(params) {
  var dates     = (params && params.dates)     || {};
  var overwrite = (params && params.overwrite) || {};

  var doc     = DocumentApp.getActiveDocument();
  var devBody = getDevelopmentTabBody(doc);
  if (!devBody) throw new Error('Could not find a "Development" tab in this document.');

  var headings = scanDevelopmentHeadings7_(devBody);
  if (headings.length === 0) throw new Error('No numbered module headings found in the Development tab.');

  var filledIn    = 0;
  var overwritten = 0;
  var preserved   = [];
  var unsupported = [];

  var H2 = DocumentApp.ParagraphHeading.HEADING2;

  for (var i = 0; i < headings.length; i++) {
    var info = headings[i];
    var para = devBody.getChild(info.childIndex).asParagraph();

    // Guard against the document having shifted since the scan.
    if (para.getHeading() !== H2) continue;

    var moduleDates = dates[info.num];

    if (info.dateIsPlaceholder) {
      if (!moduleDates || !moduleDates.start || !moduleDates.end) continue;
      // Replaces the trailing "(...)" whole rather than swapping the two words
      // in place. The surgical version left the TEMPLATE's hyphen behind, so a
      // filled-in placeholder read "Aug 24 - Sept 6" while an overwritten
      // heading read "Aug 24 – Sept 6" — two separators in one document.
      // dateIsPlaceholder guarantees the parenthetical is the placeholder, so
      // there is nothing here worth preserving.
      para.replaceText('\\([^()]*\\)\\s*$', '(' + moduleRange8_(moduleDates) + ')');
      filledIn++;
      continue;
    }

    if (info.datePart === null) {
      unsupported.push(info.displayLabel);
      continue;
    }

    if (!overwrite[info.num] || !moduleDates || !moduleDates.start || !moduleDates.end) {
      preserved.push(info.displayLabel + ' — kept "(' + info.datePart + ')"');
      continue;
    }

    var newRange = moduleRange8_(moduleDates);
    // Replaces the trailing "(...)" as a whole, whatever it currently holds —
    // unlike the placeholder branch above, there is no fixed word ("start
    // date") to target once a real date range is already sitting there.
    para.replaceText('\\([^()]*\\)\\s*$', '(' + newRange + ')');
    overwritten++;
  }

  // Holiday modules go in AFTER the dates pass. Inserting first would shift
  // every childIndex in `headings`, which was captured before this loop began —
  // and those indices are what the loop above walks.
  // A holiday module that already carries real dates is only rewritten when the
  // user asked for the numbered modules to be overwritten too — there is no
  // separate preserve/overwrite control for the breaks, so they follow that one.
  var overwriting = false;
  for (var o in overwrite) {
    if (Object.prototype.hasOwnProperty.call(overwrite, o) && overwrite[o]) { overwriting = true; break; }
  }
  var holidayResult =
    insertHolidayModules8_(devBody, headings, params && params.holidays, overwriting);

  Logger.log('applyModuleDates8: %s filled in, %s overwritten, %s preserved, %s unsupported, ' +
             '%s holiday module(s) added, %s holiday module(s) re-dated.',
             filledIn, overwritten, preserved.length, unsupported.length,
             holidayResult.added.length, holidayResult.updated.length);

  var lines = ['✅ Development tab updated.', ''];
  lines.push('Newly filled in: ' + filledIn);
  lines.push('Overwritten with new dates: ' + overwritten);

  if (holidayResult.added.length > 0) {
    lines.push('', 'Holiday modules added (' + holidayResult.added.length + '):');
    for (var a = 0; a < holidayResult.added.length; a++) lines.push('  • ' + holidayResult.added[a]);
    lines.push('  Drag one to a different spot if it did not land where you expected.');
  }
  if (holidayResult.updated.length > 0) {
    lines.push('', 'Holiday modules already present, dates written (' +
                   holidayResult.updated.length + '):');
    for (var v = 0; v < holidayResult.updated.length; v++) lines.push('  • ' + holidayResult.updated[v]);
  }
  if (holidayResult.skipped.length > 0) {
    lines.push('', 'Holiday modules already present, left alone (' +
                   holidayResult.skipped.length + '):');
    for (var k = 0; k < holidayResult.skipped.length; k++) lines.push('  • ' + holidayResult.skipped[k]);
  }

  if (preserved.length > 0) {
    lines.push('', 'Kept as-is (' + preserved.length + ') — you chose to preserve these:');
    for (var p = 0; p < preserved.length; p++) lines.push('  • ' + preserved[p]);
  }

  if (unsupported.length > 0) {
    lines.push('', 'No date placeholder found (' + unsupported.length + ') — dates cannot be ' +
                   'added to these headings by this tool, since they have no trailing ' +
                   '"(...)" left to put a date range into:');
    for (var u = 0; u < unsupported.length; u++) lines.push('  • ' + unsupported[u]);
  }

  return lines.join('\n');
}
