// ============================================================
// Blueprint Tools — Code8.gs
// Add Module Dates: fills in module start/end dates on the
// Development tab's H2 headings from the Boise State registrar's
// academic calendar.
// ------------------------------------------------------------
// Last updated on 2026-09-04 at 23:18 MDT
// ------------------------------------------------------------
//
// Split out of the "Add Module Titles & Module Dates (Beta)" tool
// (Code7.gs / Sidebar7.html) — see project memory
// project_blueprint_designmap_dev_tab.md, Phase 2. The Beta tool is left in
// place for its title half; this is the standalone dates-only replacement.
//
// Runs AFTER "Add Activity Titles, Tools, Due Date Headers, & Times", which is
// what creates the H2 headings this tool writes into.
//
// Relies on shared helpers defined elsewhere in the same flat GAS namespace —
// do NOT redefine any of these here:
//   getDevelopmentTabBody         (Code2.gs)
//   classifyHeading7_,
//   scanDevelopmentHeadings7_,
//   MODULE_PREFIX_RE_7,
//   START_PLACEHOLDER_7 (below, but its dateIsPlaceholder consumer,
//     classifyHeading7_, lives in Code7.gs),
//   END_PLACEHOLDER_7                     (all Code7.gs)
// ============================================================


// ── CONSTANTS ────────────────────────────────────────────────

// The two placeholder substrings this tool replaces. Load-bearing contract
// with the Blueprint template, same as Code7.gs's TITLE_PLACEHOLDER_7: the
// tool writes a field ONLY where its placeholder is still present, which is
// what makes "skip and report" work without ever clobbering a designer's own
// text. classifyHeading7_ (Code7.gs) reads these to compute dateIsPlaceholder.
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


// ── SIDEBAR OPENER ───────────────────────────────────────────

function showModuleDatesSidebar8() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar8')
    .setTitle('Add Module Dates')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}


// ── SIDEBAR DATA ─────────────────────────────────────────────

/**
 * Sidebar-callable. Builds the per-module panel data this tool needs: just the
 * Development tab's numbered module headings and whether each already has
 * dates. Unlike getModuleTitlesSidebarData7 (Code7.gs), this does NOT touch
 * the Design tab — a dates-only tool has no reason to require one, and
 * requiring it would block a course whose Design tab is missing or renamed
 * from adding dates at all.
 */
function getModuleDatesSidebarData8() {
  var result = { modules: [], error: '' };

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
        num:           info.num,
        displayLabel:  info.displayLabel,
        canWriteDates: info.dateIsPlaceholder,
        currentDates:  info.dateIsPlaceholder ? '' : (info.datePart || '')
      });
    }

    Logger.log('getModuleDatesSidebarData8: %s heading(s).', result.modules.length);

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
 * Dates-only extract of applyModuleTitlesAndDates7 (Code7.gs) — that function
 * is left intact for the Beta tool's combined titles+dates run; this is the
 * standalone version for the new tool. Same write strategy: replaces the
 * placeholder substrings rather than calling setText on the paragraph, which
 * keeps the parentheses, hyphen and heading's Arial 17 bold untouched.
 *
 * Never overwrites. A heading whose date placeholder is already gone is
 * skipped and reported, so a re-run cannot destroy a designer's own edits.
 *
 * @param {Object} params
 *   .dates {Object}  module number → {start: ms, end: ms}
 * @returns {string} plain-text summary for the sidebar
 */
function applyModuleDates8(params) {
  var dates = (params && params.dates) || {};

  var doc     = DocumentApp.getActiveDocument();
  var devBody = getDevelopmentTabBody(doc);
  if (!devBody) throw new Error('Could not find a "Development" tab in this document.');

  var headings = scanDevelopmentHeadings7_(devBody);
  if (headings.length === 0) throw new Error('No numbered module headings found in the Development tab.');

  var datesWritten = 0;
  var datesSkipped = [];

  var H2 = DocumentApp.ParagraphHeading.HEADING2;

  for (var i = 0; i < headings.length; i++) {
    var info = headings[i];
    var para = devBody.getChild(info.childIndex).asParagraph();

    // Guard against the document having shifted since the scan.
    if (para.getHeading() !== H2) continue;

    var moduleDates = dates[info.num];
    if (!moduleDates || !moduleDates.start || !moduleDates.end) continue;

    if (info.dateIsPlaceholder) {
      para.replaceText('\b' + START_PLACEHOLDER_7 + '\b', formatModuleDate7(moduleDates.start));
      para.replaceText('\b' + END_PLACEHOLDER_7   + '\b', formatModuleDate7(moduleDates.end));
      datesWritten++;
    } else {
      datesSkipped.push(info.displayLabel + ' — already reads "(' + info.datePart + ')"');
    }
  }

  Logger.log('applyModuleDates8: %s date range(s) written; %s skipped.',
             datesWritten, datesSkipped.length);

  var lines = ['✅ Development tab updated.', ''];
  lines.push('Date ranges written: ' + datesWritten);

  if (datesSkipped.length > 0) {
    lines.push('', 'Dates left alone (' + datesSkipped.length + ') — these headings already ' +
                   'have dates, so nothing was overwritten:');
    for (var d = 0; d < datesSkipped.length; d++) lines.push('  • ' + datesSkipped[d]);
  }

  return lines.join('\n');
}
