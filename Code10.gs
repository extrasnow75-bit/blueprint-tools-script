/**
 * ================================================================
 * BLUEPRINT TOOLS  |  Run Log
 * ================================================================
 * Last updated on 2026-09-18 at 09:40 MDT
 * ================================================================
 * Records every Blueprint Tools run against this document — what ran,
 * when, who ran it, and what it did — so that someone who needs to roll
 * the document back can tell which version in File > Version history is
 * the one they want.
 *
 * WHERE THE LOG LIVES, AND WHY
 * The log is kept in Document Properties, NOT in a tab of the document.
 * Properties are stored against the document but are not part of its
 * content, so restoring the document to an earlier version leaves them
 * untouched. A log tab would roll back with everything else — losing
 * exactly the history someone opened it to read. Properties also need no
 * extra OAuth scope and add no revisions of their own.
 *
 * Two consequences, both accepted: the log is invisible until someone
 * opens the menu item, and a COPY of a Blueprint starts with an empty
 * log. A new document genuinely has no history, so that is correct.
 *
 * STORAGE SHAPE
 * One property per run, keyed BPLOG_<epochMillis>_<sequence>. A single
 * JSON array would cap out around 45 entries against the 9KB
 * per-value limit; one property per entry against the 500KB store
 * gives well over a thousand. Epoch millis are 13 digits until the year
 * 2286, so sorting the keys lexically sorts them chronologically.
 *
 * The sequence is a zero-padded counter, NOT a random suffix. Two runs
 * can land in the same millisecond, and randomness there fails twice
 * over: their relative order becomes arbitrary, which matters because
 * the merge logic below trusts the last key to be the newest entry; and
 * two equal suffixes collide, silently overwriting an entry. A counter
 * makes both impossible.
 *
 * LOGGING MUST NEVER BREAK A TOOL RUN. Every path in this file swallows
 * its own errors. A failed write loses one log entry; it must never lose
 * the user's work.
 */

// ── STORAGE ───────────────────────────────────────────────────────
var LOG_KEY_PREFIX     = 'BPLOG_';
// Deliberately has no underscore: it must NOT start with LOG_KEY_PREFIX,
// or logEntries_ would read the counter as though it were an entry.
var LOG_SEQ_KEY        = 'BPLOGSEQ';
var LOG_SCHEMA_VERSION = 1;
// Entries kept before the oldest are pruned. 400 × ~250 bytes leaves
// generous headroom under the 500KB document property store.
var LOG_MAX_ENTRIES    = 400;
// Summaries are trimmed to keep any single entry well under the 9KB
// per-value limit even when a tool reports a long list of modules.
var LOG_SUMMARY_MAX    = 300;
// How close together two calls must be to count as one user action. See
// MERGING below.
var LOG_MERGE_WINDOW_MS = 5 * 60 * 1000;
var LOG_LOCK_WAIT_MS    = 5000;

// ── TOOL NAMES ────────────────────────────────────────────────────
// One home for these strings: the wrappers log them and the viewer
// filters on them, so a typo in either place would silently break the
// "document changes only" filter.
var LOG_TOOL_BLUEPRINT          = 'Add Activity Titles, Tools, Due Date Headers, & Times';
var LOG_TOOL_DIRECTIONS         = 'Deploy Activity Directions';
var LOG_TOOL_DIRECTIONS_MODEL   = 'Deploy Activity Directions (from model module)';
var LOG_TOOL_MODEL_MODULE       = 'Create Model Module';
var LOG_TOOL_MODULE_DATES       = 'Specialty Tool: Module Dates & Holidays';
var LOG_TOOL_DESIGN_MAP_TITLES  = 'Design Map → Dev Tab (titles)';
var LOG_TOOL_DESIGN_MAP_MODULES = 'Design Map → Dev Tab (modules)';
var LOG_TOOL_TIME_ESTIMATOR     = 'Time Estimator';

// Tools that only READ the document. These never create a revision, so
// version history has no entry at their timestamp — the viewer tells the
// user to look for the nearest version at or before it instead, and the
// "document changes only" filter hides them.
var LOG_READ_ONLY_TOOLS = [LOG_TOOL_TIME_ESTIMATOR];

// ── CURRENT USER ──────────────────────────────────────────────────
/**
 * The signed-in user's email, or '' when Google will not disclose it.
 * getActiveUser() returns an empty string for anyone outside the script
 * owner's Workspace domain, and can throw outright depending on how the
 * script was authorised — neither case is an error worth surfacing, so
 * both come back as '' and the viewer renders a dash.
 */
function logCurrentUser_() {
  try {
    var email = Session.getActiveUser().getEmail();
    return email ? email : '';
  } catch (e) {
    return '';
  }
}

// ── SUMMARY TEXT ──────────────────────────────────────────────────
/**
 * Trims a tool's own summary down to what belongs in a log entry.
 *
 * Every write tool opens its summary with a "✅ ... updated!" banner.
 * That line says nothing the entry's own outcome field does not already
 * say, and it would crowd out the informative lines in the table's
 * preview column, so it is dropped.
 */
function logCondense_(text) {
  if (text === null || text === undefined) return '';
  var lines = String(text).split('\n');
  if (lines.length && /^\s*✅/.test(lines[0])) lines.shift();

  var kept = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].replace(/\s+$/, '');
    if (line.trim() === '') continue;
    kept.push(line);
  }

  var out = kept.join('\n');
  if (out.length > LOG_SUMMARY_MAX) {
    out = out.substring(0, LOG_SUMMARY_MAX - 1) + '…';
  }
  return out;
}

/** An error's message, however it was thrown. */
function logErrorText_(e) {
  if (!e) return 'Unknown error.';
  if (e.message) return String(e.message);
  return String(e);
}

// ── READING ───────────────────────────────────────────────────────
/**
 * Every stored entry as {key, entry}, oldest first. Unparseable values
 * are skipped rather than thrown on: one corrupt property must not make
 * the whole log unreadable.
 */
function logEntries_(props) {
  var all  = props.getProperties();
  var keys = [];
  for (var k in all) {
    if (Object.prototype.hasOwnProperty.call(all, k) && k.indexOf(LOG_KEY_PREFIX) === 0) {
      keys.push(k);
    }
  }
  keys.sort();   // 13-digit epoch millis sort chronologically as text

  var out = [];
  for (var i = 0; i < keys.length; i++) {
    try {
      out.push({ key: keys[i], entry: JSON.parse(all[keys[i]]) });
    } catch (e) {
      // Skip and carry on.
    }
  }
  return out;
}

/** When a run finished, in epoch millis. */
function logEndMs_(entry) {
  var start = Date.parse(entry.ts);
  if (isNaN(start)) return 0;
  return start + (entry.ms || 0);
}

// ── MERGING ───────────────────────────────────────────────────────
// Three tools are driven by their sidebar in several server calls per
// user action: "Deploy Activity Directions" splits its modules into up
// to three groups, "Design Map → Dev Tab (modules)" runs a chunk at a
// time, and the Time Estimator gets re-run repeatedly while someone
// tunes the credits/weeks inputs. Logging each call would bury the
// entries that matter behind a dozen near-identical rows.
//
// So a merge-eligible call within LOG_MERGE_WINDOW_MS of a matching
// entry rewrites that entry instead of adding one: counters are summed,
// lists concatenated, and the ORIGINAL start timestamp is kept — because
// the version someone wants to restore is the one from before the first
// call, not before the last.
//
// Merging is opt-in per tool and deliberately so. Two runs of a
// single-call write tool five minutes apart are two separate changes to
// the document, and collapsing them would hide a restore point.

/**
 * Folds one call's data into the running entry's.
 *
 * Only the keys named in sumKeys are added together. Summing by type
 * instead would be wrong: the Time Estimator's data holds SETTINGS
 * (credits, weeks, modules per week), and a merge that added those would
 * turn two runs of a 3-credit course into a 6-credit one. Counters are
 * declared, never inferred.
 *
 * Arrays always concatenate — every array in use here is a list of
 * skipped or unmatched items, and a chunked run's lists are partial.
 */
function logMergeData_(prev, next, sumKeys) {
  if (!prev) return next || null;
  if (!next) return prev;
  sumKeys = sumKeys || [];

  var out = {};
  var key;
  for (key in prev) {
    if (Object.prototype.hasOwnProperty.call(prev, key)) out[key] = prev[key];
  }
  for (key in next) {
    if (!Object.prototype.hasOwnProperty.call(next, key)) continue;
    var a = out[key];
    var b = next[key];
    if (sumKeys.indexOf(key) !== -1 && typeof a === 'number' && typeof b === 'number') {
      out[key] = a + b;
    } else if (Object.prototype.toString.call(a) === '[object Array]' &&
               Object.prototype.toString.call(b) === '[object Array]') {
      out[key] = a.concat(b).slice(0, 50);
    } else {
      out[key] = b;
    }
  }
  return out;
}

/**
 * The key of an entry this call should be merged into, or null for a new
 * entry. Same tool, same user, same outcome, inside the window.
 */
function logMergeTarget_(props, toolName, user, outcome, nowMs) {
  var entries = logEntries_(props);
  if (entries.length === 0) return null;

  var last = entries[entries.length - 1];
  if (last.entry.tool !== toolName)   return null;
  if (last.entry.user !== user)       return null;
  if (last.entry.outcome !== outcome) return null;
  if (nowMs - logEndMs_(last.entry) > LOG_MERGE_WINDOW_MS) return null;

  return last;
}

// ── WRITING ───────────────────────────────────────────────────────
/**
 * Records one run.
 *
 * @param {string} toolName  one of the LOG_TOOL_* constants
 * @param {string} outcome   'ok' or 'error'
 * @param {string} summary   the tool's own summary, or an error message
 * @param {number} startMs   Date.now() captured before the run began
 * @param {Object} [opts]
 *   .data       {Object}   structured counters kept on the entry
 *   .merge      {boolean}  fold into a recent matching entry (see MERGING)
 *   .sumKeys    {string[]} data keys that are counters, to be added on merge
 *   .summaryFn  {Function} given merged .data, returns the merged summary
 */
function logRun_(toolName, outcome, summary, startMs, opts) {
  opts = opts || {};
  var lock = null;
  try {
    lock = LockService.getDocumentLock();
    // Skip the entry rather than delay the user's run. A lost log line is
    // a far smaller cost than a tool that appears to hang.
    if (!lock.tryLock(LOG_LOCK_WAIT_MS)) return;

    var props = PropertiesService.getDocumentProperties();
    var now   = Date.now();
    var user  = logCurrentUser_();

    var target = opts.merge ? logMergeTarget_(props, toolName, user, outcome, now) : null;

    if (target) {
      var merged    = logMergeData_(target.entry.data, opts.data, opts.sumKeys);
      var startedAt = Date.parse(target.entry.ts);
      var text      = (opts.summaryFn && merged) ? opts.summaryFn(merged) : summary;

      target.entry.data    = merged;
      target.entry.summary = logCondense_(text);
      target.entry.ms      = isNaN(startedAt) ? (target.entry.ms || 0) : (now - startedAt);
      props.setProperty(target.key, JSON.stringify(target.entry));
      return;
    }

    var entry = {
      v:       LOG_SCHEMA_VERSION,
      ts:      new Date(startMs || now).toISOString(),
      tool:    toolName,
      user:    user,
      outcome: outcome,
      summary: logCondense_(summary),
      ms:      startMs ? (now - startMs) : 0
    };
    if (opts.data) entry.data = opts.data;

    props.setProperty(logNewKey_(props, now), JSON.stringify(entry));
    logPrune_(props);

  } catch (e) {
    try { Logger.log('logRun_ failed (run itself unaffected): ' + logErrorText_(e)); } catch (ignored) {}
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (ignored2) {} }
  }
}

/**
 * A key that is unique AND correctly ordered even when two runs land in
 * the same millisecond. Both properties are load-bearing — see STORAGE
 * SHAPE at the top of this file.
 *
 * Safe to read-modify-write because every caller holds the document lock.
 */
function logNewKey_(props, nowMs) {
  var seq = Number(props.getProperty(LOG_SEQ_KEY) || 0) + 1;
  if (seq > 999999) seq = 1;           // wraps harmlessly: the millis differ
  props.setProperty(LOG_SEQ_KEY, String(seq));

  var ms = String(nowMs);
  while (ms.length  < 13) ms  = '0' + ms;
  var sq = String(seq);
  while (sq.length  <  6) sq  = '0' + sq;
  return LOG_KEY_PREFIX + ms + '_' + sq;
}

/** Drops the oldest entries once the log passes LOG_MAX_ENTRIES. */
function logPrune_(props) {
  var entries = logEntries_(props);
  var excess  = entries.length - LOG_MAX_ENTRIES;
  for (var i = 0; i < excess; i++) {
    props.deleteProperty(entries[i].key);
  }
}

// ── SIDEBAR-CALLABLE API ──────────────────────────────────────────
// NO TRAILING UNDERSCORE on anything in this section. Apps Script treats
// a trailing underscore as private: such functions cannot be reached by
// google.script.run and cannot be used as menu targets.

/**
 * The log, newest first, formatted for the viewer.
 *
 * Times are rendered server-side in the script's own time zone
 * (America/Denver, per appsscript.json) so that every viewer reads the
 * same clock as the document's version history.
 */
function getRunLog() {
  var props   = PropertiesService.getDocumentProperties();
  var entries = logEntries_(props);
  var tz      = Session.getScriptTimeZone();
  var out     = [];

  for (var i = entries.length - 1; i >= 0; i--) {
    var e    = entries[i].entry;
    var when = new Date(e.ts);
    out.push({
      ts:       e.ts,
      when:     Utilities.formatDate(when, tz, 'MMM d, yyyy') + ' · ' +
                Utilities.formatDate(when, tz, 'h:mm a'),
      tool:     e.tool,
      user:     e.user || '',
      outcome:  e.outcome,
      summary:  e.summary || '',
      readOnly: LOG_READ_ONLY_TOOLS.indexOf(e.tool) !== -1,
      data:     e.data || null
    });
  }
  return { entries: out, count: out.length };
}

/** Empties the log. The viewer confirms before calling this. */
function clearRunLog() {
  var props   = PropertiesService.getDocumentProperties();
  var entries = logEntries_(props);
  for (var i = 0; i < entries.length; i++) {
    props.deleteProperty(entries[i].key);
  }
  return entries.length;
}

/** Menu target: opens the log viewer. */
function showRunLog() {
  var html = HtmlService.createHtmlOutputFromFile('LogViewer')
    .setWidth(720)
    .setHeight(540);
  DocumentApp.getUi().showModalDialog(html, 'Blueprint Tools — Run Log');
}

// ── LOGGED ENTRY POINTS ───────────────────────────────────────────
// The sidebars call these. Each one wraps the untouched implementation,
// now named <name>Core_ in its own file, and records the run.
//
// They live here rather than beside their implementations so that the
// tool files carry a one-word rename and nothing else — the logging
// concern stays in one place, and a future change to how runs are
// recorded touches this file alone.
//
// The shape is always the same: capture the start time, run the original
// body untouched, log, return. On a throw, log the failure and RETHROW,
// so the sidebar still shows the user exactly what it shows today.
// Failures are logged because a run that died partway may still have
// modified the document — which is precisely when someone needs to
// revert.

/** Add Activity Titles, Tools, Due Date Headers, & Times. One call per run. */
function processBlueprint(params) {
  var t0 = Date.now();
  try {
    var result = processBlueprintCore_(params);
    logRun_(LOG_TOOL_BLUEPRINT, 'ok', result, t0);
    return result;
  } catch (e) {
    logRun_(LOG_TOOL_BLUEPRINT, 'error', logErrorText_(e), t0);
    throw e;
  }
}

/**
 * Deploy Activity Directions, source-document variant.
 *
 * No sidebar in the suite currently calls this — Sidebar2 uses
 * applyDirectionsFromModel — but it is public API in a flat namespace, so
 * it is logged like anything else rather than left as a silent path.
 */
function applyDirections(params) {
  var t0 = Date.now();
  try {
    var result = applyDirectionsCore_(params);
    logRun_(LOG_TOOL_DIRECTIONS, 'ok', result, t0);
    return result;
  } catch (e) {
    logRun_(LOG_TOOL_DIRECTIONS, 'error', logErrorText_(e), t0);
    throw e;
  }
}

/**
 * Deploy Activity Directions from a model module.
 *
 * Sidebar2 splits the chosen modules into up to three groups and calls
 * this once per group, so one user action arrives here as several calls.
 * They merge into a single entry carrying the totals — see MERGING.
 */
function applyDirectionsFromModel(params) {
  var t0 = Date.now();
  try {
    var r = applyDirectionsFromModelCore_(params);
    logRun_(LOG_TOOL_DIRECTIONS_MODEL, 'ok', logDirectionsSummary_(r), t0, {
      merge:     true,
      sumKeys:   ['replaced', 'noMatch', 'targets'],
      data:      { replaced: r.replaced, noMatch: r.noMatch, targets: r.targets },
      summaryFn: logDirectionsSummary_
    });
    return r;
  } catch (e) {
    logRun_(LOG_TOOL_DIRECTIONS_MODEL, 'error', logErrorText_(e), t0, { merge: true });
    throw e;
  }
}

function logDirectionsSummary_(d) {
  return 'Directions deployed: ' + (d.replaced || 0) + '\n' +
         'Modules targeted: '    + (d.targets  || 0) + '\n' +
         'No match in model: '   + (d.noMatch  || 0);
}

/** Create Model Module. One call per run. */
function applyDirectionsToModule5(params) {
  var t0 = Date.now();
  try {
    var result = applyDirectionsToModule5Core_(params);
    logRun_(LOG_TOOL_MODEL_MODULE, 'ok', result, t0);
    return result;
  } catch (e) {
    logRun_(LOG_TOOL_MODEL_MODULE, 'error', logErrorText_(e), t0);
    throw e;
  }
}

/** Specialty Tool: Add Module Dates &/or Holiday Modules. One call per run. */
function applyModuleDates8(params) {
  var t0 = Date.now();
  try {
    var result = applyModuleDates8Core_(params);
    logRun_(LOG_TOOL_MODULE_DATES, 'ok', result, t0);
    return result;
  } catch (e) {
    logRun_(LOG_TOOL_MODULE_DATES, 'error', logErrorText_(e), t0);
    throw e;
  }
}

/** Design Map → Dev Tab, step 1: module titles. One call per run. */
function applyDesignMapTitles9(params) {
  var t0 = Date.now();
  try {
    var r = applyDesignMapTitles9Core_(params);
    var skipped = (r && r.skipped) ? r.skipped.length : 0;
    logRun_(LOG_TOOL_DESIGN_MAP_TITLES, 'ok',
            'Module titles written: ' + ((r && r.written) || 0) + '\n' +
            'Left as-is (already titled): ' + skipped, t0);
    return r;
  } catch (e) {
    logRun_(LOG_TOOL_DESIGN_MAP_TITLES, 'error', logErrorText_(e), t0);
    throw e;
  }
}

/**
 * Design Map → Dev Tab, step 2: objectives and notes.
 *
 * Sidebar9 writes a chunk of modules per call, so like the directions
 * deployer above, one user action arrives as several calls and merges
 * into one entry.
 */
function applyDesignMapModules9(params) {
  var t0 = Date.now();
  try {
    var r = applyDesignMapModules9Core_(params);
    var data = {
      objectives:        r.objectives || 0,
      notes:             r.notes || 0,
      objectivesSkipped: r.objectivesSkipped || [],
      notesSkipped:      r.notesSkipped || [],
      unmatched:         r.unmatched || []
    };
    logRun_(LOG_TOOL_DESIGN_MAP_MODULES, 'ok', logDesignMapSummary_(data), t0, {
      merge:     true,
      sumKeys:   ['objectives', 'notes'],
      data:      data,
      summaryFn: logDesignMapSummary_
    });
    return r;
  } catch (e) {
    logRun_(LOG_TOOL_DESIGN_MAP_MODULES, 'error', logErrorText_(e), t0, { merge: true });
    throw e;
  }
}

function logDesignMapSummary_(d) {
  var lines = [
    'Module overviews written: ' + (d.objectives || 0),
    'Note blocks written: '      + (d.notes || 0)
  ];
  var skipped = (d.objectivesSkipped ? d.objectivesSkipped.length : 0) +
                (d.notesSkipped      ? d.notesSkipped.length      : 0);
  if (skipped > 0) lines.push('Skipped (already had content): ' + skipped);
  if (d.unmatched && d.unmatched.length > 0) {
    lines.push('Unmatched Design Map rows: ' + d.unmatched.length);
  }
  return lines.join('\n');
}

/**
 * Time Estimator. Read-only, so it never creates a revision.
 *
 * TWO THINGS HERE ARE DELIBERATE AND EASY TO GET WRONG:
 *
 * 1. This function does NOT throw. It catches internally and returns
 *    {error: '...'} — see its own try/catch, and its early returns for
 *    bad inputs or a missing Development tab. A try/catch alone would
 *    record every failed run as a success, so the RESULT is inspected.
 *
 * 2. Only successful readings are logged. A validation bounce ("please
 *    enter valid values") records nothing about the document and has no
 *    forensic value. This differs on purpose from the write tools above,
 *    where failures ARE logged because a failed write may have left the
 *    document half-modified.
 *
 * The per-module time results are not stored. They are recoverable from
 * the document itself at that point in version history. The three
 * sidebar inputs ARE stored, because they are typed into the sidebar and
 * written nowhere in the document — without them a later reader can see
 * that Module 6 totalled 15 hours but cannot tell whether that was over
 * or under target, since the target derives entirely from these numbers.
 */
function runTimeEstimator(credits, weeks, modulesPerWeek) {
  var t0     = Date.now();
  var result = runTimeEstimatorCore_(credits, weeks, modulesPerWeek);

  if (result && !result.error) {
    logRun_(LOG_TOOL_TIME_ESTIMATOR, 'ok', LOG_READ_ONLY_SUMMARY, t0, {
      merge:   true,
      sumKeys: [],   // these are settings, not counters — never add them up
      data: {
        credits:        Number(credits),
        weeks:          Number(weeks),
        modulesPerWeek: Number(modulesPerWeek)
      },
      summaryFn: function () { return LOG_READ_ONLY_SUMMARY; }
    });
  }
  return result;
}

var LOG_READ_ONLY_SUMMARY = 'Read only — no changes to the document.';
