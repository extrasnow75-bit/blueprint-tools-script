/**
 * ================================================================
 * BLUEPRINT TOOLS 3  |  Google Apps Script  |  Time Estimator
 * ================================================================
 * Scans the Development tab of a Blueprint document, tallies the
 * time estimates for each numbered module (Module 1 / Week 1, etc.),
 * and compares totals to credit-hour guidelines.
 *
 * Guideline formula (from Clock Hours table):
 *   Min hrs per course = credits × 37.5
 *   Max hrs per course = credits × 45
 *   Target per module  = (min or max) ÷ weeks ÷ modulesPerWeek
 *                        rounded UP to the nearest ¼ hour
 * ================================================================
 */


// ── SIDEBAR ───────────────────────────────────────────────────────

function showTimeEstimatorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar3')
    .setTitle('🎓 Time Estimator')
    .setWidth(340);
  DocumentApp.getUi().showSidebar(html);
}


// ── TAB UTILITIES ─────────────────────────────────────────────────

/**
 * Recursively collects all tabs (and child tabs) of a document.
 * Returns an array of { title, body } objects.
 * Uses a unique name (collectTabs3) to avoid conflicts if this file
 * is ever combined with Blueprint Tools 1 or 2 in a single project.
 */
function collectTabs3(doc) {
  const result = [];
  function walk(tab) {
    result.push({ title: tab.getTitle(), body: tab.asDocumentTab().getBody() });
    tab.getChildTabs().forEach(walk);
  }
  doc.getTabs().forEach(walk);
  return result;
}

/** Returns the Body of the Development tab, or null if not found. */
function getDevelopmentTabBody3(doc) {
  const tab = collectTabs3(doc).find(t => /\bdevelopment\b/i.test(t.title));
  return tab ? tab.body : null;
}


// ── TIME PARSING ──────────────────────────────────────────────────

/**
 * Parses a time string into total minutes.
 *
 * Handles: "30 min", "1 hr", "1.5 hr", "1.5 hrs", "1 hour",
 *          "2 hours", "1 hr 30 min", "1 hr 30 mins",
 *          "1 hour and 30 minutes", "1hr 30min", etc.
 *
 * Returns:
 *   integer ≥ 0  — parsed minutes
 *   -1           — "(TBD)" detected
 *   null         — no recognisable time expression found
 */
function parseTimeToMinutes3(text) {
  if (!text) return null;
  text = text.trim();

  if (/\btbd\b/i.test(text)) return -1;

  let total = 0;
  let found = false;

  // Hours: "1 hr", "1.5 hrs", "1 hour", "2 hours", "1.5hr"
  const hMatch = text.match(/(\d+(?:\.\d+)?)\s*(?:hrs?|hours?)/i);
  if (hMatch) { total += parseFloat(hMatch[1]) * 60; found = true; }

  // Minutes: "30 min", "30 mins", "30 minutes"
  const mMatch = text.match(/(\d+)\s*(?:mins?|minutes?)/i);
  if (mMatch) { total += parseInt(mMatch[1], 10); found = true; }

  return found ? Math.round(total) : null;
}

/**
 * Extracts a time value from the parenthesised suffix at the END of an
 * activity heading text, e.g.:
 *   "1.01 Discussion (1 hr)"         → 60
 *   "1.02 Writing Assignment (TBD)"  → -1
 *   "1.03 Activity Title"            → null
 */
function timeFromHeading3(text) {
  const m = text.match(/\(([^)]+)\)\s*$/);
  return m ? parseTimeToMinutes3(m[1]) : null;
}

/**
 * Extracts a time value from an "Estimated time: …" paragraph, e.g.:
 *   "Estimated time: 30 min"   → 30
 *   "Estimated time: (TBD)"    → -1
 *   "Estimated time:"          → null  (label only, no value entered)
 */
function timeFromEstLine3(text) {
  const m = text.match(/estimated\s*time[:\s]+(.+)/i);
  return m ? parseTimeToMinutes3(m[1].trim()) : null;
}


// ── FORMATTING HELPERS ────────────────────────────────────────────

function fmtMinutes3(mins) {
  const h = Math.floor(mins / 60);
  const m = Math.round(mins % 60);
  if (h === 0) return `${m} min`;
  if (m === 0) return `${h} hr${h !== 1 ? 's' : ''}`;
  return `${h} hr${h !== 1 ? 's' : ''} ${m} min`;
}

function fmtHours3(hrs) {
  return fmtMinutes3(Math.round(hrs * 60));
}

/** Round UP to the nearest ¼ hour (matches the Clock Hours guideline table). */
function roundUpQtr3(hrs) {
  return Math.ceil(hrs * 4) / 4;
}


// ── MAIN ──────────────────────────────────────────────────────────

/**
 * Entry point called by the sidebar Run button.
 *
 * @param {number} credits        — credit hours (e.g. 3)
 * @param {number} weeks          — course length in weeks (e.g. 7)
 * @param {number} modulesPerWeek — modules a student completes per week (e.g. 1)
 *
 * @returns {Object} results payload for the sidebar, with keys:
 *   modules      [{name, total, status}]
 *   avg          string
 *   avgStatus    {type, diff?, target?}
 *   minTarget    string
 *   maxTarget    string
 *   missing      [{module, title, reason}]
 *   error        string  (only on failure)
 */
function runTimeEstimator(credits, weeks, modulesPerWeek) {
  try {
    credits        = Number(credits);
    weeks          = Number(weeks);
    modulesPerWeek = Number(modulesPerWeek);

    if (!credits || !weeks || !modulesPerWeek) {
      return { error: 'Please enter valid values for all fields.' };
    }

    const doc  = DocumentApp.getActiveDocument();
    const body = getDevelopmentTabBody3(doc);
    if (!body) {
      return { error: 'Could not find a tab named "Development" in this document.' };
    }

    const n = body.getNumChildren();

    // ── Scan the Development tab body ─────────────────────────────
    const modules       = [];   // [{ shortName, totalMins }]
    let   currentModule = null;
    const missing       = [];   // [{ module, title, reason }]

    for (let i = 0; i < n; i++) {
      const child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

      const para    = child.asParagraph();
      const heading = para.getHeading();
      const text    = para.getText().trim();
      if (!text) continue;

      // ── H2: module boundary ──────────────────────────────────────
      if (heading === DocumentApp.ParagraphHeading.HEADING2) {
        if (/^(module|week)\s+\d+/i.test(text)) {
          // Numbered module — extract "Module X" or "Week X" as the display name.
          const nameMatch = text.match(/^((?:module|week)\s+\d+)/i);
          const shortName = nameMatch ? nameMatch[1] : text.split(':')[0].trim();
          currentModule = { shortName, totalMins: 0 };
          modules.push(currentModule);
        } else {
          // Non-numbered heading (Course Resources, Spring Break, etc.) — skip.
          currentModule = null;
        }
        continue;
      }

      // Only count activities that fall inside a numbered module.
      if (!currentModule) continue;

      // ── H4: activity heading ─────────────────────────────────────
      if (heading === DocumentApp.ParagraphHeading.HEADING4) {

        // Priority 1: time estimate at the end of the heading itself.
        let timeResult = timeFromHeading3(text);

        // Priority 2: "Estimated time:" paragraph immediately following.
        // Per spec: if BOTH are present, the heading value wins (already captured above).
        if (timeResult === null) {
          for (let j = i + 1; j < Math.min(i + 5, n); j++) {
            const nx = body.getChild(j);
            if (nx.getType() !== DocumentApp.ElementType.PARAGRAPH) break;
            const nxPara = nx.asParagraph();
            if (nxPara.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) break;
            const nxText = nxPara.getText().trim();
            if (!nxText) continue;
            if (/estimated\s*time/i.test(nxText)) {
              timeResult = timeFromEstLine3(nxText);
            }
            break;   // only ever check the first non-blank normal paragraph
          }
        }

        // Tally or flag as missing.
        if (timeResult === null || timeResult === -1) {
          missing.push({
            module: currentModule.shortName,
            title:  text,
            reason: timeResult === -1 ? 'TBD' : 'not entered'
          });
          // Counts as 0 — do not add to totalMins.
        } else {
          currentModule.totalMins += timeResult;
        }
      }
    }

    if (modules.length === 0) {
      return {
        error: 'No numbered modules (Module 1, Week 1, etc.) were found in the Development tab.'
      };
    }

    // ── Recommended range per module ─────────────────────────────
    const minPerModHrs = roundUpQtr3((credits * 37.5) / weeks / modulesPerWeek);
    const maxPerModHrs = roundUpQtr3((credits * 45)   / weeks / modulesPerWeek);

    // ── Status helper ─────────────────────────────────────────────
    function statusFor(actualHrs) {
      const excess    = actualHrs - maxPerModHrs;
      const shortfall = minPerModHrs - actualHrs;
      if (excess    > 0.001) return { type: 'over',  diff: fmtHours3(excess),    target: fmtHours3(maxPerModHrs) };
      if (shortfall > 0.001) return { type: 'under', diff: fmtHours3(shortfall), target: fmtHours3(minPerModHrs) };
      return { type: 'ok' };
    }

    // ── Per-module results ────────────────────────────────────────
    const modResults = modules.map(m => ({
      name:   m.shortName,
      total:  fmtMinutes3(m.totalMins),
      status: statusFor(m.totalMins / 60)
    }));

    // ── Average across all counted modules ────────────────────────
    const totalMins = modules.reduce((s, m) => s + m.totalMins, 0);
    const avgMins   = totalMins / modules.length;

    return {
      modules:   modResults,
      avg:       fmtMinutes3(Math.round(avgMins)),
      avgStatus: statusFor(avgMins / 60),
      minTarget: fmtHours3(minPerModHrs),
      maxTarget: fmtHours3(maxPerModHrs),
      missing
    };

  } catch (e) {
    Logger.log('runTimeEstimator error: ' + e.message);
    return { error: e.message };
  }
}
