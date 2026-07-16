// ============================================================
// Blueprint Tools — Code6.gs
// Deploy Activity Directions (AI): adapts model-module directions
// for each target module using Gemini AI.
//
// Reuses from the shared GAS namespace:
//   Code4.gs — parseCourseDesignMap, extractTextFromElements4,
//               insertFormattedText4, callGemini4_, validateGeminiKey4,
//               findPlaceholderInTool4, INDENT_4
//               (deploy model is GEMINI_DEPLOY_6, defined below)
//   Code2.gs — readModuleContent_, findMatchingModelContent_,
//               getDevelopmentTabBody, collectTabs, collectAllModuleActivities,
//               stripActivityHeading, findDirectionsPlaceholder
//   Code.gs  — FONT, RED, BLACK
// ============================================================

// ── CONSTANTS ─────────────────────────────────────────────────
// Adaptation model. Lite has the highest free-tier request throughput, which
// matters because this tool makes one Gemini call per activity across many
// modules — the full flash (thinking) model saturates the free-tier rate limit
// quickly and spends tokens "thinking" about what is really a constrained
// adaptation task. Lite also matches validateGeminiKey4 (which validates against
// the lite model), so the key check now tests the same model deployment uses.
var GEMINI_DEPLOY_6 = 'gemini-2.5-flash-lite';

// ── SIDEBAR OPENER ────────────────────────────────────────────

function showDeployAiSidebar6() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar6')
    .setTitle('Deploy Activity Directions (AI)')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}

// ── SIDEBAR INIT DATA ─────────────────────────────────────────

/**
 * Returns the module list to populate the sidebar dropdowns.
 * @returns {{ moduleList: string[] }}
 */
function getDeployAiSidebarData6() {
  var doc  = DocumentApp.getActiveDocument();
  var body = getDevelopmentTabBody(doc);
  if (!body) return { moduleList: [] };
  return { moduleList: collectAllModuleActivities(body).moduleList };
}

/**
 * Lists the activity titles found in one model module, so the sidebar can offer
 * a per-activity picker (deploy only the chosen activities across targets).
 * Returns titles in document order. readModuleContent_ already skips the due-by
 * and Canvas-header markers, so these are clean activity titles only.
 *
 * @param {string} modelModuleTitle
 * @returns {{ activities: string[] }}
 */
function getModelModuleActivities6(modelModuleTitle) {
  var doc     = DocumentApp.getActiveDocument();
  var devBody = getDevelopmentTabBody(doc);
  if (!devBody) return { activities: [] };
  var content = readModuleContent_(devBody, modelModuleTitle);
  return { activities: Object.keys(content) };
}

// ── SESSION INITIALIZATION ─────────────────────────────────────

/**
 * Validates the API key, reads the model module as plain text, collects
 * CDM context for each target module, and pre-computes the per-target
 * activity work list (only activities that have a placeholder AND a
 * matching model direction).
 *
 * @param {Object} params
 *   .apiKey             {string}
 *   .modelModuleTitle   {string}
 *   .targetModuleTitles {string[]}
 * @returns {{
 *   keyError?:             string,
 *   moduleContextByTarget: Object,   // targetTitle → CDM context or null
 *   activitiesByTarget:    Object    // targetTitle → [{actTitle, modelText}]
 * }}
 */
function initDeployAiSession6(params) {
  var apiKey             = (params.apiKey || '').trim();
  var modelModuleTitle   = params.modelModuleTitle;
  var targetModuleTitles = params.targetModuleTitles;

  // Validate API key first.
  var keyResult = validateGeminiKey4(apiKey);
  if (!keyResult.valid) return { keyError: keyResult.error || 'Invalid API key.' };

  var doc  = DocumentApp.getActiveDocument();
  var tabs = collectTabs(doc);

  // Read model module content as plain text.
  var devBody = getDevelopmentTabBody(doc);
  if (!devBody) throw new Error('Could not find a "Development" tab in this document.');

  var modelElements = readModuleContent_(devBody, modelModuleTitle);
  var modelTextByActivity = {};
  for (var actTitle in modelElements) {
    modelTextByActivity[actTitle] = extractTextFromElements4(modelElements[actTitle]);
  }

  // Restrict to the activities the user checked in the picker. Each target
  // activity is only included when it matches a model activity (below), so
  // shrinking this map is all that's needed to deploy just the chosen
  // activities. Empty/undefined means "all" (backwards compatible).
  if (params.selectedActivities && params.selectedActivities.length) {
    var keep = {};
    for (var s = 0; s < params.selectedActivities.length; s++) {
      var sel = params.selectedActivities[s];
      if (modelTextByActivity[sel] !== undefined) keep[sel] = modelTextByActivity[sel];
    }
    modelTextByActivity = keep;
  }

  if (Object.keys(modelTextByActivity).length === 0) {
    throw new Error(
      'No directions found in "' + modelModuleTitle + '". ' +
      'Run "Create Model Module" on it first, then try again.'
    );
  }

  // Parse Course Design Map for all target modules.
  var cdmModules = [];
  for (var i = 0; i < tabs.length; i++) {
    if (/\bdesign\b/i.test(tabs[i].title)) {
      cdmModules = parseCourseDesignMap(tabs[i].body);
      break;
    }
  }

  var moduleContextByTarget = {};
  for (var m = 0; m < targetModuleTitles.length; m++) {
    var target    = targetModuleTitles[m];
    var targetNum = (target.match(/\d+/) || [''])[0];
    var ctx = null;
    for (var c = 0; c < cdmModules.length; c++) {
      var labelNum = (cdmModules[c].moduleLabel.match(/\d+/) || [''])[0];
      if (labelNum === targetNum) { ctx = cdmModules[c]; break; }
    }
    moduleContextByTarget[target] = ctx;
  }

  // Fetch real titles for every target module's reading/video links in one
  // parallel batch (de-duplicated by URL inside enrichMediaTitles_).
  var ctxToEnrich = [];
  for (var e = 0; e < targetModuleTitles.length; e++) {
    var ec = moduleContextByTarget[targetModuleTitles[e]];
    if (ec) ctxToEnrich.push(ec);
  }
  enrichMediaTitles_(ctxToEnrich);

  // Pre-compute which activities to adapt in each target module.
  // An activity is included if it has a placeholder AND a matching model direction.
  // The resolved model text is stored alongside the activity title so the client
  // can pass it back per-call without a second lookup.
  var H2 = DocumentApp.ParagraphHeading.HEADING2;
  var H4 = DocumentApp.ParagraphHeading.HEADING4;
  var activitiesByTarget = {};

  for (var t = 0; t < targetModuleTitles.length; t++) {
    var tgt     = targetModuleTitles[t];
    var escaped = tgt.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    var modRe   = new RegExp('^' + escaped + '[:\\s]', 'i');
    var inMod   = false;
    var matches = [];
    var n       = devBody.getNumChildren();

    for (var j = 0; j < n; j++) {
      var child = devBody.getChild(j);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      var para    = child.asParagraph();
      var heading = para.getHeading();
      var txt     = para.getText().trim();

      if (heading === H2) {
        if (modRe.test(txt)) { inMod = true; }
        else if (inMod)      { break; }
        continue;
      }
      if (!inMod || heading !== H4) continue;

      if (!findDirectionsPlaceholder(devBody, para, j)) continue;

      var at        = stripActivityHeading(para.getText());
      var modelText = findMatchingModelContent_(modelTextByActivity, at);
      if (modelText) {
        matches.push({
          actTitle:  at,
          modelText: modelText,
          toolType:  getToolTypeForSlot(devBody, para, j)  // for the assignment help section
        });
      }
    }

    activitiesByTarget[tgt] = matches;
  }

  return {
    moduleContextByTarget: moduleContextByTarget,
    activitiesByTarget:    activitiesByTarget
  };
}

// ── PER-MODULE ADAPTATION (BATCHED) ────────────────────────────

/**
 * Adapts every activity in ONE target module with a SINGLE Gemini call, then
 * inserts each activity's adapted directions into its placeholder.
 *
 * Batching (one call per module instead of one per activity) cuts the number of
 * API requests ~10x, which is the difference between constantly tripping the
 * free-tier rate limit and finishing in a couple of minutes.
 *
 * @param {Object} params
 *   .apiKey            {string}
 *   .modelModuleTitle  {string}
 *   .targetModuleTitle {string}
 *   .activities        {Array<{activityTitle: string, modelText: string}>}
 *   .moduleContext     {Object|null}  CDM data for the target module
 * @returns {{
 *   success: boolean,
 *   results?: Array<{activityTitle: string, success: boolean, skipped?: boolean, error?: string}>,
 *   error?: string
 * }}
 */
function adaptAndDeployModule6(params) {
  var activities = params.activities || [];
  if (activities.length === 0) return { success: true, results: [] };

  try {
    var prompt = buildModuleAdaptPrompt6_(params);
    // One call now returns directions for the whole module, so give it a
    // generous output budget (default caps can truncate a multi-activity reply).
    var aiText = callGemini4_(params.apiKey, prompt, GEMINI_DEPLOY_6, { maxOutputTokens: 32768 });
    var blocks = splitModuleResponse6_(aiText, activities.length);

    // Resolve the target module's design-map link tokens (VIDEO1, READING1, …).
    var media = buildMediaTokens4_(params.moduleContext ? params.moduleContext.mediaLinks : null);

    // Course code for the assignment help section (read once per module).
    var courseCode = extractCourseCode_();

    var doc     = DocumentApp.getActiveDocument();
    var devBody = getDevelopmentTabBody(doc);
    if (!devBody) throw new Error('Could not find the Development tab.');

    var results = [];
    for (var i = 0; i < activities.length; i++) {
      var act     = activities[i];
      var aiBlock = blocks[i];

      // The model omitted this activity — report it, but leave the placeholder
      // in place so nothing is destroyed.
      if (!aiBlock || !aiBlock.trim()) {
        results.push({ activityTitle: act.activityTitle, success: false,
                       error: 'No adapted text returned for this activity.' });
        continue;
      }

      var placeholder = findPlaceholderInTool4(devBody, params.targetModuleTitle, act.activityTitle);
      if (!placeholder) {
        // Slot already filled or title mismatch — skip (not an error).
        results.push({ activityTitle: act.activityTitle, success: true, skipped: true });
        continue;
      }

      // Assignments get the standardized "Where to go for help" section appended
      // deterministically (never AI-authored) so wording and links are exact.
      // Strip any code fence first so the appended block survives insertion.
      if (isAssignmentToolType4_(act.toolType)) {
        aiBlock = stripCodeFence4_(aiBlock) +
                  assignmentHelpBlock4_(courseCode, act.activityTitle);
      }

      var indent    = placeholder.getIndentStart() || INDENT_4;
      var insertIdx = devBody.getChildIndex(placeholder);
      placeholder.removeFromParent();
      insertFormattedText4(devBody, insertIdx, aiBlock, indent, media.linkMap);
      results.push({ activityTitle: act.activityTitle, success: true });
    }

    return { success: true, results: results };
  } catch (e) {
    // Whole-module failure (e.g. a rate limit that survived callGemini4_'s
    // retries). Nothing was inserted, so the client can safely retry the module.
    Logger.log('adaptAndDeployModule6 [' + params.targetModuleTitle + ']: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── RESPONSE SPLITTER ──────────────────────────────────────────

/**
 * Splits a batched Gemini reply into per-activity blocks using the
 * "@@@ACTIVITY N@@@" delimiter the prompt requires. Mapping is by the 1-based
 * number N (robust to the model reordering activities). Any block the model
 * omits is left undefined so the caller flags it rather than mis-assigning it.
 *
 * @param {string} aiText
 * @param {number} expectedCount
 * @returns {Array<string|undefined>} blocks[i] = directions for activity i (0-based)
 */
function splitModuleResponse6_(aiText, expectedCount) {
  aiText = String(aiText || '');

  // Strip a markdown code fence ONLY when it wraps the entire reply (anchored).
  // A non-anchored match would grab just the first ```…``` block and discard
  // everything after it — fatal if the model fences each activity separately.
  // Per-activity fences (if any) are handled downstream by insertFormattedText4.
  var whole = aiText.trim().match(/^```(?:markdown|text)?\n?([\s\S]*?)\n?```$/);
  if (whole) aiText = whole[1];

  var re      = /@@@ACTIVITY\s+(\d+)@@@/g;
  var blocks  = new Array(expectedCount);
  var matches = [];
  var m;
  while ((m = re.exec(aiText)) !== null) {
    matches.push({ num: parseInt(m[1], 10), contentStart: re.lastIndex, markerStart: m.index });
  }

  // No delimiters at all — only safe to use the whole reply if a single
  // activity was requested; otherwise we cannot split reliably.
  if (matches.length === 0) {
    if (expectedCount === 1) blocks[0] = aiText.trim();
    return blocks;
  }

  for (var i = 0; i < matches.length; i++) {
    var end  = (i + 1 < matches.length) ? matches[i + 1].markerStart : aiText.length;
    var text = aiText.slice(matches[i].contentStart, end).trim();
    var idx  = matches[i].num - 1; // 1-based → 0-based
    if (idx >= 0 && idx < expectedCount) blocks[idx] = text;
  }
  return blocks;
}

// ── PROMPT BUILDER ─────────────────────────────────────────────

function buildModuleAdaptPrompt6_(params) {
  var ctx = params.moduleContext;

  var contextLines = [];
  if (ctx) {
    if (ctx.moduleTitle) contextLines.push('Module title: ' + ctx.moduleTitle);
    if (ctx.clos)        contextLines.push('Learning Outcomes: ' + ctx.clos);
    if (ctx.readings)    contextLines.push('Readings/content: ' + ctx.readings);
    if (ctx.activityDescriptions && ctx.activityDescriptions.length > 0) {
      contextLines.push('Activity notes: ' + ctx.activityDescriptions.join('; '));
    }
  }
  var contextBlock = contextLines.length > 0
    ? 'TARGET MODULE CONTEXT:\n' + contextLines.join('\n') + '\n\n'
    : '';

  // Available-links block for THIS target module — reference-by-TOKEN.
  var media = buildMediaTokens4_(ctx ? ctx.mediaLinks : null);
  var mediaBlock = media.hasLinks
    ? 'AVAILABLE LINKS for "' + params.targetModuleTitle + '" — reference each by its ' +
      'TOKEN, never by retyping the URL:\n' + media.promptLines.join('\n') + '\n\n'
    : '';
  var linkRule = media.hasLinks
    ? '- When an activity references a reading or video, embed the link in its TITLE ' +
      'using markdown syntax with the TOKEN as the target, e.g. "[Read Chapter 3](READING1)". ' +
      'When AVAILABLE LINKS supplies a title in quotes, use that exact title as the link text. ' +
      'Use only tokens from AVAILABLE LINKS; never write the raw URL or a "[Link here]" ' +
      'placeholder.\n'
    : '- Do NOT invent link placeholders such as "[Link to video here]".\n';

  var activities = params.activities || [];
  var actBlocks  = [];
  for (var i = 0; i < activities.length; i++) {
    // Neutralize any literal "@@@" in doc-sourced text so it can't spoof the
    // "@@@ACTIVITY N@@@" delimiter we split the reply on (collapsing to a single
    // "@" cannot match the 3-@ marker).
    var safeTitle = String(activities[i].activityTitle || '').replace(/@{2,}/g, '@');
    var safeModel = String(activities[i].modelText || '').replace(/@{2,}/g, '@');
    // Assignments get a standardized help section appended in code. The model
    // directions may already contain one — tell the AI to omit it so it isn't
    // duplicated (and don't let it author its own).
    var assignNote = isAssignmentToolType4_(activities[i].toolType)
      ? '\n(NOTE: This is an assignment. OMIT any "Where to go for help", technical ' +
        'support, or Help Desk section from your adaptation — a standardized one is ' +
        'added automatically.)'
      : '';
    // Keep readings-only / videos-only activities from picking up the other type.
    var actKind   = classifyActivityKind_(activities[i].activityTitle);
    var scopeNote = (actKind === 'reading')
      ? '\n(This activity covers READINGS ONLY — do not add a videos/multimedia section.)'
      : (actKind === 'video')
        ? '\n(This activity covers VIDEOS ONLY — do not add a readings section.)'
        : '';
    actBlocks.push(
      'ACTIVITY ' + (i + 1) + ': ' + safeTitle + assignNote + scopeNote + '\n' +
      'MODEL DIRECTIONS (adapt these):\n' + safeModel
    );
  }

  return (
    'You are an instructional designer adapting student-facing activity directions ' +
    'for a college online course.\n\n' +
    'The directions below were written for "' +
    (params.modelModuleTitle || 'the model module') +
    '". Adapt EACH activity for "' + params.targetModuleTitle + '". ' +
    'Keep the structure, tone, and formatting identical — only update content ' +
    'that should genuinely differ between modules, such as module-specific ' +
    'activity numbers or references to specific readings listed in the context.\n\n' +
    contextBlock +
    mediaBlock +
    'There are ' + activities.length + ' activities to adapt:\n\n' +
    actBlocks.join('\n\n---\n\n') + '\n\n' +
    'FORMATTING RULES — follow exactly:\n' +
    '- Return the adapted directions for EVERY activity, in order, each preceded ' +
    'by a delimiter line EXACTLY of the form "@@@ACTIVITY N@@@" where N is the ' +
    'activity number shown above (e.g. "@@@ACTIVITY 1@@@"). Put the delimiter on ' +
    'its own line, then that activity\'s directions beneath it.\n' +
    '- For section headings, write the heading text followed by (H2) or (H3) on the ' +
    'same line. Example: "Overview (H2)". Do NOT use # markdown syntax.\n' +
    '- For bullet lists, begin each item with "- " (dash then space).\n' +
    '- For inline bold text, wrap with **double asterisks**.\n' +
    '- Separate paragraphs with a blank line.\n' +
    '- Do not add a title heading for an activity — after its delimiter, begin ' +
    'directly with the directions content.\n' +
    linkRule +
    '- If you refer to the campus help desk, call it exactly "Boise State Help Desk".\n' +
    '- Do NOT include any "Due by … Mountain Time" date header, nor any Canvas ' +
    'header annotation such as "Unpublished text header in Canvas" — those are ' +
    'generated separately and must never appear in the directions.\n\n' +
    'Write the adapted directions for all ' + activities.length + ' activities now.'
  );
}
