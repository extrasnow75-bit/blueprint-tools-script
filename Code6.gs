// ============================================================
// Blueprint Tools — Code6.gs
// Deploy Activity Directions (AI): adapts model-module directions
// for each target module using Gemini AI.
//
// Reuses from the shared GAS namespace:
//   Code4.gs — parseCourseDesignMap, extractTextFromElements4,
//               insertFormattedText4, callGemini4_, validateGeminiKey4,
//               findPlaceholderInModule4, GEMINI_PRIMARY_4, INDENT_4
//   Code2.gs — readModuleContent_, findMatchingModelContent_,
//               getDevelopmentTabBody, collectTabs, collectAllModuleActivities,
//               stripActivityHeading, findDirectionsPlaceholder
//   Code.gs  — FONT, RED, BLACK
// ============================================================

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

      if (!findDirectionsPlaceholder(devBody, para)) continue;

      var at        = stripActivityHeading(para.getText());
      var modelText = findMatchingModelContent_(modelTextByActivity, at);
      if (modelText) matches.push({ actTitle: at, modelText: modelText });
    }

    activitiesByTarget[tgt] = matches;
  }

  return {
    moduleContextByTarget: moduleContextByTarget,
    activitiesByTarget:    activitiesByTarget
  };
}

// ── PER-ACTIVITY ADAPTATION ────────────────────────────────────

/**
 * Adapts model directions for one activity in one target module and
 * inserts the AI-generated text into the blueprint document.
 *
 * Called sequentially from the sidebar — once per activity per module.
 *
 * @param {Object} params
 *   .apiKey            {string}
 *   .modelModuleTitle  {string}
 *   .targetModuleTitle {string}
 *   .activityTitle     {string}   stripped title
 *   .modelText         {string}   plain-text model directions
 *   .moduleContext     {Object|null}  CDM data for target module
 * @returns {{ success: boolean, skipped?: boolean, error?: string }}
 */
function adaptAndDeployActivity6(params) {
  try {
    var doc     = DocumentApp.getActiveDocument();
    var devBody = getDevelopmentTabBody(doc);
    if (!devBody) throw new Error('Could not find the Development tab.');

    var placeholder = findPlaceholderInModule4(
      devBody, params.targetModuleTitle, params.activityTitle
    );
    if (!placeholder) {
      // Slot already filled or title mismatch — skip.
      return { success: true, skipped: true };
    }

    var indent    = placeholder.getIndentStart() || INDENT_4;
    var prompt    = buildAdaptPrompt6_(params);
    var aiText    = callGemini4_(params.apiKey, prompt, GEMINI_PRIMARY_4);
    var insertIdx = devBody.getChildIndex(placeholder);
    placeholder.removeFromParent();
    insertFormattedText4(devBody, insertIdx, aiText, indent);

    return { success: true };
  } catch (e) {
    Logger.log('adaptAndDeployActivity6: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── PROMPT BUILDER ─────────────────────────────────────────────

function buildAdaptPrompt6_(params) {
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

  return (
    'You are an instructional designer adapting student-facing activity directions ' +
    'for a college online course.\n\n' +
    'The directions below were written for "' +
    (params.modelModuleTitle || 'the model module') +
    '". Adapt them for "' + params.targetModuleTitle + '". ' +
    'Keep the structure, tone, and formatting identical — only update content ' +
    'that should genuinely differ between modules, such as module-specific ' +
    'activity numbers or references to specific readings listed in the context.\n\n' +
    contextBlock +
    'ACTIVITY: ' + params.activityTitle + '\n\n' +
    'MODEL DIRECTIONS (adapt these):\n' + params.modelText + '\n\n' +
    'FORMATTING RULES — follow exactly:\n' +
    '- For section headings, write the heading text followed by (H2) or (H3) on the ' +
    'same line. Example: "Overview (H2)". Do NOT use # markdown syntax.\n' +
    '- For bullet lists, begin each item with "- " (dash then space).\n' +
    '- For inline bold text, wrap with **double asterisks**.\n' +
    '- Separate paragraphs with a blank line.\n' +
    '- Do not add a title heading — begin directly with content.\n\n' +
    'Write the adapted directions now.'
  );
}
