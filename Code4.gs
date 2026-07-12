/**
 * ================================================================
 * BLUEPRINT TOOLS  |  Code4.gs
 * Tool 4: Create Model Module — AI-generated activity directions
 * ================================================================
 *
 * Shared namespace: relies on constants and helpers defined in
 * Code.gs (FONT, RED, BLACK, _fmt, collectTabs, parseCoursePattern,
 * stripActivityHeading, getTemplateIndent) and Code2.gs
 * (DEFAULT_SOURCE_URL, DIRECTION_OPTIONS, getDevelopmentTabBody,
 * collectAllModuleActivities, getActivityPattern,
 * getActivityPatternForModule, findDirectionElements,
 * findDirectionsPlaceholder, extractDocId, getDocTitleFromUrl).
 *
 * NOTE FOR USER: Add the following lines to onOpen() in Code.gs
 * (inside the createMenu chain, after the Time Estimator item):
 *
 *   .addSeparator()
 *   .addItem('Create Model Module (AI)', 'showAiDirectionsSidebar4')
 * ================================================================
 */

// ── CONSTANTS ────────────────────────────────────────────────────────

var GEMINI_BASE_URL_4  = 'https://generativelanguage.googleapis.com/v1beta/models/';
var GEMINI_PRIMARY_4   = 'gemini-2.5-flash';
var GEMINI_FAST_4      = 'gemini-2.5-flash-lite';
var INDENT_4           = 36; // default activity-slot indent (points)

// ── SIDEBAR ──────────────────────────────────────────────────────────

function showAiDirectionsSidebar4() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar4')
    .setTitle('Create Model Module')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}

// ── SIDEBAR DATA ─────────────────────────────────────────────────────

/**
 * Returns initial data for the sidebar on load.
 * @returns {{ moduleList: string[], defaultTemplateUrl: string }}
 */
function getAiSidebarData4() {
  var doc  = DocumentApp.getActiveDocument();
  var body = getDevelopmentTabBody(doc);

  if (!body) {
    return { moduleList: [], defaultTemplateUrl: DEFAULT_SOURCE_URL };
  }

  var data = collectAllModuleActivities(body);
  return {
    moduleList:         data.moduleList,
    defaultTemplateUrl: DEFAULT_SOURCE_URL
  };
}

// ── COURSE DESIGN MAP PARSER ─────────────────────────────────────────

/**
 * Scans the Design tab body for the Course Design Map table and returns
 * per-module context objects.
 *
 * Expected table structure — one repeating block per module:
 *   Row 0: "Module X"  (first cell; second cell may be empty or merged)
 *   Row 1: "Title"                   | [text]
 *   Row 2: "CLOs, MLOs"              | [text]
 *   Row 3: "Reading / Video / Content"| [text]
 *   Row 4+: "Activity or Assessment"  | [text]  (1–3 rows)
 *
 * The table is identified by the presence of "clo" or "mlo" in any cell.
 *
 * @param {GoogleAppsScript.Document.Body} designBody
 * @returns {Array<{moduleLabel, moduleTitle, clos, readings, activityDescriptions[]}>}
 */
function parseCourseDesignMap(designBody) {
  var modules = [];
  var tables  = designBody.getTables();

  for (var t = 0; t < tables.length; t++) {
    var table   = tables[t];
    var numRows = table.getNumRows();
    if (numRows < 4) continue;

    // Cache of cell text keyed by row index. Rows fetched during the
    // detection scan below are reused by the parse loop instead of being
    // re-read from the document — each getText() is a server round-trip.
    var rowCache = {};

    // Identify as Course Design Map by presence of a "clo" or "mlo" label.
    // Scan stops advancing rows once found; the current row is finished so
    // its cells are fully cached.
    var isCDM     = false;
    var scanLimit = Math.min(numRows, 15);
    for (var ri = 0; ri < scanLimit && !isCDM; ri++) {
      var scanRow = table.getRow(ri);
      var scanN   = scanRow.getNumCells();
      var rowText = [];
      for (var ci = 0; ci < scanN; ci++) {
        var cellStr = scanRow.getCell(ci).getText();
        rowText.push(cellStr);
        var lc = cellStr.toLowerCase();
        if (lc.indexOf('clo') !== -1 || lc.indexOf('mlo') !== -1) isCDM = true;
      }
      rowCache[ri] = rowText;
    }
    if (!isCDM) continue;

    var currentModule = null;

    for (var r = 0; r < numRows; r++) {
      // Reuse cached row text when available; otherwise read only cells 0–1
      // (the columns used below) once from the document.
      var cells = rowCache[r];
      if (!cells) {
        var dataRow = table.getRow(r);
        var dataN   = dataRow.getNumCells();
        cells = [];
        for (var c = 0; c < dataN && c < 2; c++) cells.push(dataRow.getCell(c).getText());
      }
      var numCells = cells.length;
      if (numCells === 0) continue;

      var firstCell = cells[0].trim();

      // Module header row: "Module 1", "Module 2", etc.
      if (/^module\s+\d+/i.test(firstCell)) {
        var labelMatch = firstCell.match(/^(module\s+\d+)/i);
        currentModule = {
          moduleLabel:          labelMatch[1],
          moduleTitle:          '',
          clos:                 '',
          readings:             '',
          activityDescriptions: []
        };
        modules.push(currentModule);
        continue;
      }

      // Data rows need a value column
      if (!currentModule || numCells < 2) continue;

      var label = firstCell.toLowerCase();
      var value = cells[1].trim();
      if (!value) continue;

      if (label === 'title') {
        currentModule.moduleTitle = value;
      } else if (label.indexOf('clo') !== -1 || label.indexOf('mlo') !== -1) {
        currentModule.clos = value;
      } else if (label.indexOf('reading') !== -1 ||
                 label.indexOf('video')   !== -1 ||
                 label.indexOf('content') !== -1) {
        currentModule.readings = value;
      } else if (label.indexOf('activity')   !== -1 ||
                 label.indexOf('assessment') !== -1) {
        currentModule.activityDescriptions.push(value);
      }
    }

    if (modules.length > 0) break; // found the CDM table — stop searching
  }

  return modules;
}

// ── SESSION INITIALIZATION ────────────────────────────────────────────

/**
 * Called once when the user clicks Generate. Opens the template doc if
 * requested, fetches the default-choice direction text for each unique
 * tool type, and parses the Course Design Map for the target module.
 *
 * @param {string} moduleTitle      e.g. "Module 1"
 * @param {string} templateSource   "standard" | "custom" | "none"
 * @param {string} templateUrl      used when templateSource === "custom"
 * @returns {{
 *   moduleContext: Object|null,
 *   templateContentsByToolType: Object,
 *   templateAccessFailed: boolean,
 *   failedTemplateUrl: string|null
 * }}
 */
function initAiSession4(moduleTitle, templateSource, templateUrl) {
  var doc  = DocumentApp.getActiveDocument();
  var tabs = collectTabs(doc);

  // ── Parse Course Design Map ─────────────────────────────────────
  var moduleContext = null;
  for (var i = 0; i < tabs.length; i++) {
    if (/\bdesign\b/i.test(tabs[i].title)) {
      var cdmModules = parseCourseDesignMap(tabs[i].body);
      var targetNum  = (moduleTitle.match(/\d+/) || ['1'])[0];
      for (var j = 0; j < cdmModules.length; j++) {
        var labelNum = (cdmModules[j].moduleLabel.match(/\d+/) || [''])[0];
        if (labelNum === targetNum) {
          moduleContext = cdmModules[j];
          break;
        }
      }
      break;
    }
  }

  // ── Fetch template content (default choice per tool type) ───────
  var templateContentsByToolType = {};
  var templateAccessFailed       = false;
  var failedTemplateUrl          = null;
  var templateError              = null;

  if (templateSource !== 'none') {
    var resolvedUrl   = (templateSource === 'standard') ? DEFAULT_SOURCE_URL : templateUrl;
    failedTemplateUrl = resolvedUrl; // will be cleared on success

    try {
      var sourceDocId = extractDocId(resolvedUrl);
      var sourceDoc   = DocumentApp.openById(sourceDocId);
      var sourceTabs  = collectTabs(sourceDoc);

      var activities = getActivityPattern(moduleTitle);
      var seenTypes  = {};

      for (var k = 0; k < activities.length; k++) {
        var tt = activities[k].toolType;
        if (!tt || seenTypes[tt]) continue;
        seenTypes[tt] = true;

        var opts = DIRECTION_OPTIONS[tt];
        if (!opts || !opts.defaultChoice || opts.defaultChoice === '[Leave Blank]') continue;

        var elements = findDirectionElements(sourceTabs, opts.defaultChoice);
        if (elements && elements.length > 0) {
          templateContentsByToolType[tt] = extractTextFromElements4(elements);
        }
      }

      failedTemplateUrl = null; // clear — fetch succeeded
    } catch (e) {
      Logger.log('initAiSession4 template error: ' + e.message);
      templateAccessFailed = true;
      templateError        = e.message || String(e);
      // failedTemplateUrl remains set so the sidebar can offer to open it
    }
  }

  return {
    moduleContext:               moduleContext,
    templateContentsByToolType:  templateContentsByToolType,
    templateAccessFailed:        templateAccessFailed,
    failedTemplateUrl:           failedTemplateUrl,
    templateError:               templateError
  };
}

// ── ELEMENT TEXT EXTRACTION ───────────────────────────────────────────

/**
 * Extracts readable plain text from an array of GDoc elements.
 * Used to give the AI a style reference from the template document.
 *
 * @param {Array<GoogleAppsScript.Document.Element>} elements
 * @returns {string}
 */
function extractTextFromElements4(elements) {
  var lines = [];
  for (var i = 0; i < elements.length; i++) {
    var el   = elements[i];
    var type = el.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      var text = el.asParagraph().getText().trim();
      if (text) lines.push(text);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      var listText = el.asListItem().getText().trim();
      if (listText) lines.push('- ' + listText);
    }
  }
  return lines.join('\n');
}

// ── GENERATE & INSERT ONE ACTIVITY ────────────────────────────────────

/**
 * Generates AI-written directions for one activity slot and inserts them
 * into the "Directions go here…" placeholder in the specified module.
 *
 * Called sequentially from the sidebar — once per activity.
 *
 * @param {Object} params
 *   .apiKey           {string}
 *   .moduleTitle      {string}  e.g. "Module 1"
 *   .activityTitle    {string}  stripped (no number prefix / time suffix)
 *   .toolType         {string|null}
 *   .moduleContext    {Object|null}   Course Design Map data for target module
 *   .templateText     {string|null}  plain-text style reference
 * @returns {{ success: boolean, error?: string }}
 */
function generateAndInsertOneActivity4(params) {
  try {
    var doc     = DocumentApp.getActiveDocument();
    var devBody = getDevelopmentTabBody(doc);
    if (!devBody) throw new Error('Could not find the Development tab in this document.');

    // Find the placeholder for this activity
    var placeholder = findPlaceholderInTool4(devBody, params.moduleTitle, params.activityTitle);
    if (!placeholder) {
      throw new Error(
        'Could not find “Directions go here…” placeholder for “' +
        params.activityTitle + '” in ' + params.moduleTitle + '.'
      );
    }

    // Capture the slot's indent before removing the placeholder
    var indent = placeholder.getIndentStart() || INDENT_4;

    // Build prompt and call Gemini
    var prompt = buildAiDirectionsPrompt4(params);
    var aiText = callGemini4_(params.apiKey, prompt, GEMINI_PRIMARY_4);

    // Remove placeholder and insert formatted content at the same position
    var insertIdx = devBody.getChildIndex(placeholder);
    placeholder.removeFromParent();
    insertFormattedText4(devBody, insertIdx, aiText, indent);

    return { success: true };
  } catch (e) {
    Logger.log('generateAndInsertOneActivity4 error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── FIND PLACEHOLDER IN MODULE ────────────────────────────────────────

/**
 * Locates the "Directions go here…" paragraph for a specific activity
 * within a specific module.
 *
 * @param {GoogleAppsScript.Document.Body} devBody
 * @param {string} moduleTitle    e.g. "Module 1"
 * @param {string} activityTitle  stripped title (no prefix/suffix)
 * @returns {GoogleAppsScript.Document.Paragraph|null}
 */
function findPlaceholderInTool4(devBody, moduleTitle, activityTitle) {
  var H2       = DocumentApp.ParagraphHeading.HEADING2;
  var H4       = DocumentApp.ParagraphHeading.HEADING4;
  var escaped  = moduleTitle.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  var moduleRe = new RegExp('^' + escaped + '[:\\s]', 'i');
  var inModule = false;
  var n        = devBody.getNumChildren();

  for (var i = 0; i < n; i++) {
    var child = devBody.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    var para    = child.asParagraph();
    var heading = para.getHeading();
    var text    = para.getText().trim();

    if (heading === H2) {
      if (moduleRe.test(text)) { inModule = true;  continue; }
      if (inModule)             { break; }           // left the module
      continue;
    }
    if (!inModule) continue;

    if (heading === H4) {
      if (stripActivityHeading(text).toLowerCase() === activityTitle.toLowerCase()) {
        return findDirectionsPlaceholder(devBody, para);
      }
    }
  }
  return null;
}

// ── PROMPT BUILDER ────────────────────────────────────────────────────

/**
 * Builds the Gemini prompt for a single activity.
 *
 * @param {Object} params  same shape as generateAndInsertOneActivity4 params
 * @returns {string}
 */
function buildAiDirectionsPrompt4(params) {
  var moduleContext = params.moduleContext || null;

  // Module context block
  var contextLines = [];
  if (moduleContext) {
    if (moduleContext.moduleTitle) {
      contextLines.push('Module title: ' + moduleContext.moduleTitle);
    }
    if (moduleContext.clos) {
      contextLines.push('Course/Module Learning Outcomes: ' + moduleContext.clos);
    }
    if (moduleContext.readings) {
      contextLines.push('Reading/video content: ' + moduleContext.readings);
    }
    if (moduleContext.activityDescriptions && moduleContext.activityDescriptions.length > 0) {
      contextLines.push('Activity notes from course design: ' +
        moduleContext.activityDescriptions.join('; '));
    }
  }
  var contextBlock = contextLines.length > 0
    ? 'MODULE CONTEXT:\n' + contextLines.join('\n') + '\n\n'
    : '';

  // Template reference block
  var templateBlock = params.templateText
    ? 'STYLE REFERENCE — standard directions for this activity type. Use this as a model ' +
      'for tone, structure, and formatting conventions, but write content specific to this ' +
      'module\'s learning outcomes and activities:\n' +
      params.templateText + '\n\n'
    : '';

  return (
    'You are an instructional designer writing student-facing activity instructions ' +
    'for a college online course.\n\n' +
    contextBlock +
    'ACTIVITY:\n' +
    'Title: '            + params.activityTitle + '\n' +
    'Canvas tool type: ' + (params.toolType || 'unspecified') + '\n\n' +
    templateBlock +
    'FORMATTING RULES — follow exactly:\n' +
    '- For section headings, write the heading text and then the Canvas style in parentheses ' +
    'on the same line. Example: "Overview (H2)" means "Overview" is the heading text and ' +
    '"(H2)" is the Canvas Heading 2 style. Use (H2) for main sections, (H3) for sub-sections.\n' +
    '- Do NOT use markdown heading syntax (# or ##). Only use the (H2)/(H3) parenthetical convention.\n' +
    '- For bullet lists, begin each item with "- " (dash then space).\n' +
    '- For inline bold text, wrap with **double asterisks**, e.g. **Submit by Sunday**.\n' +
    '- Separate paragraphs with a blank line.\n' +
    '- Use "---" on its own line to insert a horizontal rule between major sections.\n' +
    '- Omit any section entirely if it has no applicable content for this activity — ' +
    'do not write "no X assigned" or similar placeholder text.\n' +
    '- Do not start with the activity title as a heading — begin directly with content.\n' +
    '- Aim for 150–400 words. Keep instructions focused and actionable for students.\n\n' +
    'Write student-facing directions for this activity now.'
  );
}

// ── GEMINI REST API ───────────────────────────────────────────────────

/**
 * Calls the Gemini generateContent REST endpoint via UrlFetchApp.
 * Throws on HTTP error, API error, or empty response.
 *
 * @param {string} apiKey
 * @param {string} prompt
 * @param {string} model   e.g. GEMINI_PRIMARY_4 or GEMINI_FAST_4
 * @returns {string}  generated text
 */
function callGemini4_(apiKey, prompt, model, opts) {
  var url = GEMINI_BASE_URL_4 + model + ':generateContent';

  // generationConfig defaults, with optional per-call overrides (e.g. a larger
  // maxOutputTokens for a batched call that returns a whole module at once).
  var genConfig = { temperature: 0.7 };
  if (opts && opts.maxOutputTokens) genConfig.maxOutputTokens = opts.maxOutputTokens;

  var payload = {
    contents:         [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: genConfig
  };

  var options = {
    method:             'post',
    contentType:        'application/json',
    headers:            { 'x-goog-api-key': apiKey },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // Retry with exponential backoff on rate-limit (429) and transient 5xx errors.
  var maxAttempts = 5;
  var response, responseCode, data;

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    response     = UrlFetchApp.fetch(url, options);
    responseCode = response.getResponseCode();

    if (responseCode !== 429 && responseCode < 500) break; // success or non-retryable error

    if (attempt < maxAttempts) {
      // Exponential backoff (2s, 4s, 8s, 16s + jitter) as the floor, but honor
      // the API's own RetryInfo hint (e.g. "retry in 11s") when it asks for
      // longer — waiting less than requested just guarantees another 429.
      var backoffMs = Math.pow(2, attempt) * 1000 + Math.floor(Math.random() * 500);
      var serverMs  = extractRetryDelayMs4_(response);
      var waitMs    = Math.min(Math.max(backoffMs, serverMs), 60000); // cap at 60s
      Logger.log('callGemini4_ HTTP ' + responseCode + ' — retry ' + attempt +
                 ' of ' + (maxAttempts - 1) + ' after ' + waitMs + 'ms' +
                 (serverMs ? ' (server hint ' + serverMs + 'ms)' : ''));
      Utilities.sleep(waitMs);
    }
  }

  data = JSON.parse(response.getContentText());

  if (responseCode !== 200) {
    var apiMsg = (data.error && data.error.message)
      ? data.error.message
      : 'Gemini API error (HTTP ' + responseCode + ')';
    // Friendlier message for rate limits that survived all retries
    if (responseCode === 429) {
      apiMsg = 'Gemini rate limit reached (429) after ' + maxAttempts +
               ' attempts. Wait a minute and try again, or use a lighter model. (' + apiMsg + ')';
    }
    throw new Error(apiMsg);
  }

  if (!data.candidates ||
      !data.candidates[0] ||
      !data.candidates[0].content ||
      !data.candidates[0].content.parts ||
      !data.candidates[0].content.parts[0]) {
    throw new Error('Gemini returned no content. The prompt may have been blocked by safety filters.');
  }

  // Truncated output (hit the token cap) would otherwise be returned as if it
  // were complete — for a batched module reply that silently deploys partial
  // directions. Fail instead, so the caller flags/retries rather than inserting
  // a cut-off block.
  if (data.candidates[0].finishReason === 'MAX_TOKENS') {
    throw new Error('Gemini response hit the output length limit (MAX_TOKENS) and was ' +
                    'truncated. Try deploying fewer modules at once, or a smaller module.');
  }

  return data.candidates[0].content.parts[0].text || '';
}

/**
 * Pulls the server-suggested retry delay out of a 429/5xx response body.
 * Gemini returns it as a RetryInfo detail, e.g.
 *   { error: { details: [ { "@type": "...RetryInfo", retryDelay: "11s" } ] } }
 *
 * @param {GoogleAppsScript.URL_Fetch.HTTPResponse} response
 * @returns {number} delay in milliseconds, or 0 if none was provided
 */
function extractRetryDelayMs4_(response) {
  try {
    var data    = JSON.parse(response.getContentText());
    var details = data && data.error && data.error.details;
    if (!details || !details.length) return 0;
    for (var i = 0; i < details.length; i++) {
      var rd = details[i] && details[i].retryDelay;
      if (rd) {
        // e.g. "11s", "7.2s", or "1200ms"
        var m = String(rd).match(/([\d.]+)\s*(ms|s)?/);
        if (m) {
          var val = parseFloat(m[1]);
          return (m[2] === 'ms') ? Math.round(val) : Math.round(val * 1000);
        }
      }
    }
  } catch (e) { /* body not JSON or shape unexpected — fall back to backoff */ }
  return 0;
}

/**
 * Validates a Gemini API key with a minimal test call.
 *
 * @param {string} apiKey
 * @returns {{ valid: boolean, error?: string }}
 */
function validateGeminiKey4(apiKey) {
  try {
    if (!apiKey || !apiKey.trim()) return { valid: false, error: 'No API key provided.' };
    callGemini4_(apiKey.trim(), 'Reply with the single word: ok', GEMINI_FAST_4);
    return { valid: true };
  } catch (e) {
    return { valid: false, error: e.message };
  }
}

// ── FORMATTED TEXT INSERTION ──────────────────────────────────────────

/**
 * Parses AI-generated text and inserts styled GDoc elements at insertIdx.
 *
 * Recognised patterns:
 *   "Heading text (H2)"  → Normal paragraph; heading text in black,
 *                           (H2) marker in red #ff0000 bold
 *   "Heading text (H3)"  → Same treatment
 *   "- Item text"        → Bullet list item
 *   "**bold**"           → Inline bold within any element
 *   Blank line           → Paragraph separator (skipped)
 *   Everything else      → Normal paragraph
 *
 * Elements are inserted in reverse order so the final result reads
 * correctly top-to-bottom (follows the pattern used throughout Code.gs
 * and Code2.gs to avoid index drift).
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {number} insertIdx
 * @param {string} aiText
 * @param {number} indent   left indent in points for inserted paragraphs
 */
function insertFormattedText4(body, insertIdx, aiText, indent) {
  indent = indent || INDENT_4;

  // Strip markdown code fences if the AI includes them despite instructions
  var fenceMatch = aiText.match(/```(?:markdown|text)?\n?([\s\S]*?)\n?```/);
  if (fenceMatch) aiText = fenceMatch[1].trim();

  // ── Pass 1: tokenise ───────────────────────────────────────────────
  var tokens = [];
  var lines  = aiText.split('\n');

  for (var i = 0; i < lines.length; i++) {
    var trimmed = lines[i].trim();
    if (!trimmed) continue; // blank lines are paragraph separators — skip

    // Horizontal rule
    if (trimmed === '---') {
      tokens.push({ type: 'rule' });
      continue;
    }

    // Bullet list item: starts with "- " or "* "
    if (/^[-*]\s+/.test(trimmed)) {
      tokens.push({ type: 'list', text: trimmed.replace(/^[-*]\s+/, '') });
      continue;
    }

    // Heading marker: ends with (H2), (H3), or (H4)
    var hMatch = trimmed.match(/^(.*?)\s*(\(H[2-4]\))\s*$/i);
    if (hMatch) {
      // Strip markdown heading prefix (###) the AI sometimes emits despite instructions
      var headingText = hMatch[1].trim().replace(/^#+\s*/, '');
      tokens.push({ type: 'heading', text: headingText, marker: hMatch[2] });
      continue;
    }

    tokens.push({ type: 'para', text: trimmed });
  }

  // If AI returned nothing usable, restore the placeholder so the slot isn't empty
  if (tokens.length === 0) {
    var fallback = body.insertParagraph(insertIdx, 'Directions go here…');
    fallback.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    fallback.setIndentStart(indent);
    _fmt(fallback.editAsText(), { font: FONT, size: 11, bold: false, italic: false, color: BLACK });
    return;
  }

  // ── Pass 2: insert in reverse order ───────────────────────────────
  for (var j = tokens.length - 1; j >= 0; j--) {
    var token = tokens[j];

    if (token.type === 'list') {
      var li = body.insertListItem(insertIdx, '');
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
      li.setIndentStart(indent);
      setParagraphText4(li, token.text, BLACK, false);

    } else if (token.type === 'rule') {
      body.insertHorizontalRule(insertIdx);

    } else if (token.type === 'heading') {
      var segs        = parseBoldSegments4(token.text);
      var cleanHeading = segs.map(function(s) { return s.text; }).join('');
      var fullText    = cleanHeading + ' ' + token.marker;
      var markerStart = cleanHeading.length + 1; // +1 for the space
      var markerEnd   = fullText.length - 1;

      var hPara = body.insertParagraph(insertIdx, fullText);
      hPara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      hPara.setIndentStart(indent);

      var pt = hPara.editAsText();
      _fmt(pt, { font: FONT, size: 11, bold: false, italic: false, color: BLACK });

      // (HX) marker: red and bold (applied inline — character positions are known here)
      if (markerStart <= markerEnd) {
        pt.setForegroundColor(markerStart, markerEnd, RED);
        pt.setBold(markerStart, markerEnd, true);
      }
      // Heading text bold is applied in the post-processing pass below

    } else {
      // Normal paragraph
      var para = body.insertParagraph(insertIdx, '');
      para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      para.setIndentStart(indent);
      setParagraphText4(para, token.text, BLACK, false);
    }
  }

  // Post-processing pass: bold heading text after all elements are inserted.
  // Done here (not inline) so the GAS batch has fully flushed before we apply bold.
  applyHeadingBold4(body, insertIdx, tokens.length);
}

// ── HEADING BOLD POST-PROCESSOR ───────────────────────────────────────

/**
 * Scans inserted paragraphs and bolds the heading text (everything before
 * the space + (HX) marker). Called after all elements are inserted so the
 * GAS write batch has flushed and setBold calls are not silently dropped.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {number} startIdx   index where insertion began
 * @param {number} tokenCount number of tokens that were inserted
 */
function applyHeadingBold4(body, startIdx, tokenCount) {
  var headingRe  = /^(.*?)\s*\(H[2-4]\)\s*$/i;
  var end        = startIdx + tokenCount;
  var childCount = body.getNumChildren();   // cache once — avoid a server round-trip per iteration
  for (var i = startIdx; i < end && i < childCount; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para = child.asParagraph();
    var text = para.getText();
    var m    = text.match(headingRe);
    if (!m || !m[1] || m[1].length === 0) continue;
    para.editAsText().setBold(0, m[1].length - 1, true);
  }
}

// ── TEXT FORMATTING HELPERS ───────────────────────────────────────────

/**
 * Sets the text of a paragraph or list item, processes **bold** markers,
 * and applies base font formatting.
 *
 * @param {GoogleAppsScript.Document.Paragraph|GoogleAppsScript.Document.ListItem} el
 * @param {string}  rawText  may contain **bold** markers
 * @param {string}  color    foreground colour (hex)
 * @param {boolean} bold     base bold state
 */
function setParagraphText4(el, rawText, color, bold) {
  var segs      = parseBoldSegments4(rawText);
  var cleanText = segs.map(function(s) { return s.text; }).join('');

  el.setText(cleanText);
  var pt = el.editAsText();
  pt.setFontFamily(FONT);
  pt.setFontSize(11);
  pt.setBold(bold || false);
  pt.setItalic(false);
  pt.setForegroundColor(color || BLACK);

  var pos = 0;
  for (var i = 0; i < segs.length; i++) {
    var len = segs[i].text.length;
    if (len > 0 && segs[i].bold) {
      pt.setBold(pos, pos + len - 1, true);
    }
    pos += len;
  }
}

/**
 * Splits a string with **...** bold markers into an array of segments:
 *   [{ text: string, bold: boolean }, ...]
 *
 * @param {string} raw
 * @returns {Array<{text: string, bold: boolean}>}
 */
function parseBoldSegments4(raw) {
  var segments = [];
  var re       = /\*\*(.+?)\*\*/g;
  var match;
  var lastEnd  = 0;

  while ((match = re.exec(raw)) !== null) {
    if (match.index > lastEnd) {
      segments.push({ text: raw.slice(lastEnd, match.index), bold: false });
    }
    segments.push({ text: match[1], bold: true });
    lastEnd = match.index + match[0].length;
  }
  if (lastEnd < raw.length) {
    segments.push({ text: raw.slice(lastEnd), bold: false });
  }
  if (segments.length === 0) {
    segments.push({ text: raw, bold: false });
  }
  return segments;
}
