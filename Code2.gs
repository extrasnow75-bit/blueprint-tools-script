// ============================================================
// Blueprint Tools — Code2.gs
// Adds activity directions to the Development tab of a blueprint doc.
// ============================================================

// -----------------------------------------------------------
// CONSTANTS
// -----------------------------------------------------------

const DEFAULT_SOURCE_URL =
  'https://docs.google.com/document/d/1jTBx3itcf-sI8RFbxzFEpGjkPZ-uFRsXJlWcgKhgywM/edit';

const DIRECTIONS_PLACEHOLDER_TEXT = 'Directions go here\u2026'; // U+2026 ellipsis

// Direction options keyed by tool type (as it appears before the semicolon in the doc).
const DIRECTION_OPTIONS = {
  'Page': {
    defaultChoice: 'Readings and Multimedia',
    options: [
      'Readings and Multimedia',
      'Zoom Meetings (Scheduled)',
      'Proctored Exams Instructions Page',
      'Course Evaluation Page',
      '[Leave Blank]'
    ]
  },
  'Assignment': {
    defaultChoice: 'Document Upload',
    options: [
      'Document Upload',
      'Text Entry',
      'Video Assignment (Panopto)',
      'Group Assignment Submission',
      'Peer Review of Assignment',
      'Peer Review of Assignment - Reminder',
      'Journal Assignment (Private: Student-Instructor)',
      'Journal Discussion (Not Private: Student-Student)',
      'LTI Assignments (External Tool Integrations)',
      'Perusall LTI Reading Annotations Assignment',
      'Perusall LTI Video Annotations Assignment',
      'Canvas Annotation Tool Practice Activity',
      'Canvas Annotation Tool Activity',
      'Course Evaluation Assignment',
      'Proof of Completion: AI Literacy Course',
      '[Leave Blank]'
    ]
  },
  'Assignment (Not Graded)': {
    defaultChoice: 'Discussion Reply Reminder',
    options: [
      'Discussion Reply Reminder',
      'Video Discussion Reply Reminder',
      '[Leave Blank]'
    ]
  },
  'Discussion': {
    defaultChoice: 'Discussion with Reply Reminder',
    options: [
      'Discussion with Reply Reminder',
      'Video Discussion with Reply Reminder',
      'Discussion: Class Introductions',
      'Video Discussion: Class Introductions',
      '[Leave Blank]'
    ]
  },
  'Quiz (Classic)': {
    defaultChoice: 'Module Quiz',
    options: [
      'Module Quiz',
      'Syllabus Quiz',
      'Proctored Exam',
      'Module Evaluation',
      'Mid-Course Evaluation',
      '[Leave Blank]'
    ]
  },
  'Quiz (New)': {
    defaultChoice: 'Module Quiz',
    options: [
      'Module Quiz',
      '[Leave Blank]'
    ]
  }
};

// Known tool-type strings that appear as standalone lines in the source doc (skip them).
const KNOWN_TOOL_NAMES = [
  'page',
  'assignment',
  'assignment (not graded)',
  'discussion',
  'quiz (classic)',
  'quiz (new)'
];


// -----------------------------------------------------------
// MENU
// -----------------------------------------------------------

/**
 * Creates the custom menu when the document is opened.
 */
// -----------------------------------------------------------
// SIDEBAR
// -----------------------------------------------------------

/**
 * Opens the directions sidebar (360 px wide).
 */
function showDirectionsSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar2')
    .setTitle('Deploy Activity Directions (No AI)')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}


// -----------------------------------------------------------
// SERVER FUNCTIONS CALLED BY SIDEBAR
// -----------------------------------------------------------

/**
 * Returns the data the sidebar needs on load:
 *   { defaultUrl, activities, directionOptions, moduleList }
 *
 * activities:    [{title, toolType}] — deduplicated, ordered (from first module)
 * directionOptions: the DIRECTION_OPTIONS constant (serialised for the client)
 * moduleList:    ['Module 1', 'Module 2', …] — all numbered modules found
 */
/**
 * Single-pass scan of the Development tab body that returns both the ordered
 * module list and each module's deduplicated activity pattern.
 * Avoids the double tab-walk that separate getModuleList / getActivityPattern calls produce.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @returns {{ moduleList: string[], activitiesByModule: Object }}
 */
function collectAllModuleActivities(body) {
  var H2 = DocumentApp.ParagraphHeading.HEADING2;
  var H4 = DocumentApp.ParagraphHeading.HEADING4;
  var moduleList         = [];
  var activitiesByModule = {};
  var seenByModule       = {};
  var currentModule      = null;
  var numChildren        = body.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para    = child.asParagraph();
    var heading = para.getHeading();
    var text    = para.getText().trim();

    if (heading === H2) {
      var match = text.match(/^((?:module|week)\s+\d+)/i);
      if (match) {
        currentModule = match[1];
        moduleList.push(currentModule);
        activitiesByModule[currentModule] = [];
        seenByModule[currentModule]       = {};
      } else {
        currentModule = null;
      }
      continue;
    }

    if (!currentModule) continue;

    if (heading === H4) {
      var actTitle = stripActivityHeading(text);
      var toolType = getToolTypeForSlot(body, para);
      if (!seenByModule[currentModule][actTitle]) {
        seenByModule[currentModule][actTitle] = true;
        activitiesByModule[currentModule].push({ title: actTitle, toolType: toolType });
      }
    }
  }

  return { moduleList: moduleList, activitiesByModule: activitiesByModule };
}

function getSidebarData() {
  var doc  = DocumentApp.getActiveDocument();
  var body = getDevelopmentTabBody(doc);

  if (!body) {
    return {
      defaultUrl:         DEFAULT_SOURCE_URL,
      activities:         [],
      directionOptions:   DIRECTION_OPTIONS,
      moduleList:         ['Module 1'],
      activitiesByModule: { 'Module 1': [] }
    };
  }

  var data       = collectAllModuleActivities(body);
  var moduleList = data.moduleList.length > 0 ? data.moduleList : ['Module 1'];

  return {
    defaultUrl:         DEFAULT_SOURCE_URL,
    activities:         data.activitiesByModule[moduleList[0]] || [],
    directionOptions:   DIRECTION_OPTIONS,
    moduleList:         moduleList,
    activitiesByModule: data.activitiesByModule
  };
}

/**
 * Returns the activity pattern for a specific module title.
 * Called by the sidebar when the user changes the model-module dropdown.
 *
 * @param {string} moduleTitle — e.g. "Module 2" or "Week 3"
 * @returns {Array<{title: string, toolType: string|null}>}
 */
function getActivityPatternForModule(moduleTitle) {
  return getActivityPattern(moduleTitle);
}

/**
 * Returns an ordered list of numbered module titles found in the Development tab.
 * e.g. ['Module 1', 'Module 2', 'Module 3']
 *
 * @returns {Array<string>}
 */
function getModuleList() {
  var doc  = DocumentApp.getActiveDocument();
  var body = getDevelopmentTabBody(doc);
  if (!body) return ['Module 1'];

  var modules     = [];
  var H2          = DocumentApp.ParagraphHeading.HEADING2;
  var numChildren = body.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para = child.asParagraph();
    if (para.getHeading() !== H2) continue;
    var match = para.getText().trim().match(/^((?:module|week)\s+\d+)/i);
    if (match) modules.push(match[1]);
  }

  return modules.length > 0 ? modules : ['Module 1'];
}


/**
 * Given a Google Doc URL (or bare doc ID), opens the document and returns its title.
 * Called by the sidebar when the URL field loses focus.
 *
 * @param {string} url
 * @returns {string} document title, or an error message starting with "Error:"
 */
function getDocTitleFromUrl(url) {
  try {
    var docId = extractDocId(url);
    if (!docId) return 'Error: Could not read the document ID from that URL. Please check the link and try again.';
    var doc = DocumentApp.openById(docId);
    return doc.getName();
  } catch (e) {
    if (e.message && e.message.toLowerCase().includes('action not allowed')) {
      return 'Error: You don’t have permission to access this document. Make sure it’s shared with your Google account.';
    }
    return 'Error: Could not open the document. ' + e.message;
  }
}


/**
 * Main entry point called by the sidebar Run button.
 *
 * @param {Object} params
 *   params.sourceUrl  {string}  — URL of the Activity Directions source document
 *   params.selections {Object}  — map of activityTitle → directionName
 *
 * @returns {string} summary text
 */
function applyDirections(params) {
  var sourceUrl  = params.sourceUrl;
  var selections = params.selections; // { "Readings and Multimedia": "Readings and Multimedia", ... }

  // ---- Open source document and collect all tab bodies ----
  var sourceDocId = extractDocId(sourceUrl);
  if (!sourceDocId) throw new Error('Invalid source document URL.');
  var sourceDoc  = DocumentApp.openById(sourceDocId);
  var sourceTabs = collectTabs(sourceDoc);

  // ---- Open the blueprint (active) document and find the Development tab ----
  var blueprintDoc = DocumentApp.getActiveDocument();
  var devBody = getDevelopmentTabBody(blueprintDoc);
  if (!devBody) throw new Error('Could not find a "Development" tab in this document.');

  // ---- Walk the blueprint body, find every H4 activity slot, replace placeholder ----
  var replacedCount   = 0;
  var skippedCount    = 0;
  var notFoundCount   = 0;

  var numChildren = devBody.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = devBody.getChild(i);

    // We only care about H4 paragraphs (activity headings).
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para = child.asParagraph();
    if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING4) continue;

    // Find the "Directions go here…" paragraph inside this slot.
    var placeholder = findDirectionsPlaceholder(devBody, para);
    if (!placeholder) continue;

    // Determine which activity title this slot belongs to (strip number + time estimate).
    var rawTitle    = para.getText();
    var actTitle    = stripActivityHeading(rawTitle);

    // Look up the user's selection for this activity title.
    var directionName = selections[actTitle];
    if (!directionName) {
      // Title not in selections map — skip silently (shouldn't happen if sidebar is correct).
      continue;
    }

    // "[Leave Blank]" means do not touch the placeholder.
    if (directionName === '[Leave Blank]') {
      skippedCount++;
      continue;
    }

    // Find the content elements in the source document.
    var sourceElements = findDirectionElements(sourceTabs, directionName);
    if (!sourceElements || sourceElements.length === 0) {
      notFoundCount++;
      Logger.log('Direction not found in source doc: ' + directionName);
      continue;
    }

    // Replace the placeholder paragraph with the copied content.
    replaceWithCopiedElements(devBody, placeholder, sourceElements);
    replacedCount++;

    // After inserting, the body has grown — recalculate numChildren and i.
    // The placeholder was removed (–1) and N elements were inserted (+N).
    var delta = sourceElements.length - 1;
    numChildren += delta;
    i += delta;
  }

  // ---- Build summary string ----
  var summary =
    '\u2705 Activity Directions Added!\n\n' +
    'Directions replaced: '     + replacedCount  + '\n' +
    'Left blank (skipped): '    + skippedCount   + '\n' +
    'Not found in source: '     + notFoundCount;

  return summary;
}


// -----------------------------------------------------------
// ACTIVITY PATTERN DISCOVERY
// -----------------------------------------------------------

/**
 * Scans the specified module of the Development tab and returns an ordered,
 * deduplicated list of activity descriptors found there.
 *
 * @param {string} [moduleTitle] — e.g. "Module 1" or "Week 3". Defaults to "Module 1".
 * @returns {Array<{title: string, toolType: string|null}>}
 */
function getActivityPattern(moduleTitle) {
  moduleTitle = moduleTitle || 'Module 1';
  // Escape regex special chars in the title, then match the H2 heading start.
  var escaped  = moduleTitle.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  var moduleRe = new RegExp('^' + escaped + '[:\\s]', 'i');

  var doc  = DocumentApp.getActiveDocument();
  var body = getDevelopmentTabBody(doc);
  if (!body) {
    Logger.log('getActivityPattern: Development tab not found.');
    return [];
  }

  var activities     = [];
  var seenTitles     = {};
  var inTargetModule = false;

  var numChildren = body.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para = child.asParagraph();

    var heading = para.getHeading();
    var text    = para.getText().trim();

    // Detect H2 boundaries.
    if (heading === DocumentApp.ParagraphHeading.HEADING2) {
      if (moduleRe.test(text)) {
        inTargetModule = true;
        continue;
      }
      if (inTargetModule) {
        // Hit the next H2 — target module is over.
        break;
      }
      continue;
    }

    if (!inTargetModule) continue;

    // H4 = activity slot heading.
    if (heading === DocumentApp.ParagraphHeading.HEADING4) {
      var actTitle  = stripActivityHeading(text);
      var toolType  = getToolTypeForSlot(body, para);

      if (!seenTitles[actTitle]) {
        seenTitles[actTitle] = true;
        activities.push({ title: actTitle, toolType: toolType });
      }
    }
  }

  return activities;
}


/**
 * Given the body and an H4 paragraph (activity heading), returns the tool type
 * found in the "Tool; Link to settings tab" line immediately following the heading.
 *
 * The tool type is the text before the first semicolon, trimmed.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {GoogleAppsScript.Document.Paragraph} headingPara
 * @returns {string|null}
 */
function getToolTypeForSlot(body, headingPara) {
  var startIndex = body.getChildIndex(headingPara);
  var numChildren = body.getNumChildren();

  // Scan forward up to 6 siblings.
  for (var j = startIndex + 1; j < Math.min(startIndex + 7, numChildren); j++) {
    var child = body.getChild(j);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para = child.asParagraph();

    // Stop if we hit another H4 or a higher heading.
    var h = para.getHeading();
    if (h === DocumentApp.ParagraphHeading.HEADING4 ||
        h === DocumentApp.ParagraphHeading.HEADING3 ||
        h === DocumentApp.ParagraphHeading.HEADING2) break;

    var text = para.getText();

    // The "Tool; Link to settings tab" line contains a semicolon.
    if (text.indexOf(';') !== -1) {
      var toolType = text.split(';')[0].trim();
      return toolType || null;
    }
  }

  return null;
}


// -----------------------------------------------------------
// TAB UTILITIES
// -----------------------------------------------------------

/**
 * Returns the Body of the "Development" tab in the given document.
 * Returns null if no such tab is found.
 *
 * @param {GoogleAppsScript.Document.Document} doc
 * @returns {GoogleAppsScript.Document.Body|null}
 */
function getDevelopmentTabBody(doc) {
  var tabs = collectTabs(doc);
  for (var i = 0; i < tabs.length; i++) {
    if (/\bdevelopment\b/i.test(tabs[i].title)) {
      return tabs[i].body;
    }
  }
  return null;
}


/**
 * Recursively walks all tabs (and child tabs) of a document.
 * Returns an array of { title, body } objects.
 *
 * @param {GoogleAppsScript.Document.Document} doc
 * @returns {Array<{title: string, body: GoogleAppsScript.Document.Body}>}
 */
function collectTabs(doc) {
  var result = [];

  function walk(tabs) {
    if (!tabs) return;
    for (var i = 0; i < tabs.length; i++) {
      var tab = tabs[i];
      var docTab = tab.asDocumentTab();
      result.push({
        title: tab.getTitle(),
        body:  docTab.getBody()
      });
      walk(tab.getChildTabs());
    }
  }

  walk(doc.getTabs());
  return result;
}


// -----------------------------------------------------------
// SOURCE DOCUMENT SEARCHING
// -----------------------------------------------------------

/**
 * Searches all tabs of the source document for an H2 heading that exactly matches
 * directionName, then collects the body content that follows (skipping intro lines)
 * until the next H2.
 *
 * Uses body.getNumChildren() / body.getChild(i) to capture ALL element types.
 *
 * @param {Array<{title: string, body: GoogleAppsScript.Document.Body}>} sourceTabs
 * @param {string} directionName  — e.g. "Readings and Multimedia"
 * @returns {Array<GoogleAppsScript.Document.Element>|null}
 */
function findDirectionElements(sourceTabs, directionName) {
  for (var t = 0; t < sourceTabs.length; t++) {
    var body = sourceTabs[t].body;
    var numChildren = body.getNumChildren();

    for (var i = 0; i < numChildren; i++) {
      var child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

      var para = child.asParagraph();
      if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING2) continue;
      if (para.getText().trim() !== directionName) continue;

      // Found the H2. Now find where content actually starts (skip intro lines).
      var contentStartIndex = findContentStartByIndex(body, i);

      // Collect all elements from contentStartIndex until the next H2 (exclusive).
      var elements = [];
      for (var j = contentStartIndex; j < numChildren; j++) {
        var el = body.getChild(j);

        if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
          var elPara = el.asParagraph();
          if (elPara.getHeading() === DocumentApp.ParagraphHeading.HEADING2) {
            // Reached the next section — stop.
            break;
          }
        }

        elements.push(el);
      }

      // Trim trailing blank paragraphs.
      while (elements.length > 0) {
        var last = elements[elements.length - 1];
        if (last.getType() === DocumentApp.ElementType.PARAGRAPH &&
            last.asParagraph().getText().trim() === '') {
          elements.pop();
        } else {
          break;
        }
      }

      return elements;
    }
  }

  return null; // not found in any tab
}


/**
 * Given the body and the index of an H2 heading, returns the index of the first
 * element that is "real content" — i.e., after skipping:
 *   - H3 paragraphs (demo activity headers)
 *   - Blank / smart-chip lines (getText().trim() === '')
 *   - "Estimated time…" lines
 *   - Known tool-name lines
 *   - Demo activity number lines (start with "X.XX" or /^\d+\.\d+/)
 *   - "Note to add…" lines
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {number} h2Index
 * @returns {number}
 */
function findContentStartByIndex(body, h2Index) {
  var H2 = DocumentApp.ParagraphHeading.HEADING2;
  var H3 = DocumentApp.ParagraphHeading.HEADING3;
  var numChildren = body.getNumChildren();

  // Pass 1: find the H3 demo-activity header that marks the start of the
  // example block (look within the next 10 elements to avoid overshooting).
  var h3Index = -1;
  for (var k = h2Index + 1; k < Math.min(h2Index + 10, numChildren); k++) {
    var c = body.getChild(k);
    if (c.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var p = c.asParagraph();
    if (p.getHeading() === H2) break;   // hit the next section — no H3 found
    if (p.getHeading() === H3) { h3Index = k; break; }
  }

  // Pass 2: start scanning from just after the H3 (or after H2 if no H3 found),
  // skipping boilerplate lines until we reach real content.
  var i = (h3Index >= 0) ? h3Index + 1 : h2Index + 1;

  while (i < numChildren) {
    var child = body.getChild(i);
    var type  = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      var para    = child.asParagraph();
      var heading = para.getHeading();
      var text    = para.getText().trim();

      // Stop scanning if we hit another H2 — nothing to copy.
      if (heading === H2) break;

      // Skip any additional H3 headings.
      if (heading === H3)                                              { i++; continue; }

      // Skip blank lines (also catches smart chips which return '').
      if (text === '')                                                 { i++; continue; }

      // Skip "Estimated time" lines.
      if (/^estimated time/i.test(text))                              { i++; continue; }

      // Skip known tool-name lines.
      if (KNOWN_TOOL_NAMES.indexOf(text.toLowerCase()) !== -1)        { i++; continue; }

      // Skip demo activity number prefixes (e.g. "X.XX …" or "1.01 …").
      if (/^X\.XX/i.test(text) || /^\d+\.\d+/.test(text))            { i++; continue; }

      // Skip "Note to add…" instructor notes.
      if (/^note to add/i.test(text))                                 { i++; continue; }

      // Skip the "[Tool chip]; Link to settings tab" line (and any variant).
      if (/link to settings tab/i.test(text))                         { i++; continue; }

      // This line is real content — stop skipping.
      break;
    } else {
      // Non-paragraph element (e.g. table, list) — treat as content.
      break;
    }
  }

  return i;
}


// -----------------------------------------------------------
// PLACEHOLDER FINDING & REPLACEMENT
// -----------------------------------------------------------

/**
 * Starting from the H4 heading paragraph, scans forward to find the
 * "Directions go here…" paragraph within the same activity slot.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {GoogleAppsScript.Document.Paragraph} headingPara
 * @returns {GoogleAppsScript.Document.Paragraph|null}
 */
function findDirectionsPlaceholder(body, headingPara) {
  var startIndex  = body.getChildIndex(headingPara);
  var numChildren = body.getNumChildren();

  for (var i = startIndex + 1; i < Math.min(startIndex + 8, numChildren); i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    var para    = child.asParagraph();
    var heading = para.getHeading();

    // Stop if we leave this slot (hit another H4 or higher).
    if (heading === DocumentApp.ParagraphHeading.HEADING4 ||
        heading === DocumentApp.ParagraphHeading.HEADING3 ||
        heading === DocumentApp.ParagraphHeading.HEADING2) break;

    var text = para.getText();

    // Match the placeholder — handle both U+2026 and three literal dots.
    if (text === DIRECTIONS_PLACEHOLDER_TEXT ||
        text === 'Directions go here...') {
      return para;
    }
  }

  return null;
}


/**
 * Removes the placeholder paragraph and inserts copies of sourceElements
 * at the same position in the correct forward order.
 *
 * @param {GoogleAppsScript.Document.Body}           body
 * @param {GoogleAppsScript.Document.Paragraph}      placeholder
 * @param {Array<GoogleAppsScript.Document.Element>} sourceElements
 */
function replaceWithCopiedElements(body, placeholder, sourceElements) {
  var insertIdx = body.getChildIndex(placeholder);
  placeholder.removeFromParent();

  // Insert in REVERSE order at insertIdx so final order is correct.
  for (var k = sourceElements.length - 1; k >= 0; k--) {
    var el   = sourceElements[k];
    var type = el.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      body.insertParagraph(insertIdx, el.asParagraph().copy());
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      var srcItem = el.asListItem();
      var newItem = body.insertListItem(insertIdx, srcItem.copy());
      // Re-apply glyph type and nesting level — copying a ListItem across
      // documents drops the bullet/number because the source ListId doesn't
      // exist in the destination document.
      newItem.setGlyphType(srcItem.getGlyphType());
      newItem.setNestingLevel(srcItem.getNestingLevel());
    } else if (type === DocumentApp.ElementType.TABLE) {
      body.insertTable(insertIdx, el.asTable().copy());
    } else {
      Logger.log('replaceWithCopiedElements: skipping unsupported element type: ' + type);
    }
  }
}


// -----------------------------------------------------------
// HELPERS
// -----------------------------------------------------------

/**
 * Strips the number prefix (e.g. "1.01 ") and the time estimate suffix
 * (e.g. " (30 min)") from an H4 activity heading, returning just the title.
 *
 * Examples:
 *   "1.01 Readings and Multimedia (30 min)"  →  "Readings and Multimedia"
 *   "2.03 Module Quiz (20 min)"              →  "Module Quiz"
 *
 * @param {string} raw
 * @returns {string}
 */
function stripActivityHeading(raw) {
  // Remove leading number prefix like "1.01 " or "10.03 ".
  var stripped = raw.replace(/^\d+\.\d+\s+/, '');
  // Remove trailing time estimate like " (30 min)" or " (1 hr)".
  stripped = stripped.replace(/\s*\(\s*(?:TBD|\d[\d\w\s./]*(?:min|mins|hr|hrs|hour|hours)?)\s*\)\s*$/i, '');
  return stripped.trim();
}


/**
 * Extracts a Google Doc ID from a full URL or returns the input unchanged
 * if it looks like a bare ID.
 *
 * @param {string} urlOrId
 * @returns {string|null}
 */
function extractDocId(urlOrId) {
  if (!urlOrId) return null;

  // Try to pull /d/<id> from a standard Docs URL.
  var match = urlOrId.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];

  // If it doesn't contain a slash it might already be a bare ID.
  if (urlOrId.indexOf('/') === -1 && urlOrId.length > 20) {
    return urlOrId.trim();
  }

  return null;
}
