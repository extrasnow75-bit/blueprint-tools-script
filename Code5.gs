// ============================================================
// Blueprint Tools — Code5.gs
// Create Model Module (No AI): inserts templated activity
// directions into one user-selected module only.
//
// Relies on shared helpers in Code2.gs (same GAS namespace):
//   extractDocId, collectTabs, getDevelopmentTabBody,
//   findDirectionsPlaceholder, findDirectionElements,
//   replaceWithCopiedElements, stripActivityHeading,
//   DIRECTIONS_PLACEHOLDER_TEXT, DEFAULT_SOURCE_URL
// ============================================================

// ── SIDEBAR OPENER ───────────────────────────────────────────

function showModelModuleNoAiSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar5')
    .setTitle('Create Model Module (No AI)')
    .setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}


// ── MAIN SERVER FUNCTION ─────────────────────────────────────

/**
 * Applies user-chosen direction selections to every activity placeholder
 * in one module of the Development tab.
 *
 * @param {Object} params
 *   .moduleTitle {string}  e.g. "Module 1"
 *   .sourceUrl   {string}  URL of the directions source document
 *   .selections  {Object}  { activityTitle → directionName }
 * @returns {string}  plain-text summary shown in the sidebar
 */
function applyDirectionsToModule5(params) {
  var moduleTitle = params.moduleTitle;
  var sourceUrl   = params.sourceUrl;
  var selections  = params.selections;

  // ── Open source document ────────────────────────────────
  var sourceDocId = extractDocId(sourceUrl);
  if (!sourceDocId) throw new Error('Invalid source document URL.');
  var sourceDoc  = DocumentApp.openById(sourceDocId);
  var sourceTabs = collectTabs(sourceDoc);

  // ── Open blueprint Development tab ──────────────────────
  var blueprintDoc = DocumentApp.getActiveDocument();
  var devBody      = getDevelopmentTabBody(blueprintDoc);
  if (!devBody) throw new Error('Could not find a "Development" tab in this document.');

  // Build a regex that matches the target module's H2 heading.
  var escaped  = moduleTitle.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  var moduleRe = new RegExp('^' + escaped + '[:\\s]', 'i');

  var H2 = DocumentApp.ParagraphHeading.HEADING2;
  var H4 = DocumentApp.ParagraphHeading.HEADING4;

  var replacedCount = 0;
  var skippedCount  = 0;
  var notFoundCount = 0;
  var inTarget      = false;
  var numChildren   = devBody.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = devBody.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    var para    = child.asParagraph();
    var heading = para.getHeading();
    var text    = para.getText().trim();

    // ── Track module boundaries ──────────────────────────
    if (heading === H2) {
      if (moduleRe.test(text)) {
        inTarget = true;
      } else if (inTarget) {
        break; // moved past the target module — done
      }
      continue;
    }

    if (!inTarget || heading !== H4) continue;

    // ── Activity slot found ──────────────────────────────
    var placeholder = findDirectionsPlaceholder(devBody, para, i);
    if (!placeholder) continue;

    var actTitle      = stripActivityHeading(para.getText());
    var directionName = selections[actTitle];
    if (!directionName) continue;

    if (directionName === '[Leave Blank]') {
      skippedCount++;
      continue;
    }

    var sourceElements = findDirectionElements(sourceTabs, directionName);
    if (!sourceElements || sourceElements.length === 0) {
      notFoundCount++;
      Logger.log('applyDirectionsToModule5: direction not found in source doc: ' + directionName);
      continue;
    }

    replaceWithCopiedElements(devBody, placeholder, sourceElements);
    replacedCount++;

    // Body grew — keep index in sync.
    var delta = sourceElements.length - 1;
    numChildren += delta;
    i += delta;
  }

  return (
    '✅ Directions applied to ' + moduleTitle + '!\n\n' +
    'Directions inserted: ' + replacedCount  + '\n' +
    'Left blank:          ' + skippedCount   + '\n' +
    'Not found in source: ' + notFoundCount  + '\n\n' +
    'Review and revise the inserted directions, then run\n' +
    '"Deploy Activity Directions" to push them to other modules.'
  );
}
