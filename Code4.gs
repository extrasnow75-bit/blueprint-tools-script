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
          readings:             '',   // merged (fallback for combined activities)
          readingsText:         '',   // readings-only rows
          videosText:           '',   // videos-only rows
          activityDescriptions: [],
          mediaLinks:           []    // [{ kind: 'video'|'reading', url, text }]
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
        // Keep a merged blob (fallback for combined/ambiguous activities) plus
        // separate readings/videos text so per-activity scoping can pick one.
        currentModule.readings = currentModule.readings
          ? currentModule.readings + '\n' + value
          : value;
        var labelSaysVideo   = label.indexOf('video')   !== -1;
        var labelSaysReading = label.indexOf('reading') !== -1;
        if (labelSaysReading && !labelSaysVideo) {
          currentModule.readingsText = currentModule.readingsText
            ? currentModule.readingsText + '\n' + value : value;
        } else if (labelSaysVideo && !labelSaysReading) {
          currentModule.videosText = currentModule.videosText
            ? currentModule.videosText + '\n' + value : value;
        } else {
          // Combined / "content" row — contributes to both buckets.
          currentModule.readingsText = currentModule.readingsText
            ? currentModule.readingsText + '\n' + value : value;
          currentModule.videosText = currentModule.videosText
            ? currentModule.videosText + '\n' + value : value;
        }
        // Pull the real hyperlink targets so they can be embedded downstream.
        // Tag each link's kind by its URL (a video host → 'video'), falling back
        // to the row label — this stays correct even in a combined cell that
        // mixes readings and videos ("Readings & Multimedia").
        var cellLinks = collectCellLinks_(table.getRow(r).getCell(1));
        for (var lk = 0; lk < cellLinks.length; lk++) {
          var kind = isVideoUrl_(cellLinks[lk].url)
            ? 'video'
            : (labelSaysVideo && !labelSaysReading ? 'video' : 'reading');
          currentModule.mediaLinks.push({
            kind: kind, url: cellLinks[lk].url, text: cellLinks[lk].text
          });
        }
      } else if (label.indexOf('activity')   !== -1 ||
                 label.indexOf('assessment') !== -1) {
        currentModule.activityDescriptions.push(value);
      }
    }

    if (modules.length > 0) break; // found the CDM table — stop searching
  }

  return modules;
}

/**
 * Extracts hyperlink targets from a table cell — both real hyperlink runs
 * (whose display text may be a title rather than the URL) and bare http(s) URLs
 * typed as plain text. Returns [{ url, text }] in document order, de-duplicated
 * by URL. Shared by the Course Design Map parser.
 *
 * @param {GoogleAppsScript.Document.TableCell} cell
 * @returns {Array<{url: string, text: string}>}
 */
function collectCellLinks_(cell) {
  var links = [];
  var seen  = {};
  var n     = cell.getNumChildren();

  for (var i = 0; i < n; i++) {
    var el   = cell.getChild(i);
    var type = el.getType();
    var textEl;
    if      (type === DocumentApp.ElementType.PARAGRAPH) textEl = el.asParagraph().editAsText();
    else if (type === DocumentApp.ElementType.LIST_ITEM) textEl = el.asListItem().editAsText();
    else continue;

    var s = textEl.getText();
    if (!s) continue;

    // 1) Real hyperlink runs — getLinkUrl changes at attribute boundaries.
    var idxs = textEl.getTextAttributeIndices();
    for (var k = 0; k < idxs.length; k++) {
      var start = idxs[k];
      var url   = textEl.getLinkUrl(start);
      if (url && !seen[url]) {
        seen[url] = true;
        var stop = (k + 1 < idxs.length) ? idxs[k + 1] : s.length;
        links.push({ url: url, text: s.substring(start, stop).trim() || url });
      }
    }

    // 2) Bare URLs typed as plain text (visible in the doc but not hyperlinked).
    var re = /https?:\/\/[^\s\]\)]+/g;
    var m;
    while ((m = re.exec(s)) !== null) {
      var bare = m[0].replace(/[.,;]+$/, '');
      if (!seen[bare]) { seen[bare] = true; links.push({ url: bare, text: bare }); }
    }
  }
  return links;
}

/**
 * Builds the media-link scaffolding shared by the AI tools from a module's
 * parsed mediaLinks array. Returns:
 *   linkMap     { TOKEN: url }          — resolves [title](TOKEN) at insert time
 *   promptLines [ "VIDEO1 — context" ]  — lists available links in the prompt
 *   hasLinks    boolean
 * Tokens are VIDEO1, VIDEO2, … and READING1, READING2, … in document order.
 *
 * @param {Array<{kind, url, text}>} mediaLinks
 * @returns {{ linkMap: Object, promptLines: string[], hasLinks: boolean }}
 */
function buildMediaTokens4_(mediaLinks) {
  var linkMap  = {};
  var lines    = [];
  var counters = { video: 0, reading: 0 };

  (mediaLinks || []).forEach(function (lk) {
    if (!lk || !lk.url) return;
    var kind  = (lk.kind === 'video') ? 'video' : 'reading';
    counters[kind]++;
    var token = kind.toUpperCase() + counters[kind];
    linkMap[token] = lk.url;
    // Prefer the fetched page/video title; mark it so the prompt knows it is the
    // canonical title to use verbatim as the link text. Scrub "@@@" and newlines
    // from fetched titles so a crafted page title can't spoof the batched-deploy
    // "@@@ACTIVITY N@@@" delimiter (Code6) or break the prompt across lines.
    var ctx;
    if (lk.title) {
      var safeTitle = String(lk.title).replace(/@{2,}/g, '@').replace(/[\r\n]+/g, ' ');
      ctx = '"' + safeTitle + '" (use this exact title as the link text)';
    } else {
      // Fallback to the doc-supplied hyperlink display text (or the URL). Scrub
      // "@@@" and newlines here too so a crafted link label can't spoof the
      // "@@@ACTIVITY N@@@" delimiter or break the prompt across lines.
      var raw = (lk.text && lk.text !== lk.url) ? lk.text : lk.url;
      ctx = String(raw).replace(/@{2,}/g, '@').replace(/[\r\n]+/g, ' ');
    }
    lines.push(token + ' — ' + ctx);
  });

  return { linkMap: linkMap, promptLines: lines, hasLinks: lines.length > 0 };
}

/**
 * Resolves a markdown link target to a real URL. A target that already looks
 * like an http(s) URL is used as-is; otherwise it is treated as a media TOKEN
 * (e.g. VIDEO1, READING2) and looked up in linkMap. Returns null if unresolved,
 * so the caller can leave the visible text unlinked rather than link to junk.
 *
 * @param {string} target
 * @param {Object|null} linkMap
 * @returns {string|null}
 */
function resolveLinkTarget4_(target, linkMap) {
  if (!target) return null;
  var t = String(target).trim();
  if (/^https?:\/\//i.test(t)) return t;
  if (linkMap) {
    var key = t.toUpperCase().replace(/[\s_\-]/g, '');
    if (linkMap[key]) return linkMap[key];
  }
  return null;
}

// ── PER-ACTIVITY CONTEXT SCOPING ──────────────────────────────────────

/**
 * True when a URL points at a known video host or a video file. Used to tag a
 * link's kind independently of the CDM row label, so a combined cell that mixes
 * readings and videos still classifies each link correctly.
 *
 * @param {string} url
 * @returns {boolean}
 */
function isVideoUrl_(url) {
  return /(?:youtube\.com|youtu\.be|vimeo\.com|panopto\.|kaltura\.|mediaspace|loom\.com|wistia\.|brightcove|\.mp4(?:[?#]|$))/i
    .test(String(url || ''));
}

/**
 * Classifies an activity by its title so its directions can be scoped to the
 * matching slice of module context:
 *   'reading' — readings-only activity
 *   'video'   — videos-only activity
 *   'all'     — combined or unrecognized (gets the FULL module context)
 *
 * Biased toward 'all': the only harmful error is hiding content, so it scopes
 * down only for a simple single-type title. Any conjunction ("and", "&", "+",
 * "/", ",") signals a combined activity ("Readings and Videos", "Readings &
 * Multimedia", "Articles and Videos") and returns 'all'.
 *
 * @param {string} title
 * @returns {'reading'|'video'|'all'}
 */
function classifyActivityKind_(title) {
  var t = String(title || '').toLowerCase();
  if (!t) return 'all';

  // A conjunction joins multiple content types → treat as combined.
  if (/\band\b/.test(t) || /[&/+,]/.test(t)) return 'all';

  var isRead  = /read|article|text|chapter|paper|essay/.test(t);
  var isVideo = /video|watch|multimedia|media|recording|lecture|screencast/.test(t);

  if (isRead && !isVideo) return 'reading';
  if (isVideo && !isRead) return 'video';
  return 'all';
}

/**
 * Returns a module-context object scoped to one activity's kind — a readings
 * activity gets readings text + non-video links; a videos activity gets videos
 * text + video links. A combined/unrecognized activity ('all') gets the shared
 * context unchanged. Shallow-clones before overriding so the module context
 * (reused by sibling activities) is never mutated.
 *
 * @param {Object|null} moduleContext
 * @param {string} activityTitle
 * @returns {Object|null}
 */
function scopeModuleContextForActivity_(moduleContext, activityTitle) {
  if (!moduleContext) return moduleContext;
  var kind = classifyActivityKind_(activityTitle);
  if (kind === 'all') return moduleContext;

  var scoped = {};
  for (var k in moduleContext) {
    if (moduleContext.hasOwnProperty(k)) scoped[k] = moduleContext[k];
  }

  if (kind === 'reading') {
    scoped.readings   = moduleContext.readingsText || moduleContext.readings || '';
    scoped.mediaLinks = (moduleContext.mediaLinks || []).filter(function (lk) {
      return lk && lk.kind !== 'video';
    });
  } else { // 'video'
    scoped.readings   = moduleContext.videosText || moduleContext.readings || '';
    scoped.mediaLinks = (moduleContext.mediaLinks || []).filter(function (lk) {
      return lk && lk.kind === 'video';
    });
  }
  return scoped;
}

// ── LINK METADATA (TITLES) ────────────────────────────────────────────

/**
 * Fetches a human-readable title for each media link in the given module
 * contexts and stores it on the link as `.title`. Uses oEmbed for YouTube/Vimeo
 * (no key, reliable) and a best-effort og:title / <title> scrape for everything
 * else. All fetches run in ONE parallel UrlFetchApp.fetchAll batch and every
 * failure is swallowed — a missing title just falls back to the raw URL text.
 *
 * These are UrlFetchApp calls, NOT Gemini calls, so they add no model-side
 * rate-limit pressure. De-duplicates by URL so a link shared across modules is
 * fetched once.
 *
 * @param {Array<Object>} moduleContexts  parsed CDM module objects (with mediaLinks)
 */
/**
 * True only for URLs that are safe to fetch from Google's egress: an http(s)
 * scheme and a public host. Blocks loopback/private/link-local IP literals and
 * bare intranet names (no dot, or .local/.internal) so a link in a shared doc
 * can't make the tool probe internal endpoints. Public article hosts still pass,
 * so title-scraping keeps working for arbitrary journals/magazines.
 *
 * @param {string} url
 * @returns {boolean}
 */
function isFetchableUrl_(url) {
  var m = String(url || '').trim().match(/^https?:\/\/([^\/?#]+)/i);
  if (!m) return false;                       // non-http(s): javascript:, data:, file:, ftp:, …
  var host = m[1].toLowerCase();
  var at = host.indexOf('@');
  if (at !== -1) host = host.slice(at + 1);   // strip any user:pass@ credentials
  host = host.replace(/:\d+$/, '');           // strip port
  if (!host) return false;
  if (host === 'localhost' || /\.(local|internal|localdomain)$/.test(host)) return false;

  var v4 = host.match(/^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/);
  if (v4) {
    var a = +v4[1], b = +v4[2];
    if (a === 0 || a === 10 || a === 127) return false;          // this-network, private, loopback
    if (a === 169 && b === 254) return false;                    // link-local
    if (a === 192 && b === 168) return false;                    // private
    if (a === 172 && b >= 16 && b <= 31) return false;           // private
    if (a >= 224) return false;                                  // multicast / reserved
    return true;                                                 // public IPv4
  }
  if (host.indexOf(':') !== -1) {                                // IPv6 literal
    if (/^\[?::1\]?$/.test(host)) return false;                  // loopback
    if (/^\[?(fe80|fc|fd)/i.test(host)) return false;            // link-local / unique-local
    return true;
  }
  return host.indexOf('.') !== -1;            // require a dotted public domain
}

function enrichMediaTitles_(moduleContexts) {
  if (!moduleContexts || !moduleContexts.length) return;

  // Group link objects by URL so each unique URL is fetched exactly once.
  var byUrl = {};
  for (var mi = 0; mi < moduleContexts.length; mi++) {
    var mc = moduleContexts[mi];
    if (!mc || !mc.mediaLinks) continue;
    for (var li = 0; li < mc.mediaLinks.length; li++) {
      var lk = mc.mediaLinks[li];
      if (!lk || !lk.url) continue;
      if (!byUrl[lk.url]) byUrl[lk.url] = [];
      byUrl[lk.url].push(lk);
    }
  }

  // Only fetch public http(s) URLs. The links come from a document that may have
  // been authored/shared by someone else, so guard against using UrlFetchApp as
  // an SSRF vector: drop non-http(s) schemes and any host that is loopback,
  // private, link-local, or a bare intranet name, and cap the batch size.
  var urls = Object.keys(byUrl).filter(isFetchableUrl_);
  if (!urls.length) return;
  var MAX_FETCH = 40;
  if (urls.length > MAX_FETCH) urls = urls.slice(0, MAX_FETCH);

  var requests = [];
  var meta     = []; // parallel to requests: { url, mode }
  for (var u = 0; u < urls.length; u++) {
    var oembed = oembedEndpointFor_(urls[u]);
    if (oembed) {
      requests.push({ url: oembed, muteHttpExceptions: true, followRedirects: true });
      meta.push({ url: urls[u], mode: 'oembed' });
    } else {
      requests.push({
        url: urls[u], muteHttpExceptions: true, followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; BlueprintTools/1.0)' }
      });
      meta.push({ url: urls[u], mode: 'html' });
    }
  }

  var responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    Logger.log('enrichMediaTitles_ fetchAll failed: ' + ((e && e.message) || e));
    return; // leave titles unset — callers fall back to URL text
  }

  for (var ri = 0; ri < responses.length; ri++) {
    try {
      var resp = responses[ri];
      if (resp.getResponseCode() !== 200) continue;
      var title = (meta[ri].mode === 'oembed')
        ? extractOembedTitle_(resp.getContentText())
        : extractHtmlTitle_(resp.getContentText());
      if (title) {
        var refs = byUrl[meta[ri].url];
        for (var rf = 0; rf < refs.length; rf++) refs[rf].title = title;
      }
    } catch (e2) {
      Logger.log('enrichMediaTitles_ parse [' + meta[ri].url + ']: ' + ((e2 && e2.message) || e2));
    }
  }
}

/**
 * Returns the oEmbed JSON endpoint for a YouTube or Vimeo URL, or null for any
 * other host (which is handled by the generic HTML scrape instead).
 *
 * @param {string} url
 * @returns {string|null}
 */
function oembedEndpointFor_(url) {
  var u = String(url);
  if (/(?:youtube\.com\/(?:watch|shorts|embed|live)|youtu\.be\/)/i.test(u)) {
    return 'https://www.youtube.com/oembed?url=' + encodeURIComponent(u) + '&format=json';
  }
  if (/(?:player\.)?vimeo\.com\//i.test(u)) {
    return 'https://vimeo.com/api/oembed.json?url=' + encodeURIComponent(u);
  }
  return null;
}

/**
 * Extracts the "title" field from an oEmbed JSON response body.
 *
 * @param {string} body
 * @returns {string}  title, or '' if unparseable/absent
 */
function extractOembedTitle_(body) {
  try {
    var j = JSON.parse(body);
    return (j && j.title) ? String(j.title).trim() : '';
  } catch (e) {
    return '';
  }
}

/**
 * Best-effort title from an HTML page: prefers the Open Graph og:title meta
 * tag, falling back to the <title> element. Returns '' when neither is present
 * (paywalls, JS-rendered pages, bot blocks). Decodes common HTML entities and
 * caps the length.
 *
 * @param {string} html
 * @returns {string}
 */
function extractHtmlTitle_(html) {
  if (!html) return '';
  // Cap length before matching — the title lives in <head>, and matching greedy
  // patterns against a multi-MB body risks burning the script's execution quota.
  html = String(html).slice(0, 200000);
  var m = html.match(/<meta[^>]+property=["']og:title["'][^>]*content=["']([^"']*)["']/i)
       || html.match(/<meta[^>]+content=["']([^"']*)["'][^>]+property=["']og:title["']/i);
  var raw = m ? m[1] : null;
  if (!raw) {
    var tm = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    if (tm) raw = tm[1];
  }
  if (!raw) return '';
  raw = decodeEntities_(raw).replace(/\s+/g, ' ').trim();
  if (raw.length > 200) raw = raw.slice(0, 197) + '…';
  return raw;
}

/**
 * Decodes the handful of HTML entities that commonly appear in page titles.
 *
 * @param {string} s
 * @returns {string}
 */
function decodeEntities_(s) {
  return String(s)
    .replace(/&amp;/g,  '&')
    .replace(/&quot;/g, '"')
    .replace(/&#0*39;/g, "'")
    .replace(/&#x27;/gi, "'")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g,   '<')
    .replace(/&gt;/g,   '>')
    .replace(/&nbsp;/g, ' ');
}

// ── ASSIGNMENT "WHERE TO GO FOR HELP" SECTION ─────────────────────────

/**
 * True for Canvas assignment slots ("assignment" and "assignment (not graded)")
 * — the only tool types that receive the standard help section. Discussions,
 * quizzes, and blank/page slots return false.
 *
 * @param {string|null} toolType
 * @returns {boolean}
 */
function isAssignmentToolType4_(toolType) {
  return !!toolType && /^assignment\b/i.test(String(toolType).trim());
}

/**
 * Reads the active document's name and returns the course code as "DEPT ###"
 * (e.g. "MATH 101", "RESPCARE 516"). Blueprints are named like "MATH 101
 * Blueprint" or "ENGL 101 Blueprint Redesign (SU26)"; only the department +
 * number is used. Returns "CLASS 101" when no code can be found.
 *
 * @returns {string}
 */
function extractCourseCode_() {
  try {
    var name = DocumentApp.getActiveDocument().getName() || '';
    var m = name.match(/\b([A-Za-z]{2,10})\s*(\d{3}[A-Za-z]?)\b/);
    if (m) return m[1].toUpperCase() + ' ' + m[2].toUpperCase();
  } catch (e) {
    Logger.log('extractCourseCode_: ' + ((e && e.message) || e));
  }
  return 'CLASS 101';
}

/**
 * Returns the standard "Where to go for help" block as markdown text, ready to
 * append to an activity's AI directions before insertFormattedText4 renders it.
 * The Help Desk phrase carries a markdown link so the formatter embeds the real
 * OIT URL; the professor bullet is plain guidance text with a suggested subject.
 *
 * @param {string} courseCode     e.g. "MATH 101"
 * @param {string} activityTitle  the activity's title
 * @returns {string}
 */
function assignmentHelpBlock4_(courseCode, activityTitle) {
  var HELP_URL = 'https://www.boisestate.edu/oit/assistance/';
  return '\n\n---\n' +
    'Where to go for help (H2)\n' +
    '- If you have any technical questions, contact the ' +
      '[Boise State Help Desk](' + HELP_URL + ').\n' +
    '- For questions about content or expectations, email the professor using a ' +
      'title something like, "' + courseCode + ': Help with ' + activityTitle + '".';
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

  // Fetch real titles for this module's reading/video links (best-effort, one
  // parallel batch). Missing titles fall back to the raw URL text downstream.
  if (moduleContext) enrichMediaTitles_([moduleContext]);

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

    // Assignments get a standardized "Where to go for help" section appended
    // deterministically (never AI-authored) so wording and links are exact.
    // Strip any code fence first so the appended block survives insertion.
    if (isAssignmentToolType4_(params.toolType)) {
      aiText = stripCodeFence4_(aiText) +
               assignmentHelpBlock4_(extractCourseCode_(), params.activityTitle);
    }

    // Resolve the design-map link tokens the AI may have embedded (VIDEO1, …).
    // Scope to the activity's kind so tokens line up with the (scoped) prompt.
    var scopedCtx = scopeModuleContextForActivity_(params.moduleContext, params.activityTitle);
    var media = buildMediaTokens4_(scopedCtx ? scopedCtx.mediaLinks : null);

    // Remove placeholder and insert formatted content at the same position
    var insertIdx = devBody.getChildIndex(placeholder);
    placeholder.removeFromParent();
    insertFormattedText4(devBody, insertIdx, aiText, indent, media.linkMap);

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
  // Scope the module context to this activity's kind (readings-only / videos-
  // only / all) so a "Readings" activity never sees the module's videos, etc.
  var activityKind  = classifyActivityKind_(params.activityTitle);
  var moduleContext = params.moduleContext
    ? scopeModuleContextForActivity_(params.moduleContext, params.activityTitle)
    : null;
  var media = buildMediaTokens4_(moduleContext ? moduleContext.mediaLinks : null);

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

  // Available-links block — real URLs offered as reference-by-TOKEN.
  var mediaBlock = media.hasLinks
    ? 'AVAILABLE LINKS — the course design map provides these readings/videos with ' +
      'real URLs. Reference each by its TOKEN, never by retyping the URL:\n' +
      media.promptLines.join('\n') + '\n\n'
    : '';

  // Link-embedding rule, tailored to whether any links are available.
  var linkRule = media.hasLinks
    ? '- Embed each relevant reading/video link directly in its TITLE using markdown ' +
      'link syntax with the TOKEN as the target, e.g. "[Introduction to Vectors](VIDEO1)". ' +
      'When AVAILABLE LINKS supplies a title in quotes, use that exact title as the link ' +
      'text. Use the exact token from AVAILABLE LINKS; never write out the raw URL and never ' +
      'invent placeholders like "[Link to video 1 here]". If a link has no title, put it on ' +
      'its own line as "[Watch the video](VIDEO1)". Only use tokens listed above.\n'
    : '- Do NOT invent link placeholders such as "[Link to video 1 here]"; if no links are ' +
      'provided, simply omit them.\n';

  // Content-scope rule: reinforce the withheld context so the AI doesn't invent
  // the other type's section for a single-type activity.
  var scopeRule = (activityKind === 'reading')
    ? '- This activity covers READINGS ONLY — do not include a videos or ' +
      'multimedia section.\n'
    : (activityKind === 'video')
      ? '- This activity covers VIDEOS ONLY — do not include a readings section.\n'
      : '';

  // Help-section rule: assignments get a standardized help block appended in
  // code, so the AI must not author its own; other tool types keep the mention.
  var helpDeskRule = isAssignmentToolType4_(params.toolType)
    ? '- Do NOT write your own "Where to go for help", technical support, or Help Desk ' +
      'section — a standardized one is appended automatically.\n'
    : '- If you refer to the campus help desk, call it exactly "Boise State Help Desk".\n';

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
    mediaBlock +
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
    scopeRule +
    linkRule +
    helpDeskRule +
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
  // Fail FAST: a sustained free-tier 429 won't clear inside a few retries, and
  // each doomed attempt is itself a slow, large generation. Two attempts with a
  // 15s cap surfaces "rate limit — wait and retry" in ~30s instead of stalling
  // the sidebar for 10+ minutes on a call that ultimately fails anyway.
  var maxAttempts = 2;
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
      var waitMs    = Math.min(Math.max(backoffMs, serverMs), 15000); // cap at 15s (fail fast)
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
 * Strips a surrounding markdown code fence (```…```) the AI sometimes wraps its
 * reply in, returning the inner content. Idempotent: text with no fence is
 * returned unchanged. Callers append their own content AFTER calling this, so
 * appended blocks are never swallowed by a later fence strip.
 *
 * @param {string} text
 * @returns {string}
 */
function stripCodeFence4_(text) {
  var s = String(text == null ? '' : text);
  var m = s.match(/```(?:markdown|text)?\n?([\s\S]*?)\n?```/);
  return m ? m[1].trim() : s;
}

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
 * @param {Object} [linkMap] TOKEN → URL map for resolving [title](TOKEN) links
 */
function insertFormattedText4(body, insertIdx, aiText, indent, linkMap) {
  indent = indent || INDENT_4;

  // Strip markdown code fences if the AI includes them despite instructions.
  // (Callers that append content after the AI text strip first, so this is a
  // no-op there and their appended block is never discarded.)
  aiText = stripCodeFence4_(aiText);

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
      setParagraphText4(li, token.text, BLACK, false, linkMap);

    } else if (token.type === 'rule') {
      body.insertHorizontalRule(insertIdx);

    } else if (token.type === 'heading') {
      var segs        = parseInlineSegments4(token.text);
      var cleanHeading = segs.map(function(s) { return s.text; }).join('');
      var fullText    = cleanHeading + ' ' + token.marker;

      var hPara = body.insertParagraph(insertIdx, fullText);
      hPara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      hPara.setIndentStart(indent);

      var pt = hPara.editAsText();
      _fmt(pt, { font: FONT, size: 11, bold: false, italic: false, color: BLACK });

      // Embed any [title](TOKEN|url) links found within the heading text.
      var hpos = 0;
      for (var hs = 0; hs < segs.length; hs++) {
        var hlen = segs[hs].text.length;
        if (hlen > 0 && segs[hs].link) {
          var hurl = resolveLinkTarget4_(segs[hs].link, linkMap);
          if (hurl) pt.setLinkUrl(hpos, hpos + hlen - 1, hurl);
        }
        hpos += hlen;
      }

      // The (HX) marker's red+bold AND the heading-text bold are BOTH applied in
      // the deferred applyHeadingBold4 pass below — inline character formatting
      // on freshly inserted text can be silently dropped before the write batch
      // flushes, which left markers black on large/link-heavy activities.

    } else {
      // Normal paragraph
      var para = body.insertParagraph(insertIdx, '');
      para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      para.setIndentStart(indent);
      setParagraphText4(para, token.text, BLACK, false, linkMap);
    }
  }

  // Post-processing pass: bold heading text after all elements are inserted.
  // Done here (not inline) so the GAS batch has fully flushed before we apply bold.
  applyHeadingBold4(body, insertIdx, tokens.length);

  // Normalize + hyperlink any "help desk" mention in the inserted text.
  linkifyHelpDesk4_(body, insertIdx, tokens.length);

  // Safety net: attach real links to any leftover "[Link to video N here]"
  // placeholders the AI emitted despite the token rule, and auto-link bare URLs.
  linkifyMediaFallback4_(body, insertIdx, tokens.length, linkMap);
}

// ── HELP DESK LINKIFIER ───────────────────────────────────────────────

/**
 * In the just-inserted paragraphs, normalizes any "help desk" mention to the
 * canonical "Boise State Help Desk" and hyperlinks it to OIT assistance. The AI
 * is also prompted to use that exact wording; this pass makes it reliable and
 * adds the link (the formatter itself produces no hyperlinks). Shared by the
 * Create Model Module (AI) and Deploy Activity Directions (AI) tools.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {number} startIdx   index where insertion began
 * @param {number} count      number of inserted elements to scan
 */
function linkifyHelpDesk4_(body, startIdx, count) {
  var URL    = 'https://www.boisestate.edu/oit/assistance/';
  var PHRASE = 'Boise State Help Desk';
  var end    = Math.min(startIdx + count, body.getNumChildren());

  for (var i = startIdx; i < end; i++) {
    var child = body.getChild(i);
    var type  = child.getType();
    if (type !== DocumentApp.ElementType.PARAGRAPH &&
        type !== DocumentApp.ElementType.LIST_ITEM) continue;

    try {
      // Cheap in-memory guard: skip the replaceText round-trips (2 per paragraph)
      // unless this paragraph actually mentions a help desk — most don't.
      var textEl = child.editAsText();
      var s      = textEl.getText();
      if (s.toLowerCase().indexOf('help desk') === -1) continue;

      // Collapse any existing "Boise State Help Desk" back to "Help Desk", then
      // promote every "help desk" (any case) to the canonical phrase — so the
      // result carries exactly one "Boise State" prefix regardless of AI casing.
      child.replaceText('(?i)boise state\\s+help desk', 'Help Desk');
      child.replaceText('(?i)help desk', PHRASE);

      textEl = child.editAsText();
      s      = textEl.getText();
      var from   = 0, idx;
      while ((idx = s.indexOf(PHRASE, from)) !== -1) {
        textEl.setLinkUrl(idx, idx + PHRASE.length - 1, URL);
        from = idx + PHRASE.length;
      }
    } catch (e) {
      // Linking is cosmetic — never let it fail the whole insertion.
      Logger.log('linkifyHelpDesk4_: skipped a paragraph (' + ((e && e.message) || e) + ')');
    }
  }
}

// ── HEADING STYLE POST-PROCESSOR ──────────────────────────────────────

/**
 * Scans inserted paragraphs and applies heading styling: bolds the heading
 * text (everything before the marker) and colours the "(HX)" marker red + bold.
 *
 * Runs after ALL elements are inserted so the GAS write batch has flushed —
 * inline character formatting on freshly inserted text can be silently dropped,
 * which previously left markers black on large / link-heavy activities. Applying
 * both the text bold and the marker colour here (not inline) makes them reliable
 * regardless of how big the activity is.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {number} startIdx   index where insertion began
 * @param {number} tokenCount number of tokens that were inserted
 */
function applyHeadingBold4(body, startIdx, tokenCount) {
  var headingRe  = /^(.*?)\s*(\(H[2-4]\))\s*$/i;
  var end        = startIdx + tokenCount;
  var childCount = body.getNumChildren();   // cache once — avoid a server round-trip per iteration
  for (var i = startIdx; i < end && i < childCount; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para = child.asParagraph();
    var text = para.getText();
    var m    = text.match(headingRe);
    if (!m) continue;

    var pt = para.editAsText();

    // Bold the heading text (everything before the marker).
    if (m[1] && m[1].length > 0) {
      pt.setBold(0, m[1].length - 1, true);
    }

    // Colour the (HX) marker red + bold. The marker sits at the end of the line
    // (ignoring any trailing whitespace), so its span is the last m[2].length
    // characters of the trimmed text.
    var trimmed     = text.replace(/\s+$/, '');
    var markerStart = trimmed.length - m[2].length;
    var markerEnd   = trimmed.length - 1;
    if (markerStart >= 0 && markerStart <= markerEnd) {
      pt.setForegroundColor(markerStart, markerEnd, RED);
      pt.setBold(markerStart, markerEnd, true);
    }
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
 * @param {Object}  [linkMap] TOKEN → URL map for resolving [title](TOKEN) links
 */
function setParagraphText4(el, rawText, color, bold, linkMap) {
  var segs      = parseInlineSegments4(rawText);
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
    if (len > 0) {
      if (segs[i].bold) pt.setBold(pos, pos + len - 1, true);
      if (segs[i].link) {
        var url = resolveLinkTarget4_(segs[i].link, linkMap);
        if (url) pt.setLinkUrl(pos, pos + len - 1, url);
      }
    }
    pos += len;
  }
}

/**
 * Splits inline text into styled segments, honoring both **bold** markers and
 * [text](target) markdown links. A link's visible text may itself contain bold.
 *
 * @param {string} raw
 * @returns {Array<{text: string, bold: boolean, link: string|null}>}
 *   link is the raw target (a media TOKEN or a URL), or null.
 */
function parseInlineSegments4(raw) {
  raw = String(raw == null ? '' : raw);
  var segments = [];
  var linkRe   = /\[([^\]]+)\]\(([^)]+)\)/g;
  var lastEnd  = 0;
  var match;

  while ((match = linkRe.exec(raw)) !== null) {
    if (match.index > lastEnd) {
      pushBoldSegments4_(segments, raw.slice(lastEnd, match.index), null);
    }
    pushBoldSegments4_(segments, match[1], match[2]);
    lastEnd = match.index + match[0].length;
  }
  if (lastEnd < raw.length) {
    pushBoldSegments4_(segments, raw.slice(lastEnd), null);
  }
  if (segments.length === 0) segments.push({ text: raw, bold: false, link: null });
  return segments;
}

/**
 * Splits `text` on **bold** markers and appends the resulting segments to
 * `segments`, each carrying the given link target (or null). Helper for
 * parseInlineSegments4.
 */
function pushBoldSegments4_(segments, text, link) {
  var re      = /\*\*(.+?)\*\*/g;
  var lastEnd = 0;
  var m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > lastEnd) {
      segments.push({ text: text.slice(lastEnd, m.index), bold: false, link: link });
    }
    segments.push({ text: m[1], bold: true, link: link });
    lastEnd = m.index + m[0].length;
  }
  if (lastEnd < text.length) {
    segments.push({ text: text.slice(lastEnd), bold: false, link: link });
  }
}

// ── MEDIA-LINK FALLBACK ───────────────────────────────────────────────

/**
 * Safety net for AI output that ignored the token convention. In the just-
 * inserted range it (1) swaps "[Link to video N here]" / "[Link to reading N
 * here]" placeholders for the matching real URL, and (2) auto-links any bare
 * http(s) URL that isn't already a hyperlink. No-ops when linkMap is empty.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {number} startIdx   index where insertion began
 * @param {number} count      number of inserted elements to scan
 * @param {Object} [linkMap]  TOKEN → URL map
 */
function linkifyMediaFallback4_(body, startIdx, count, linkMap) {
  var end = Math.min(startIdx + count, body.getNumChildren());

  for (var i = startIdx; i < end; i++) {
    var child = body.getChild(i);
    var type  = child.getType();
    if (type !== DocumentApp.ElementType.PARAGRAPH &&
        type !== DocumentApp.ElementType.LIST_ITEM) continue;

    try {
      var el = child.editAsText();
      var s  = el.getText();

      // 1) Swap explicit "[Link to <kind> N here]" placeholders for the real URL
      //    — only when a placeholder is actually present (rare), so the common
      //    path skips the replaceText round-trips entirely.
      if (linkMap && s.toLowerCase().indexOf('[link to') !== -1) {
        replaceMediaPlaceholders4_(child, 'video',   linkMap);
        replaceMediaPlaceholders4_(child, 'reading', linkMap);
        el = child.editAsText();
        s  = el.getText();
      }

      // 2) Auto-link any bare URL still sitting as plain text.
      var re = /https?:\/\/[^\s\]\)]+/g;
      var m;
      while ((m = re.exec(s)) !== null) {
        var url = m[0].replace(/[.,;]+$/, '');
        var st  = m.index;
        var en  = st + url.length - 1;
        if (!el.getLinkUrl(st)) el.setLinkUrl(st, en, url);
      }
    } catch (e) {
      // Linking is cosmetic — never let it fail the whole insertion.
      Logger.log('linkifyMediaFallback4_: skipped a paragraph (' + ((e && e.message) || e) + ')');
    }
  }
}

/**
 * Replaces "[Link to <kind> N here]" placeholders (any case) in one element with
 * the real URL from linkMap, for as many numbered tokens as exist for that kind.
 *
 * @param {GoogleAppsScript.Document.Element} child
 * @param {string} kind    'video' | 'reading'
 * @param {Object} linkMap TOKEN → URL map
 */
function replaceMediaPlaceholders4_(child, kind, linkMap) {
  var n = 1;
  while (linkMap[kind.toUpperCase() + n]) {
    var pattern = '(?i)\\[\\s*link to ' + kind + '\\s+' + n + '\\s+here\\s*\\]';
    child.replaceText(pattern, linkMap[kind.toUpperCase() + n]);
    n++;
  }
}
