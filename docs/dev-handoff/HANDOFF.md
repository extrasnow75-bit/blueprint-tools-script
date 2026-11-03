# Blueprint Tools: development handoff

Written 2026-11-02 at the end of a long design conversation in Claude Chat. Everything below was decided or learned there.

**Status (updated 2026-11-02, Claude Code session):** Tool 1 (Link Project Documents) is BUILT but UNTESTED: `Code10.gs`, `Sidebar10.html`, menu item in `Code.gs`, and the throwaway `Spikes.gs`. The Drive scope was added to `appsscript.json` with Greg's OK (see 3.6). Tools 2 to 5 are not started. Two new tool ideas were added this session (sections 11 and 12), and one idea is parked (section 13). Everything else below is the original spec.

Read this whole file, then `docs/dev-handoff/mockups/*.png` (sidebar mockups), then start at "Build order".

## 0. Who you are working for

Greg is an instructional designer and UX/UI designer. He is not a trained coder; he "vibe codes." The organization (Boise State eCampus) runs on Google Workspace, so everything is Google Apps Script.

- He wants workflow automation and efficiency.
- He does not want flattery. Do not compliment his questions or ideas.
- Explain things in plain language. Say what to click or test, not just what changed.
- He cannot always run scripts (he was driving when part of this was discussed). Do not assume he can test immediately; batch questions and tests.
- Ask before building when a wrong guess is expensive. Otherwise decide, state what you decided, and move on.

## 1. The project

Repo: `extrasnow75-bit/blueprint-tools-script` (GitHub). A Google Apps Script project bound to course "blueprint" Google Docs. It adds a **🎓 Blueprint Tools** menu (see `onOpen()` in `Code.gs`).

Existing files: `Code.gs`, `Code2.gs`, `Code3.gs`, `Code5.gs`, `Code8.gs`, `Code9.gs` and matching `Sidebar*.html`. About 5,500 lines. Work is on branch **`link-project-documents`**, not `master`. Do not merge to master without Greg asking.

### Conventions you must follow

- **File naming:** new files continue the numbering: `Code10.gs` / `Sidebar10.html`, `Code11.gs` / `Sidebar11.html`, and so on. (Greg stated this.)
- **Flat shared namespace.** Every `.gs` file shares one global scope. A function or constant defined twice does not error; the later one silently wins. So:
  - Give every new function and constant the file-number suffix (existing examples: `showDesignMapSidebar9`, `classifyHeading7_`, `normTitle9_`). Private helpers end with `_`.
  - Before defining anything, grep the whole repo for the name.
  - Reuse existing helpers rather than copying them (list in section 6).
- **File header style:** see the top of `Code9.gs`. It has a title, a short description, a `Last updated on YYYY-MM-DD at HH:MM MDT` line, and notes on shared helpers. Match it and keep the date current.
- **Sidebars:** 320px wide, Arial 13px, blue `#0033a0` headings, 44px minimum button height. The mockups show the style.
- **Scopes:** `appsscript.json` currently has `documents`, `script.container.ui`, `script.external_request`, plus the Docs v1 advanced service. **Any new scope changes the authorization prompt for every user of the script.** List exactly what you want to add and why, and get Greg's OK before editing `appsscript.json`.
- **Never overwrite user content silently.** Existing tools follow a "never-overwrite" contract. New tools show a preview and apply only what the user confirms.
- Commit messages and PR text: follow the attribution lines Claude Code gives you for the session.

## 2. Blueprint structure (what the code reads)

Three tabs: **Dashboard**, **Design**, **Development**.

- Development tab: `H2` = module; `H3` = "Due by … Mountain Time" headers; `H4` = activities (for example `1.03 Discussion: Topic 1`). Under each H4 is a tool line shaped like `<Tool> ⏺ Link to settings tab`. `TOOL_MARKER = '⏺'`. The tool name may be plain text or a dropdown smart chip. The "Link to settings tab" wording varies in real blueprints, and sometimes there is extra text after the link. Directions placeholder text: `Directions go here…`
- **Development tab order in the template (read 2026-11-02 from the template blueprint copy):** a "Development" banner, then single-cell header tables **Home (Course Navigation Menu Link)**, **Syllabus (Course Navigation Menu Link)** (with a "Select Syllabus Format" label), **Modules (Course Navigation Menu Link)**; then the H2 **Instructor Resources** (unpublished module; H3 items `COURSE ### Instructor Guide` and `COURSE ### Course Suggestions (if applicable)`, each followed by "Unpublished external link; load in a new tab"); then H2 **Course Resources** (published module; H3 `Page Title/Link Title` items, each a Page); then Module 1, Module 2, and so on. Untouched Home and Syllabus text is the placeholder `[Refer to the Template Blueprint Customization by Program document … to populate this section.]`.
- Other tabs in the template: Dashboard, Design, **Creative Brief** (sections "Creative Brief", "Media Deliverables", "Helpful Media Resources"), and **Media Notes** (5 tabs total). Dashboard title placeholder: `COURSE ###: Title`. Design tab: "Add catalog description here" is a bullet under **Course Information**, right after a line linking the undergrad and grad catalogs.
- The Media Deliverables table (template) has **eight** columns: Media Type, Title, Duration, Notes, Final Link (Embed in Canvas), Link Added to Dev Tab (IDC), Media Status, **Vetting Status**. Template rows are "Banner + Course Card" and six "Select Type" rows, all with empty Title/Final Link and status "Not Started". Untouched rows must never be flagged. In the template's text export, the "Link Added to Dev Tab" cell shows up as an **empty list item**, which fits a checklist checkbox (not a chip).
- Tool names normalize through `normalizeToolName(raw)` in `Code.gs`. `External Assignment` (a type-in variant) already maps to `Assignment`. `Assignment (Not Graded)` exists and uses the normal Assignment tab of the settings table template.
- Color Key (the document `11b03xRjWkJ6MvQvtUhg6CaRGmwjNAE478HQkzLYNZ3s`): light green 3 = ready for QA, light purple 3 = approved by QA, yellow = revised, strikethrough = built in Canvas. **Read the exact hex values from the Color Key or a test blueprint with `getBackgroundColor`; do not rely on memory.** The mockups use `#b7e1cd` (green) and `#d9d2e9` (purple) as stand-ins.
- Dashboard: has a Team Information chip column and a "Course Documents" column with labels (Settings Tables, Syllabus, Instructor Guide, Course Suggestions, Schedule of Deliverables, Google Folder, and others). Course length appears as text such as `Course Length: 7 Weeks` in the 505 copy (the template may use a chip, such as "15 Weeks"). Greg will add a **Credits dropdown chip** to a test blueprint (numeric option values) and ask a colleague to add it to the template.
- **Quickbase integration was dropped** (the API token is too risky). Never propose it again.

## 3. The tools, in build order

Menu order (final, with a separator between every item):

1. Link Project Documents
2. Add Activity Titles, Tools, Due Date Headers, & Times *(existing)*
3. Design Map → Dev Tab *(existing)*
4. Create Model Module *(existing)*
5. Deploy Activity Directions *(existing)*
6. Time Estimator *(existing)*
7. Draft Settings Tables
8. Deploy Settings Tables
9. Specialty Tool: Add Module Dates &/or Holiday Modules *(existing)*
10. Development Status Check
11. Blueprint Tools Version Check (second from the bottom, directly above KB Article)
12. KB Article *(existing)*

Build order: **10 → 11 → 12 → 13 → 14**, one branch per tool after the first (for example `deploy-settings-tables`). Greg preferred building the settings-table tools before the Status Check. The Version Check (`Code14.gs` / `Sidebar14.html`) comes last.

### 3.1 Link Project Documents (`Code10.gs` / `Sidebar10.html`)

Mockup: `mockups/LinkDocs.png` (updated 2026-11-02 to show Dept / Number / Name, the catalog description with the name-mismatch warning, the six-document preview with checkboxes, and the End-of-Project Report). **Greg has not yet reviewed this updated version.** The sample text is illustrative: the catalog title is invented, the description is a placeholder, and the "APPAI 507" → "APPAI 505" change is only there to show replace-old-title mode.

**A. Link documents**
- Find the blueprint's parent Drive folder (almost always the project folder). List its files.
- Match files to the Dashboard "Course Documents" labels. Names usually follow a pattern but vary.
- The user confirms each match with a dropdown, and can paste a URL for anything missing.
- Add links on the Dashboard. Also add the Instructor Guide and Course Suggestions links in the Development tab's Instructor Resources.
- Course Suggestions (and the End-of-Project Report) may not exist; skip with a notice and let the user paste a URL.
- Greg always has edit access to the project documents, but not to the template files.

**B. Fill the course identity into project documents.** Sidebar fields: **Dept**, **Course number**, **Course name**, with a live preview. Per-document preview with a checkbox per document, shown before any change. Documents (six):

| Document | What to replace |
|---|---|
| Blueprint | Dashboard title `COURSE ###: Title`; also the Development tab's Instructor Resources headings `COURSE ### Instructor Guide` and `COURSE ### Course Suggestions (if applicable)` |
| Agenda / Schedule of Deliverables | title line |
| Course Suggestions | `Course ###: Course Title` (H2) |
| Instructor Guide | title line; `Welcome to Course Name.` in the sample announcement (separate checkbox, on by default) |
| Syllabus | `Syllabus: Course ###` |
| End-of-Project Report | `COURSE XXX` → dept + number only (no name); case-sensitive; **three tabs**, so search every tab (the Media Summary tab has no course reference) |

Rules Greg confirmed: in `Course ###` the word "Course" stands for the department code, so it becomes e.g. `CS 321`. The end-of-project report gets dept + number only.

**Re-run / replace-old-title mode.** Finished projects are already filled in, and titles are inconsistent. Real examples from APPAI 505:

- Dashboard and Course Suggestions: `APPAI 505: Mathematical Foundations for Artificial Intelligence`
- Syllabus: `... (AI) Syllabus`
- Instructor Guide: `APPAI 505- Mathematical Foundations for Artificial Intelligence (Incomplete)` (hyphen, no colon, trailing note)

So exact-string replacement does not work. Find the title line by position and pattern (a short heading paragraph starting with a department code and number), show the match in the preview, and **replace only the identifier part**; leave extras like "(Incomplete)" or "Syllabus" alone. (Proposed; Greg had not confirmed this yet. Ask.) Any other occurrence (for example link text such as "CS 321 Agreement" in the CS 321 Agenda) must appear as an **unchecked** item in the preview, never replaced by default.

Shared tokens proposed for templates: `{{DEPT}}`, `{{COURSE_NUMBER}}`, `{{COURSE_NAME}}`, with a "COURSE ###" style fallback and the existing placeholder patterns above. Greg is asking colleagues to add tokens to templates; until then, match the existing placeholders.

**C. Catalog description.** Fetch the Boise State catalog description for the course and place it in the blueprint's **Design tab**, in the "Add catalog description here" bullet under "Course Information". Preview before writing.

- Undergraduate catalog: `https://www.boisestate.edu/registrar-catalog/#/home`; graduate catalog: `https://www.boisestate.edu/graduatecatalog/#/home`.
- **Rule: course number 500 and above uses the graduate catalog; below 500, undergraduate.** Show which catalog was used with a "check the other catalog" link (some courses appear in both). If it is in neither, say so and allow pasting.
- **Unresolved and unverified.** Both catalogs are hash-routed single-page apps. `UrlFetchApp` cannot run JavaScript. I could not confirm a machine-readable endpoint (my fetch tool was blocked). **First step: a spike.** Open a course page in a browser, check the Network tab for the request that returns the course text, and see whether it can be called with `UrlFetchApp` (the scope is already present; `Code8.gs` already fetches the registrar calendar this way). If there is no usable endpoint, fall back to: sidebar builds the right catalog search link from Dept and Number, and the user pastes the description; the script still places it with a preview.
- Compare the catalog title to the typed course name. If it differs, show both side by side with a warning but still allow insertion. Normalize before comparing: ignore case, punctuation, "&" vs "and", and text in parentheses (the Syllabus title has "(AI)").
- Record the catalog year or retrieval date in the preview. Older catalogs probably exist only as archived PDFs (ScholarWorks), not live pages; do not try to fetch older versions.

**Implementation notes (added after the build).**
- Files: `Code10.gs` (server), `Sidebar10.html` (UI), menu item `showLinkDocsSidebar10` first in `onOpen()`. Every place the handoff left a question open is marked `DEFAULT:` in the code; grep for it.
- Flow: fill in Dept / Number / Name, press **Preview all changes**, review, press **Apply**. Apply stays disabled until a preview has run on the current inputs. Apply re-scans every document and applies only ticked items.
- Dashboard "Course Documents" labels are plain paragraphs in the template; the link goes ON the existing label text (no new text). Google Folder links the parent folder.
- Instructor Resources has two forms: template (heading `COURSE ### Instructor Guide`, then "Unpublished external link...") and finished (heading plus a body sentence with a linked phrase). Default: if a linked phrase exists in the body text under the heading, that is the link spot; otherwise the heading text is linked. Existing links are never replaced unless "Replace links" is ticked.
- Identity: placeholder lines (`COURSE ###: Title`, `Course ###: Course Title`, `Syllabus: Course ###`, `COURSE XXX`) get the new identity; lines that already carry a code get only the code swapped. The Agenda and Instructor Guide placeholders are still unconfirmed, so only those two shapes are matched.
- Catalog: found 2026-11-02 in Greg's browser Network tab. The catalogs run on Kuali: `GET https://boisestate.kuali.co/api/v1/catalog/search/<catalogId>?q=cs%20321&limit=20` returns a JSON array (fields: code, subjectCode.name, number, title, description, pid, type). No sign-in needed in the browser. `fetchCatalogCourse10` uses it, with a button in the sidebar; the user can still paste. Still unverified: discovering the current catalog IDs (`/api/v1/catalog/public/catalogs`, run `spikeCatalogSpike10`) and the graduate catalog's ID. Prerequisites are not fetched yet (a separate course-detail request would be needed). The tool places the description only where the "Add catalog description here" bullet still exists.
- Schedule of Deliverables: Greg's existing project setup tool copies a template (A to E) into the project folder and puts the project name in the FILE TITLE only; the body is left unfilled. So Link Project Documents does not pick or copy templates and never opens the template folder. It fills the body of whatever copy it finds: A and C (Project Information table cells beside "Course Prefix/Number:" and "Course Title:"), B (title line `DEPT ###: Course Title`), E (title `DEPT ###: [Course Name]` plus `DEPT ### ...` link labels, handled as "Leftover placeholders", checked by default). D has no course fields and reports "No title line recognized". Unconfirmed: whether the `<<link Blueprint here>>` placeholders in A and the E link labels should also be linked to files.
- Syllabus (Greg, 2026-11-02): keep it mostly manual. The tool only fixes the heading: the program template has the dept already filled (`AMI ### Syllabus`), the faculty-owned template has `DEPT ### Syllabus`. The subtitle line `Course Title` is left alone for now. A .docx syllabus cannot be edited (link only).
- Catalog source link (Greg, 2026-11-02): wherever the catalog description is written (Design tab bullet, syllabus Course Description), it ends with a linked entry such as "(2026-2027 Undergraduate Catalog: CS321 - Data Structures)" (year, level, course code and title). The link goes to the course's own catalog page, `.../#/courses/<pid>?q=...&bc=true&bcCurrent=<title>&bcItemType=courses`, mirroring the address the catalog itself produced. Only when the description came from the lookup. Syllabus: the description is inserted as a new black paragraph above the purple "Insert Boise State catalog description" instructions, which are left in place (default; replacing them would delete text). Unchecking is per document in the preview.
- No catalog description (Greg, 2026-11-02): new courses in programs that have not launched may not be in the catalog yet. The lookup then returns a plain-English message with a timestamp ("Checked the graduate catalog for ... on <date and time>. The course is not listed yet..."), writes nothing, and the Design tab and syllabus placeholders stay untouched. A listed course with an empty description is treated the same way. Network errors get a different message ("could not be reached ... nothing was checked").
- Schedule links (Greg, 2026-11-02: yes): Template A's `<<link Blueprint here>>` placeholders become "<CODE> Blueprint" linked to the blueprint; Template E's `DEPT ### Blueprint / Syllabus / Instructor Guide / Google Drive folder` lines are linked. Other labels and Templates B, C, D are untouched. One checkbox in section 4.
- Syllabus subtitle (Greg): fill the course name only when the syllabus has a standalone `Course Title` / `[Course Name]` placeholder paragraph; templates vary by department, so nothing else is guessed. Program and master syllabi follow departmental templates; faculty-owned courses share one template.
- Only fill what has a placeholder (Greg): no Dashboard link for the Student-Facing Content folder. Greg will not add a test row to the live multi-semester sheet.
- Known gaps: Syllabus titles of the form `... (AI) Syllabus` with no code are reported as "No title line recognized". The `^` anchor in `Body.findText` (used to find Dashboard labels) is unverified; `spikeLabelsSpike10` checks it.
- Test results from Greg: pending.

### 3.2 Deploy Settings Tables (`Code11.gs` / `Sidebar11.html`)

Mockups: `DeployA.png` (flat list), `DeployB.png` (by module). Greg liked the grouped/by-module direction; confirm which one before building.

- Find the settings table (Google Sheet) through the Dashboard's Settings Tables link, or let the user choose a file.
- Match settings-table **tabs** to Development-tab **activities**. Many activities can share one tab (many-to-one). Dropdown override per activity. Preview before any change.
- Write a link onto the tool line (the "Link to settings tab" phrase). Page activities have no settings tab; skip them.
- Never match the Course Settings tab or the template tabs.
- Option to replace links that already exist.

**Built (2026-11-03, branch `deploy-settings-tables`; untested in Apps Script).** `Code11.gs` / `Sidebar11.html`, menu item after Time Estimator. DEFAULT: by-module layout (B). Finds the settings table from the Dashboard's "Settings Tables" link (or the user pastes one), reads tab names, previews a match for every activity, user fixes picks, "Add links" writes the link onto the words "Link to settings tab" only (`.../edit#gid=<tab id>`). Matching (many activities to one tab): title equals tab name or the part before the colon = exact; tab name opening the title, shared keyword, or a tab named for the Canvas tool = suggested (pre-selected, yellow); otherwise no match (dropdown). Course Settings, anything with "Template" in the name, and hidden tabs are never matched. Page activities and tool lines without the phrase are skipped. Existing links are kept unless "Replace" is ticked. Sheets scope: Greg approved the full `spreadsheets` scope on 2026-11-03 (one prompt now; Draft needs write). Added to `appsscript.json` in its own commit.

### 3.3 Draft Settings Tables (`Code12.gs` / `Sidebar12.html`)

Mockup: `DraftTabs.png`. Also `A1Title.png`.

- Purpose: quick draft tabs in the settings table by duplicating the template tabs (Assignment, Discussion, Classic Quiz, New Quiz). Everything else (editing tab contents) stays manual.
- Group activities by similar titles (strip numbers and topic labels). Modes: **one per kind of activity (default)**, one per Canvas tool, one per activity. Never default to ~30 tabs.
- Section is called **"Settings Table Tab Names"** (not "tab names"; that confuses people with blueprint tabs). Names are editable, up to 100 characters, and unique.
- "Add another tab (exam, evaluation, survey…)" button.
- **Cell A1 stays exactly as in the template** (Greg chose option A). Do not edit any cell contents.
- Copy whole tabs; never rebuild them. Conditional formatting depends on fixed cell positions.
- **The tool must never tick the "Check when settings are finalized:" box.** If a template tab arrives ticked, warn; do not untick it.
- A "Split off" control (to pull one activity out of a group) was proposed but not mocked up.

### 3.4 Development Status Check (`Code13.gs` / `Sidebar13.html`)

Mockups: `Main.png` (setup), `StatusResultsB.png` (results; Greg picked B), `StatusResultsA.png` (alternative, not chosen).

Counts **whole activities** (not paragraphs) by stage:

1. untouched template text 2. drafted, not highlighted 3. green = ready for QA 4. purple = approved by QA 5. yellow = revised 6. strikethrough = built in Canvas.

Mixed formatting inside one activity: any yellow makes it "Revised" (Greg's rule); otherwise count it at the lowest stage (my proposal; not explicitly confirmed).

Other checks: Home text present; Instructor Guide URL; Course Suggestions URL; a settings-table link on every non-Page tool line (**any link attribute counts**, not the wording; placeholder text without a link counts as missing; report this per activity next to the red tool name); module time vs. credit-hour guideline.

Time target per module = credits × 37.5 to credits × 45 hours ÷ weeks ÷ modules per week, rounded up to ¼ hour. Example: 3 credits over 7 weeks gives 16¼ – 19½ hrs. Reuse `runTimeEstimator(credits, weeks, modulesPerWeek)` from `Code3.gs`. Credits come from the Dashboard dropdown chip; weeks from Course Length.

**Settings-table finalized check (added late).** Row text `Check when settings are finalized:` followed by a checkbox, one per settings-table tab. Match on the shorter phrase "settings are finalized" and find the row by text, not position. Report "N of M tabs finalized", with each tab listed. Tabs with "Template" in the name start unchecked; every other tab (including Course Settings) starts checked; the user can toggle each one. A tab with no such row is listed as "no finalize box found", not counted as unfinished. **Read-only: no tool may tick this box.** Activity-level counting ("31 of 35 activities have finalized settings") is a possible later addition via the links Deploy writes.

**Syllabus section (added late; Greg's request).** The Development tab has a **Syllabus (Course Navigation Menu Link)** section before the modules. Sometimes it is filled out completely (the syllabus becomes a Canvas page); sometimes it holds a link to a separate syllabus document. **Apply the same stage rules as everything else:** untouched placeholder text, text present but not highlighted, green, purple, yellow (revised), struck through. Also report which format it is (content written in place vs. a link to a document). Proposed, not confirmed with Greg: (a) apply the same stage rules to the **Home** section; (b) report Home and Syllabus on their own lines and **do not count them in the "N of M activities" totals**, since they are not module activities; (c) for a linked syllabus, the "text present" check is a link in the section, and the stage still comes from its highlight.

**Media deliverables check (added late; Greg's screenshot and description).** Each blueprint has a **Creative Brief** tab (the blueprint has about 5 tabs: Dashboard, Design, Development, Creative Brief, and one more). It has a section called **"Media Deliverables"** with a table. Columns (from the template and the filled-in APPAI 507 report's Media Summary): Media Type | Title | Duration | Notes | **Final Link (Embed in Canvas)** | **Link Added to Dev Tab (IDC)** | **Media Status**. In Greg's screenshot, "Link Added to Dev Tab (IDC)" holds a **checkbox** (one row ticked, one not), "Media Status" holds a **dropdown chip** (for example "Completed"), and Notes holds date chips.

- Find the section by its heading text and the table after it. **Find the columns by header text, not position** ("Final Link", "Link Added to Dev Tab", "Media Status"; the headers have extra words and sometimes a link).
- **Headline result, shown prominently (a callout near the top of the results, above the activity counts):** media items whose **Media Status is Completed and that have a Final Link, but whose "Link Added to Dev Tab (IDC)" box is not ticked.** List each by media type and title, with a "Go to it" link to its row. If there are none, say so in one line.
- Secondary: counts of media items by Media Status value.
- **Read-only, like the settings-table box: never tick the "Link Added to Dev Tab" box.**
- **Unverified and the main risk:** whether the checkbox's ticked/unticked state can be read. It may be a Docs checklist item, a smart chip, or something else; a checked checklist item may only show as struck-through text, which would not exist in an empty cell. **Spike first (item 8 below).** If the state cannot be read, a fallback is to ask Greg to have the template's column changed to a dropdown chip (for example Yes / No), which the Docs API can read the same way as the Credits chip. Ask Greg before assuming.
- **Answered by Greg (2026-11-02):**
  - Media Status dropdown options are exactly: **Not Started, In progress, Under review, Completed, Cancelled.** "Completed" means the item is done. Treat **Cancelled** as out of scope (never flag a Cancelled item).
  - **Also flag rows that are ticked "Link Added to Dev Tab" but have no Final Link.** Greg says this would be unusual, so show it as a smaller secondary warning, not in the headline callout.
  - The tab is **always named exactly "Creative Brief."**
- The mockups (`Main.png`, `StatusResultsB.png`) do **not** show this check yet. Add a "Media deliverables (Creative Brief tab)" checkbox to the setup screen and the callout to the results screen.

**Report order (Greg's rule: roughly the order of the Development tab, Home first, media last; confirmed 2026-11-02, except where noted).** This replaces the order in the current `StatusResultsB` mockup, which has the totals at the very top.

1. Header (the report's top line, not a Development tab section; Greg approved it as is): date and time the check was run, number of activities counted, number of modules counted. It shows which version of the blueprint the report describes.
2. Home: whether home screen text has been added, and its stage (see the Syllabus note above).
3. Syllabus: format (written in place or linked document) and stage.
4. Instructor Resources: Instructor Guide URL present; Course Suggestions URL present.
5. **Course totals** (Greg asked for these before the modules): activities ready for the dev team out of the total; the six stage counts (untouched, drafted, green, purple, yellow, built); total settings table links present out of those expected. *Placement after Instructor Resources (which is the last item before the modules in the Development tab) and before the modules is my reading; Greg did not say whether totals go above Home. Ask if it matters.*
6. Modules in order (Module 1, 2, 3 …): time estimate vs. target; ready count with stage breakdown; expandable activity list with each activity's stage and, for non-Page tools, its settings table link status.
7. Settings table tabs: how many tabs are marked finalized out of those counted, each tab listed (Course Settings shown as "no finalize box found" if it has none).
8. Media deliverables (Creative Brief tab), last: a colored alert callout for Completed items with a Final Link that are not ticked as added to the Dev tab; a smaller warning for ticked items with no Final Link; counts by media status (Cancelled ignored). Open question put to Greg: also show a one-line pointer to this alert at the very top of the report when there is something to act on.

Output: always shown in the sidebar. Ask (on both the setup and results screens) whether to also save the report as a new blueprint tab, a new Google Doc in the blueprint's folder, or neither. Wording for the "ready" count is still Greg's call ("Ready for dev team" vs. "Highlighted for QA or later").

The "Status Check" settings-table reading needs a Sheets permission. For this tool read-only would be enough; Draft needs write (it creates tabs); Deploy only needs to read tab names. **Also decide the Drive scope for "save to a new Google Doc in the blueprint's folder"** (used by the Status Check and the Version Check). **Decide the Sheets scope for all three tools together** and list it for Greg before editing `appsscript.json`.

### 3.5 Blueprint Tools Version Check (`Code14.gs` / `Sidebar14.html`)

Added late. No mockup yet. Greg pastes code into Apps Script by hand, so this tells him which files to paste after an update.

**Format: a pop-up dialog (light box), not a sidebar** (Greg's choice). Nothing in it is editable. Use `showModalDialog` (or `showModelessDialog` if Greg wants to keep working in the document; ask) with a fixed width and height, and let the list scroll. Keep the `Sidebar14.html` file name so it follows Greg's `Code#.gs` / `Sidebar#.html` naming rule, even though it is a dialog.

**What the dialog shows (keep it this short; Greg said more detail is too much):**

1. **Out-of-date files:** file name, the version date in the document, and the date of the version on GitHub `master`.
2. **New files on GitHub that are missing from the document.**

Nothing else: no "what changed" notes, no per-file explanations. If everything is current, say so in one line.

**Export option (Greg asked for it):** a choice to save the results as **a new tab in the document** or **a new Google Doc in the same Drive folder as the document** (or neither). A new tab needs no new permission beyond `documents` (check that `Document.addTab` or the Docs API `addDocumentTab` request works; spike item 3). **Creating a Doc inside the document's folder very likely needs a broader Drive permission than this project has**: a new Doc is created in My Drive and moving it needs more than read-only metadata access. Not verified; test it, and if it needs a broader scope, list it for Greg before touching `appsscript.json`. This is the same question as the Status Check's identical export option (3.4).

**Design (proposed; confirm details with Greg):**

- Compare against the **`master`** branch only (Greg confirmed). The repo is public, so no token is needed, and never add one.
- The repo gets a manifest on `master` (for example `version-manifest.json`) listing every script file and its version date. Fetch it with `UrlFetchApp` from `https://raw.githubusercontent.com/extrasnow75-bit/blueprint-tools-script/master/version-manifest.json`. This uses the existing `script.external_request` scope; no new scope is needed. (Checked on 2026-11-02: the raw URL returns 200.)
- **A script cannot read its own `.gs` source.** So each `.gs` file carries a version stamp equal to its `Last updated on` header date, as a top-level `var` (not `const`; in the V8 runtime only `var` and function declarations become properties of the global object, so `globalThis[...]` lookups would miss a `const`; verify this in a spike). A `.gs` file is "missing" when its stamp is undefined.
- Sidebar `.html` files can be read with `HtmlService.createHtmlOutputFromFile(name).getContent()`. Put a version comment in each (for example `<!-- version: 2026-11-02 -->`) and read it. A file that cannot be opened is "missing."
- Comparison: document older than GitHub = out of date. Document **newer** than GitHub = show as "newer than GitHub" (probably a branch version), not an error.
- Whenever a commit changes a script file, update its stamp and the manifest in the same commit. Consider a small repo script that fails if a changed file's stamp and the manifest disagree, since a forgotten stamp makes the check wrongly say "current."
- Keep the manifest format backward compatible, so an old copy of this tool still lists new files as missing, as long as the tool itself was pasted.

**Known limits (do not hide them from Greg):**

- Files that exist in the document but not in the repo cannot be detected (listing a script's own files needs a broader permission; do not add it).
- A forgotten stamp bump is invisible.
- `appsscript.json` (permissions) cannot be read from inside the script, so a release that adds a scope cannot be flagged by this tool. Ask Greg whether to mention this in the sidebar.


### 3.6 Permissions summary (decide these together, before editing `appsscript.json`)

Current scopes: `documents`, `drive` (added 2026-11-02 for Link Project Documents; Greg chose the full Drive scope, option b, over the Advanced Drive service with metadata-only), `script.container.ui`, `script.external_request`, plus the Docs v1 advanced service. A new scope changes the authorization prompt for every user of the script. Get Greg's OK on the exact list first. This table is from the design discussion, not from testing; confirm each row in a spike.

| Tool | Needs | New scope? |
|---|---|---|
| Link Project Documents | List the blueprint's parent folder (read-only Drive metadata). Edit the other project documents (the existing `documents` scope should cover documents Greg can edit). Catalog fetch (`script.external_request`, already present). | Drive read-only metadata |
| Draft Settings Tables | Copy template tabs inside the settings table spreadsheet, so write access to Sheets. | Sheets (read and write) |
| Deploy Settings Tables | Read the settings table's tab names. Writes only to the blueprint. | Sheets (read-only is enough) |
| Development Status Check | Read settings-table tabs and the "Check when settings are finalized:" checkboxes. Save the report to a new blueprint tab (no new scope). Save it to a new Google Doc in the blueprint's folder (**see below**). | Sheets (read-only); Drive for the new-Doc option |
| Blueprint Tools Version Check | Fetch the manifest from GitHub (`script.external_request`, already present). Same export options as the Status Check. | Drive for the new-Doc option only |

**Shared decision, "save to a new Google Doc in the same folder":** the Status Check and the Version Check both offer it, so one scope decision covers both. A new Doc is created in My Drive and moving it into the blueprint's folder very likely needs a broader Drive permission than read-only metadata. Not verified. Test it. If it needs a broad Drive scope, tell Greg plainly and offer to drop the option and keep only "new tab in the blueprint" (which needs no new scope).

**Sheets:** pick the narrowest scope that works for each tool, but remember all tools share one authorization prompt, so the effective scope is the widest one. Draft Settings Tables needs write access.

## 4. Spikes to run first (throwaway code in a separate file)

Put these in something like `Spikes.gs` and delete them afterward. Greg must run them; he was not at a computer.

1. **Read a dropdown chip with the Docs API** (the Credits chip or tool-name chip). Expected: `ParagraphElement.dropdown` with `dropdownDefinitionId` and `selectedOptionId`; the definitions are in `documentTab.dropdownDefinitions` with `displayValue`. `DocumentApp` returns empty text for chips.
2. **List the blueprint's parent folder** using only the read-only Drive metadata scope.
3. **Create a report tab** in the blueprint and write text into it.
4. **Catalog endpoint** (see 3.1 C).
5. **Read settings-table checkboxes** (`getValue()` returns true/false) with the narrowest Sheets scope that works.
6. **Global lookup of stamps:** confirm a top-level `var` in one `.gs` file is readable from another via `globalThis` (needed by 3.5).
7. **Fetch the manifest** from raw.githubusercontent.com with `UrlFetchApp` inside a blueprint-bound script.
8. **Read the Creative Brief media table:** with the Docs API, fetch the Creative Brief tab of a real blueprint, find the Media Deliverables table, and print for each row the Final Link text, the Media Status dropdown's display value, and whatever distinguishes a ticked "Link Added to Dev Tab (IDC)" box from an unticked one (compare two rows: one ticked, one not). This decides whether the media check is feasible as designed.

## 5. Test files (Google Drive)

All owned by gregsnow@boisestate.edu and shared with Greg's extrasnow75 account. Claude Code needs a Google Drive connector, or Greg pastes content; if you cannot open them, say so.

**CS 321 folder** `1lMaGFtJJ9E2sVCIOVefBJjwmWTpZn0Km` ("Copy of CS 321 w/ Kathryn (Blueprint Tools Test)"). Undergraduate course. Greg planned to repair its Design tab because the faculty member's own scripts damaged the course pattern table and design map. Do not use it for Design-tab testing until he says so.

- Blueprint: `1VhLrk5t2l3IVX44bQ-6cxFbNjmTurp7SdChZv8O1Buw` ("Copy of partially completed CS 321 Blueprint")
- Settings Tables (Sheet): `1IPBT4z5An3Ki9AbiDqmPwn7tKi7mYlZY5CbSLgknYnM`
- Course Suggestions `1g-XLpDbpIfTKXPz20zH1bKIBMG4q7JyV3A4hA0jr5GM`
- Instructor Guide `1GiRD68MszmcEWo0e7Ljn_4_HE87KVMByVHvH1ZnFDgY`
- Syllabus `1J7eug1jcJIl1IxqMQPPMiv1klBfothLZzHcjH2hcL5w`
- Agenda & Schedule of Deliverables `1Zxoxn50HMZHQ3jL3nH3t_EEWuce9WfYx1EbPqhCywEA`
- A generated "Video List - Copy of partially completed CS 321 Blueprint …" sheet `1OnotcNztYCFU35bCm4be4rIrdtQ5_UzVJIW0vIjUkv4`. **Trap:** its name contains the blueprint's full title. A name match on "Blueprint" hits both files. Skip the bound blueprint by ID, filter by document type, and treat the Video List as unrecognized.
- No End-of-Project Report in this folder. All file names carry a "Copy of" prefix (ignore it when matching).

**APPAI 505 folder** `1pmJsy8Do--2tWuHPBDFTnqXb26oySoKV` ("Copy of APPAI 505 (for Blueprint Tools)"). Graduate course (505), well-filled Design tab. Finished before end-of-project reports existed.

- Blueprint `1h5YaRDE0MdR3AvgiOI8AvRdUKErPZ1yFs96lhtSFOAM`
- Course Suggestions `1HpQx4BWi5RQWKxsCs4O5cuilNWEGt5kr3yXNJGlbLSk`
- Agenda `1u898ECI_bNqNJRvOXzhKlPo-1idz4uhjknjsOCZhedE`
- Syllabus `1wNwrULr9YWnDZwfz5ypmCvzcHwYt_TgblGr0EOTwjLo`
- Instructor Guide `1iYVT8B21Bg-GoKSl-_OKVqGmN43Zk7Hay3aXyfvLVTc`
- **No settings table** and **no End-of-Project Report** in this folder (Greg was asked whether a settings table exists; unanswered). All five documents are already filled in, so this is the **replace-old-title** test case.

**Templates and examples** (Greg's staging folder "Blueprint Tools Stuff" is `1nUm-lWP5EdLnk-7ws3Qc0rEwDSAUpcxE`; not every file below lives in it)

- Template blueprint copy: `1-KtR2g67N4qBlo7wELkNaWnu1XKSSfEpezk2ChN3_Ho` ("Copy of Template Blueprint 6.0 for GAS Test")
- Settings table template: `1p5k4FUd2US257Om8FHywHFbdnYia6NfEMA6HjNVVqUI` (tabs: Assignment, Discussion, Classic Quiz, New Quiz templates, plus Course Settings)
- End-of-Project Report template copy: `1rtMwwAuLfLvtB7iKJYk-RGJVnd9mbg4qADzheELsRVQ` (still has `COURSE XXX`)
- Filled-in End-of-Project Report (APPAI 507, a different project): `1dTJt6Bb_qzUQOGAjy_Pom9riEPO4Jl0sVYd5g5NCAMI`
- Other blueprint for script tests: "Fourth Example Blueprint for Script Test" `1EzDtfO4tMSLEGDSzrCfso5VeUsfvqCjSuujCjlAQN1I`
- Blueprint Color Key `11b03xRjWkJ6MvQvtUhg6CaRGmwjNAE478HQkzLYNZ3s`

Greg said he would send **unfilled** copies of the Instructor Guide and Agenda (their titles are already filled in in these copies, so the real placeholder is not visible) and the Schedule of Deliverables template. Until then, do not guess their placeholders.

**Not accessible to the earlier session:** the "Template Activity Directions" document and the original (uncopied) CS 321 settings table.

## 6. Existing code to reuse

| Need | Where |
|---|---|
| Menu | `onOpen()` in `Code.gs` (add new items there; keep the separators) |
| Tabs list | `collectTabs(doc)` in `Code.gs` |
| Development tab body | `getDevelopmentTabBody(doc)` in `Code2.gs` |
| Tool name normalization | `normalizeToolName(raw)` in `Code.gs` |
| Doc ID from URL | `extractDocId` in `Code2.gs` |
| Directions placeholder | `findDirectionsPlaceholder` in `Code2.gs` |
| Time targets | `runTimeEstimator(credits, weeks, modulesPerWeek)` in `Code3.gs` |
| Docs API request pattern | `borderDueHeaders_` / `flattenDocsTabs_` in `Code.gs` |
| `UrlFetchApp` precedent | `fetchAcademicCalendar7` / `fetchPage7_` in `Code8.gs` |
| Title matching | `normTitle9_`, `keywordScore9_`, `matchActivities9_` in `Code9.gs` |
| Constants | `TOOL_MARKER`, `TOOL_SUFFIX`, `CANVAS_TOOL_OPTIONS` in `Code.gs` |

Confirm each signature by reading the code before calling it; this list was made from a summary, not a fresh read.

## 7. Open questions for Greg

1. **Answered:** Greg pastes code into the Apps Script editor himself. Do not set up `clasp` or run any push on his behalf. When you change files, tell him exactly which files to paste.
2. **Accepted as default (2026-11-02):** replace only the identifier in a filled-in title and keep extras like "(Incomplete)" and "Syllabus".
3. **Answered:** APPAI 505's Dashboard links a Settings Tables sheet, but it points to the original, not to a copy in the test folder.
4. Which Deploy layout: flat (A) or by module (B)? (Needed before tool 2.)
5. "Ready for dev team" wording vs. "Highlighted for QA or later".
6. Whether to also prefill End-of-Project Report fields from the Dashboard. Offered, not requested; keep the first version to dept + number.
7. Which Sheets scope (see 3.4 and 3.6). Now needed by tools 2, 3, 4 and the setup-data tool (section 11). Greg must approve before `appsscript.json` changes.
8. Mixed-formatting rule below "yellow wins" (my proposal: count at the lowest stage).
9. Results of the Link Project Documents tests and spikes (see section 10).
10. Setup-data tool (section 11) and ownership tool (section 12): the questions listed there.

## 8. Mockups

`docs/dev-handoff/mockups/` holds each mockup as `.png` (view these) and `.dc.html` (source; they render standalone in a browser but are design-canvas files, not shipping code). The canvas also lives as a private claude.ai artifact: https://claude.ai/artifact/9VaZVq4C24jP8o1hEt9WKH (only Greg can open it). Sample numbers in the mockups are invented (49 activities, 7 modules).

Updated this session: `Main.png` and `StatusResultsB.png` now include the settings-finalized check, and `LinkDocs.png` now shows the full project setup (title fields, catalog description, document preview). No mockup exists yet for the Version Check dialog; Greg said none is needed.

## 9. Greg's last stated decisions (quick list)

- Draft Settings Tables: leave A1 alone; no content edits; no ticking of the finalize box.
- Counting unit: whole activities; any yellow means Revised.
- Menu order and tool names as in section 3.
- Credits via a Dashboard dropdown chip; Course Length gives weeks.
- Graduate catalog for course numbers 500 and up.
- Catalog name mismatch: warn, but still allow insertion.
- Quickbase: dropped.

## 10. Pending from Greg (updated 2026-11-02)

- **Run and report:** the Link Project Documents test checklist and the three spikes in `Spikes.gs` (labels with the `^` anchor, folder listing through `DriveApp`, catalog endpoint; for the catalog, the URL of the request in the browser's Network tab that returns a course description).
- **Sheets scope decision** (see 3.6). Needed before tools 2, 3, 4 and section 11.
- **Setup-data tool:** the original Google Sheet (not the .xlsx copy) or its ID, and the answers listed in section 11.
- **Ownership tool:** answers listed in section 12.
- Unfilled copies of the Instructor Guide, Agenda and Schedule of Deliverables, so their placeholders can be matched.

## 11. Setup-data tool: read the project's row from the Multi-Semester List (idea, not started)

Greg shared an Excel copy of the list in his personal Drive (read-only for Claude; nothing from it is stored in this repo, and the file contains faculty names, so do not paste rows here). The real tool needs the original Google Sheet.

What the workbook looks like (columns and tabs only):
- One tab per semester (named like `Fall 2026`), plus history tabs, an IDC capacity tab and a small extra tab. About 900 rows by 35 columns each. The tab is the **development** semester, not the semester the course is taught. eCampus work runs on **cycles** (14-week development cycles; Greg: "Fall 2026 Cycle 1", "Fall 2026 Cycle 2"), and the cycle number is in the row's `Project Start Date` column. Use the word "cycle" in the UI. Cycle dates are NOT computed or written by this tool (Greg described the cycle timing two different ways, and the side table in the copy looked overlapping); copy the tab's season and year plus the cycle label exactly as written.
- Row 1 headers include: Course Title, Project Type, OPDP, Crs Dev Shell Creator, Status, Project Start Date (holds values such as "Cycle 1"), First taught (values such as `SP27 15wk`, `FA26 7W1`, `SU27 10wks`, `n/a`), Copy Date Overlap, FD Meeting Pref., College, Assigned IDC, IDC / QA / Build / MM / CAS (flags, not names), Dev Team Priority, Notes.
- The project key is column C, shaped `DEPT NUM-Faculty name`. Some rows hold several codes or several faculty (for example `ENGL 101/102-...`).
- Checked against the APPAI 505 Dashboard: `Next taught` equals column J, and `Cycle: Summer Cycle 1 (2026)` equals the tab name plus column I.
- CS 321 is in the Fall 2026 tab and APPAI 505 is in the Summer 2026 tab.

Candidate mapping to the Dashboard: Next taught = First taught; Cycle = tab name plus Project Start Date; Course Length (weeks) from the `7W1` / `7W2` / `15wk` / `10wks` / `14wk` suffix; Course title feeds Link Project Documents' Course name; Project Type and the person fields are chips.

People chips: the Docs API can insert a person chip only from an email address, not a name (not yet tested). Greg says the IDC is usually the blueprint's owner, so the script can read the owner's email from Drive (`File.getOwner()`; covered by the Drive scope). Do this BEFORE any ownership transfer (section 12). Faculty developer and other names stay manual unless a directory lookup is approved (new scope, would need Greg's OK).

Needs Greg's answers: which fields to scrape; where they are written (Dashboard, a preview in the sidebar, or both); what values besides "Cycle 1" / "Cycle 2" appear in the cycle column; a course deferred and listed in two tabs should ask, showing each tab and cycle (proposed, unconfirmed); the original sheet's ID and who has view access; the Sheets scope (read-only is enough). Proposed file: `Code15.gs` / `Sidebar15.html`, or folded into Link Project Documents. Ask Greg.

## 12. Ownership transfer tool (idea, not started)

Today the IDC moves ownership of the project folder and documents to shared eCampus accounts by hand at the end of a CDD cycle. Rules depend on the course type (Master, Program, Faculty). They come from an internal eCampus document in Greg's Drive; **do not copy that document's text or the account addresses into this public repo.** Keep the rules and addresses in a local config block that Greg fills in when he pastes the script (blank placeholders in the repo).

Design points:
- Roles that appear in the rules: a central development account, a vetting account, a maintenance account, plus the Program Coordinator/Chair and the Faculty Developer. Greg says files for the last two can be handled manually, so the tool skips and lists them.
- Course type probably comes from the Project Type column in section 11.
- `File.setOwner` / `Folder.setOwner` exist in `DriveApp` (full Drive scope already added). Untested inside the Boise State domain. A transfer of a folder moves the folder only, not its contents. Non-Google files (PDF, mp4) cannot be transferred this way. Large transfers can take a day. Only the current owner (or an admin) can transfer. Possible admin policy limits are unknown.
- It is hard to undo (the old owner stays an editor; mistakes go through a request process). So: list every file with its current owner and its target, skip what cannot move, change nothing until the user confirms, and keep it separate from Link Project Documents.
- New project folders: the media folder (with an inner accessible-media folder) is being renamed `DEPT ### Student-Facing Content`. Match it by name, accepting `Student-Facing Content`, `Student Facing Content` and the older `Media` / `Accessible Media` names. The Instructor Guide and Course Suggestions in a new project folder are Drive shortcuts to documents that sit beside the project folder, and the syllabus may be a .docx; the ownership tool has to resolve shortcuts to their targets and skip files it cannot transfer.
- Order: fill the IDC chip from the blueprint's owner first, transfer the blueprint last (the sidebar runs inside it).
- A spike is needed first: transfer one throwaway doc inside the Boise State domain. Greg runs it.
- Proposed file: `Code16.gs` / `Sidebar16.html`.

## 13. Parked: project work station notebook (back burner)

Greg's plan document describes a Gemini Notebook "project work station" per project. Findings: consumer notebooks have no API; a separate Google Cloud NotebookLM Enterprise API has `notebooks.create`, but whether Boise State's Gemini Notebook is that product, and whether the licence and a Cloud project are available, is unverified; notebooks are not Drive files, so "create it in a chosen folder" does not apply. A no-new-scope fallback is a "Notebook Sources" subfolder of Drive shortcuts to the project and template documents (needs a test of whether the app imports a shortcut's target). Not planned for now.

## 14. Next steps (proposed order)

1. Greg runs the Link Project Documents tests and spikes and reports back.
2. Fix whatever the tests find in `Code10.gs` / `Sidebar10.html`.
3. Greg decides the Sheets scope and Deploy layout; then build tool 2 (Deploy Settings Tables, `Code11`).
4. Tool 3 (Draft Settings Tables, `Code12`), then tool 4 (Development Status Check, `Code13`), then tool 5 (Version Check, `Code14`).
5. Setup-data tool (section 11) can start once Greg supplies the sheet and answers; it shares the Sheets scope decision and may come earlier than tool 3.
6. Ownership tool (section 12) after its spike.
