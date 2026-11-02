# Blueprint Tools (Google Apps Script)

Google Apps Script project bound to course blueprint Google Docs. Menu: 🎓 Blueprint Tools (`onOpen()` in `Code.gs`).

**Start here:** read `docs/dev-handoff/HANDOFF.md`. It holds the specs for the planned tools (Link Project Documents, Deploy Settings Tables, Draft Settings Tables, Development Status Check), test-file IDs, open questions, and mockups in `docs/dev-handoff/mockups/`.

## Working rules

- Work on branch `link-project-documents` (or a per-tool branch). Do not touch `master` unless asked.
- New files follow the numbering: `Code10.gs` / `Sidebar10.html`, then 11, 12, 13.
- All `.gs` files share one global namespace. Suffix every new function and constant with the file number (e.g. `showLinkDocsSidebar10`), end private helpers with `_`, and grep before defining a name. Reuse existing helpers; do not copy them.
- Match the header style at the top of `Code9.gs`, including the `Last updated on` line.
- Do not add OAuth scopes to `appsscript.json` without listing them and getting Greg's OK; a new scope changes the authorization prompt for every user.
- Tools preview changes and apply only what the user confirms. Never overwrite user content silently.
- No tool may ever tick the settings table's "Check when settings are finalized:" box.

## About the user

Greg is an instructional designer and UX/UI designer who vibe codes; he is not a trained coder. Explain in plain language, say how to test, and skip compliments. He may not be able to run scripts right away, so batch your questions and test requests.
