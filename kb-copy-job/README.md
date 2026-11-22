# KB Copy Job (prototype)

Copies eCampus Knowledge Base (Confluence) pages into Google Docs in the
**Gemini Notebook Project Work Station Stuff** folder, so the Blueprint Tools'
**Set Up Gemini Notebook** can list them as notebook sources. Gemini Notebook
can't open Confluence pages that need a sign-in, but it can open Google Docs.

- **Not part of the Blueprint Tools script.** It runs in its own Apps Script
  project, attached to a private Google Sheet that only Greg edits.
- **Only Greg's account holds the Confluence token.** The 14 IDCs never run this
  and never see the token. They only need *view* access to the output folder.
- **Free to run.** Google runs it on a schedule. Claude isn't involved.

## Which pages it copies

1. Every Confluence page linked in the **Gemini Notebook Reference List** Sheet
   (one row per reference: Category, Name, Notes, with each name linked to its page).
2. Any page link pasted into column A of the Sheet's **Pages to copy** tab. Use
   this for testing, for example the *CDD Project Documents – Ownership and Share
   Settings* page.

Use full page links (`…/wiki/spaces/EKB/pages/2274689025/…`). Short `/wiki/x/…`
links aren't recognised yet.

## One-time setup (about 15 minutes)

1. **Make the control Sheet.** Create a new Google Sheet named
   *KB Copy Job (do not share)*. Keep it private: anyone who can edit it can edit
   its script.
2. **Add the script.** In the Sheet, open **Extensions → Apps Script**. Replace
   the contents of `Code.gs` with `KbCopyJob.gs` from this folder.
3. **Turn on the Drive service.** In the editor's left bar, click **Services
   (+)**, choose **Drive API**, keep version **v3** and identifier **Drive**,
   and click **Add**.
4. **Set the permissions file.** Open **Project Settings** (gear icon), tick
   **Show "appsscript.json" manifest file in editor**, open `appsscript.json`,
   and replace its contents with `appsscript.json` from this folder.
5. **Reload the Sheet.** A **📚 KB Copy Job** menu appears. The first time you
   use it, Google asks you to approve the permissions.
6. **Create the Confluence token.** Go to
   <https://id.atlassian.com/manage-profile/security/api-tokens> and click
   **Create API token with scopes**:
   - Name: `KB Copy Job`
   - Expiry: as short as practical (set a calendar reminder to renew it)
   - App: **Confluence**
   - Scope: **`read:page:confluence`** (read-only)

   Copy the token. Atlassian shows it only once.
7. **Save it.** Choose **📚 KB Copy Job → 1. Save Confluence sign-in**. Enter your
   Atlassian email, then paste the token. Don't paste it anywhere else.
8. **Test.** Paste one page link into the **Pages to copy** tab. Choose
   **2. Test connection**, then **3. Copy pages now**. The **Copies** tab shows each
   page, its Google Doc, the version copied, and the result.
9. **Nightly copy.** When the test works, choose **Turn ON nightly copy**
   (runs 2–3 a.m.).

## What to expect

- One Doc per page, named **KB – <page title>**. The Doc keeps the same link on
  every refresh, so a notebook source made from it can just be re-synced.
- A page is rewritten only when its Confluence version changes.
- **Images are not copied** (Confluence images need a sign-in). Each one becomes
  *[Image – see the original page]*.
- Don't edit the Docs. The next copy overwrites them. Edit the page in Confluence.

## If something goes wrong

The **Copies** tab's *Result* column says what happened in plain language:

| Message mentions | Do this |
|---|---|
| "did not accept the email and token" | Re-save the sign-in. Check the token hasn't expired. |
| "not allowed to read page" | For a scoped token, add the `read:page:confluence` scope. Otherwise check you can open the page. |
| "not found" | The page was moved or deleted. Update its link. |

**If the token may have leaked:** delete it at id.atlassian.com → Security → API
tokens, choose **Forget Confluence sign-in**, then create and save a new one.

**Handing it over:** the nightly copy runs under the account that turned it on.
If that person leaves, the Docs stay but stop updating. To take over, the new owner
makes their own copy of the Sheet, then repeats steps 6–9 with their own token.
