# Gemini Notebook API test (personal account only)

A throwaway test to find out whether a script can create a Gemini Notebook and
fill it with the kinds of sources **Set Up Gemini Notebook** would add. Run it
in a **personal Google account**, never your Boise State account. The API is in
preview and limited to personal accounts. Use **copies** of test documents.

## What we already know (from Google's published API description, v1, Nov 20, 2026)

| The API can | The API can't |
|---|---|
| Create, list, rename and read notebooks | Add a **Google Drive** file as a Drive source (the kind that syncs automatically) |
| Add a source from text or Markdown | Delete notebooks or sources |
| Add a source from a web page URL | |
| Upload a file (such as a PDF) as a source | |

So the test checks whether a Google Doc works when the script **converts it to a
PDF or to Markdown text and uploads it**, or when the Doc is made public and added
as a **web page**. All of these are snapshots: in Gemini Notebook only sources
added from Drive update automatically.

## Setup (about 20 minutes, once)

1. **Google Cloud project.** Signed in to your personal account, go to
   <https://console.cloud.google.com>, create a project (for example
   *Notebook API test*), and note its **Project number** (on the project's
   Dashboard).
2. **Turn on the API.** In that project, go to **APIs & Services → Library**,
   search for **Gemini Notebook API**, and click **Enable**. Also enable the
   **Google Drive API**.
3. **Consent screen.** Go to **APIs & Services → OAuth consent screen**
   (it may be called *Google Auth Platform*). Choose **External**, fill in the
   app name and your email, and add **yourself as a test user**. Leave it in
   *Testing*.
4. **The script.** Go to <https://script.google.com>, create a **New project**,
   and paste `NotebookApiTest.gs` over the default code.
5. **Permissions file.** In **Project Settings** (gear icon), tick **Show
   "appsscript.json" manifest file in editor**, then replace that file's
   contents with `appsscript.json` from this folder.
6. **Link the Cloud project.** Still in Project Settings, under **Google Cloud
   Platform (GCP) Project**, click **Change project** and enter the project
   number from step 1.
7. **A test Doc.** Put a copy of a test Google Doc in this personal Drive.
   Copy its ID (the long part of its link between `/d/` and `/edit`) into
   `TEST_DOC_ID` at the top of the script.
8. **Public-link tests (7a–7c).** Set that test Doc's sharing to **Anyone with
   the link (Viewer)**. Optionally also choose **File → Share → Publish to web →
   Publish**, and paste the published link (it contains `/pub`) into
   `PUBLISHED_URL`. Do this only with test copies: "anyone with the link" means
   anyone on the internet who has the link can read it.

## Run it

1. Choose **runAllTests** in the function menu and click **Run**. Approve the
   permissions prompt. Google may warn that the app isn't verified: click
   **Advanced → Go to … (unsafe)**. It's your own test script.
2. Open **View → Execution log** (or the log panel). You'll get a summary like:

   ```
   ✅ 1. List notebooks (setup check)
   ✅ 2. Create a test notebook
   ✅ 3. Text source
   ✅ 4. Public web page source
   ✅ 5. Google Doc uploaded as a PDF
   ✅ 6. Google Doc added as Markdown text
   ❌ 7a. Doc's normal /edit link as a web page — …
   ✅ 7b. Doc's plain-text export link as a web page
   8. Source states after 20 seconds:
      ✅ API test – plain text: ACTIVE
      ⏳ My test doc (PDF upload): PROCESSING
   ```

3. If anything shows **PROCESSING**, wait a few minutes and run
   **checkSourcesAgain**.
4. Open Gemini Notebook in that personal account and look at the test notebook.
   Ask it a question about your test Doc to check the content came through.
5. **Delete the test notebook by hand**; the API can't.

Send Claude a copy of the log, or a screenshot of it.

## If step 1 fails

The log adds a hint after each error:

| Message mentions | Likely fix |
|---|---|
| SERVICE_DISABLED / "has not been used" | Step 2: enable the Gemini Notebook API in the linked project |
| scope / insufficient | Step 5: check `appsscript.json`, then run again and re-approve |
| other 403 | The API may not be available to this account yet |
| 401 | Run again and approve the permissions prompt |
