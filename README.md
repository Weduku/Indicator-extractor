# Indicator Extractor — Web App (GitHub Pages)

This is a browser-only version of the Indicator Extractor. It runs Python
*inside your browser* via [Pyodide](https://pyodide.org) — your `.docx`
files are read and processed entirely on your device. Nothing is uploaded
anywhere unless you explicitly use the Cloud Sync tab to push rows to your
own Excel file on OneDrive/SharePoint.

## Files

- `index.html` — the page
- `style.css` — styling
- `app.js` — all the app logic (Pyodide setup, storage, UI, Cloud Sync)
- `extractor.py` — the extraction/breakdown/pivot logic (same tested
  behavior as the desktop app), executed inside the browser via Pyodide

## 1. Test it locally first

Pyodide requires the page to be served over `http://`, not opened directly
as a file — browsers block the necessary requests under `file://`. From
this folder, run a simple local server:

```
python3 -m http.server 8000
```

Then open **http://localhost:8000** in your browser. The first load takes
a few seconds while Pyodide and `python-docx` download (this needs
internet the first time; after that your browser caches them).

**Try it**: go to the Indicators tab, add a label, go to Extract, upload a
`.docx` file with that label in it, click Process. Reload the page — your
data should still be there (that's the "append, don't overwrite" dataset,
now stored in this browser's IndexedDB instead of a local file).

## 2. Put it on GitHub

1. Create a new GitHub repository (private is fine, or public — see the
   note on visibility below).
2. Add these four files to the repo root (or a subfolder — see step 3).
3. Commit and push.

## 3. Turn on GitHub Pages

1. In the repo, go to **Settings → Pages**.
2. Under **Source**, choose **Deploy from a branch**.
3. Pick your branch (usually `main`) and the folder these files are in
   (`/root` if they're at the top level).
4. Save. GitHub gives you a URL like:
   ```
   https://<your-username>.github.io/<repo-name>/
   ```
   It can take a minute or two to go live after the first save.

**Note on visibility**: GitHub Pages sites are publicly reachable by
default — anyone with the URL can open the app (not your data, just the
tool itself, since everything runs in *their* browser with *their* data).
If that's not acceptable for your org, GitHub Pages supports private/
restricted access only on **GitHub Enterprise** plans — worth checking
with your GitHub admin if this needs to stay internal.

## 4. Cloud Sync setup (Google Drive folder)

Someone with access to a Google Cloud project needs to do this once:

1. Go to **Google Cloud Console → APIs & Services → Library** and enable
   the **Google Drive API** and **Google Sheets API**.
2. **APIs & Services → Credentials → Create Credentials → OAuth client ID**.
   Application type: **Web application**.
3. Under **Authorized JavaScript origins**, add the URL this page is hosted
   at from step 3 above (just the origin, e.g. `https://you.github.io` —
   no path or trailing slash).
4. If prompted, configure the **OAuth consent screen** — you can restrict
   it to internal users only if this is a Google Workspace org.
5. Copy the **Client ID** into the Cloud Sync tab in the app.

Then in the app: paste a link to the Google Drive **folder** you want to
work with, set a name for the results Sheet (created automatically inside
that folder the first time you sync), and click **Sync Now**. That:

- Signs you in via a Google popup (no separate device code, unlike the
  desktop app's Microsoft flow).
- Lists the `.docx` files in that folder and downloads/extracts only the
  ones that are new or have changed since your last sync (tracked locally,
  so re-running Sync doesn't reprocess everything every time).
- Pushes any not-yet-synced rows into the results Sheet in that same
  folder, creating the Sheet (and its header row) automatically if it
  doesn't exist yet.

The app only requests read access to files in the folder you link, plus
permission to create/edit its own results Sheet — not your whole Drive.

## 5. Where your data actually lives

Everything is stored **in this specific browser, on this specific
device** — via IndexedDB (the dataset) and localStorage (your indicator
list, Cloud Sync settings, and the record of which Drive files have
already been synced). There's no server-side storage at all.

Practical implications:
- Switching browsers or devices means starting with an empty dataset there
  — use the **"Download data backup (.db)"** button (bottom of the page)
  to export, and **"Restore backup"** on the other browser/device to bring
  it back.
- Clearing your browser's site data/cache for this page will delete your
  dataset **and** your "already synced" file-tracking — a fresh sync would
  then reprocess every file in the folder again (harmless, just re-adds
  the same rows, which then get skipped on the push step since the Sheet
  already has them... actually re-adding will create duplicate rows in
  your local dataset, so restore a backup rather than starting fresh if
  you want to avoid that).
- The **Sync Now** step is the one place your data leaves this browser,
  and only to a folder/Sheet you already own.

## 6. Reading data from images (OCR)

You can also add standalone image files (screenshots, scans, photos of a
printed report) directly in the Extract tab, and images embedded inside a
`.docx` file (like a table pasted in as a picture) are picked up
automatically. This runs via **Tesseract.js** — Tesseract's OCR engine
compiled to WebAssembly — entirely in your browser, same as everything
else here. Nothing is uploaded to a server for this either.

- There's a checkbox on the Extract tab ("OCR images") to turn this on or
  off per run — on by default.
- The first time you OCR anything, your browser downloads Tesseract's
  language data (a few MB) and caches it, same one-time-cost pattern as
  Pyodide itself loading on first visit.
- Results from images are tagged with source `image` (standalone file) or
  `image in document` (found inside a `.docx`) in the Data tab, so you can
  always tell which values came from OCR versus native document text.

## 7. What's different from the desktop version

- **No local NLP fallback** (spaCy) in this version — running a full NLP
  model inside the browser is a large, heavy download, so this build
  sticks to the rule-based extraction only, which is the same reliable
  core the desktop app leans on anyway. (OCR is still included — that's a
  separate, much lighter WebAssembly library.)
- **Cloud Sync targets Google Drive/Sheets** instead of OneDrive/
  SharePoint Excel, and reads source documents directly from a linked
  Drive folder rather than requiring manual upload.
- **Storage is per-browser** (IndexedDB/localStorage) instead of a file
  next to an `.exe`.

Everything else — indicator matching, breakdown detection (population/
gender/household), reporting periods, append-not-overwrite behavior, OCR
support, and the pivot table builder — is the same tested logic as the
desktop app.

## 8. Troubleshooting

- **"ModuleNotFoundError: No module named 'X'" during boot** — Pyodide
  ships a slimmed-down Python standard library and loads a few modules
  (like `sqlite3`) on demand rather than bundling them by default. This is
  already handled for everything this app uses, but if you extend
  `extractor.py` with a new stdlib import and hit this, add the module
  name to the `pyodide.loadPackage([...])` list near the top of `boot()`
  in `app.js` — the error message itself will tell you the exact fix.
- **Nothing happens / stuck on "Loading Python engine"** — check the
  browser's developer console (F12) for the actual error; the most common
  cause is no internet connection on first load (Pyodide needs to download
  itself and a couple of packages once).

## 9. A note on testing

I built and unit-tested the Python extraction logic itself (identical to
the desktop app's tested behavior), including the OCR-matching path. I was
not able to do a live, in-browser test of the full Pyodide + Tesseract.js +
IndexedDB + Google sign-in flow together — that requires an actual browser
and internet connection I don't have access to while building this. Treat
your first real run (step 1 above) as the real test, and let me know what
happens if anything doesn't work as expected — any error message helps me
fix it quickly.
