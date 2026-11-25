# Word DocXray

A Microsoft Word task-pane add-in for inspecting OOXML that Word exposes for a selection or document body. It is deliberately **not** a DOCX package editor: it does not edit `word/document.xml` in a ZIP file.

## Current implementation

- Selection and body OOXML exploration through Office.js
- Monaco XML editor, explicit formatting, dirty-state protection, XML/DTD safety checks, and a change review
- Selection-only replacement with confirmation and post-apply reload
- Structural preview and basic OOXML counts/style inspector

Body scope is read-only by design until replacement behavior is validated on current Word for Mac builds.

## Develop and sideload

1. Install Node 20+ and run `npm install`.
2. Run `npm run dev`. Vite serves HTTPS on `https://localhost:3000`; trust its development certificate if Word does not load it.
3. In Word, sideload `manifest.dev.xml` using the Office Add-in upload/sideload flow.
4. Open **DocXray** from the Home ribbon, select content, and choose **Refresh from Word**.

`manifest.dev.xml` points to `https://localhost:3000` and is only for development.

## Install a release

A `v*` tag triggers GitHub Actions to build the task pane, deploy it to GitHub Pages at `https://brendancooper.github.io/word-docXray/`, and attach two downloads to the GitHub Release:

- `word-docXray-vX.Y.Z-manifest.xml` — the only file an individual user needs to install.
- `word-docXray-vX.Y.Z.zip` — the static build plus manifest, retained for audit/offline deployment.

Download the versioned **manifest XML** and upload it in Word: **Insert → Add-ins → My Add-ins → Manage My Add-ins → Upload My Add-in**. The add-in then opens in Word's task pane and loads its static UI from GitHub Pages; the user does not run a local server or install the ZIP.

Repository administrators must make the repository public and set **Settings → Pages → Build and deployment → Source** to **GitHub Actions** before publishing the first tag. GitHub Pages hosting is static HTTPS hosting, not a separate user-facing application process.

## Safety model

Never apply unreviewed XML. DocXray rejects malformed XML, DTD/entity declarations, and obvious external URLs; Word remains the final validator. Failed insertion retains the editor content. Save normally in Word and use Word Undo to recover from an accepted but unwanted operation.

## Compatibility

The add-in requires `WordApi` 1.1 and checks that requirement set at runtime. Test on the specific Word for macOS versions you intend to support; returned OOXML and insertion acceptance can differ between Word clients.
