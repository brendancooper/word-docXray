# Word DocXray — Implementation Plan

## Purpose

Build a Microsoft Word **task-pane add-in** for exploring and carefully editing WordprocessingML/OOXML from inside Word desktop, including **Word for macOS**.

The product is an XML inspection and editing tool. It deliberately does **not** include Markdown import/export, document conversion, or a general DOCX-to-Markdown pipeline.

## Goals

- Run in desktop Microsoft Word on **macOS** and Windows through the supported Office Add-in / Office.js model.
- Open a task pane from Word that shows formatted OOXML for the current selection, range, or document body.
- Provide a capable XML editor: syntax highlighting, pretty formatting, search, validation, and an unsaved-change indicator.
- Let users refresh OOXML from Word and apply validated changes back to Word.
- Provide a readable preview and a formatting inspector for common Word content.
- Make edits safe: clear scope, confirmation/diff before apply, validation, and recovery guidance.

## Non-goals

- Markdown import, export, or conversion.
- A full Word rendering engine; Word itself remains the authoritative renderer.
- A general-purpose editor for every file in a DOCX ZIP package.
- Editing media, relationships, custom XML parts, headers/footers, themes, or arbitrary package files in the first release.
- Replacing Word's normal save mechanism or modifying the open `.docx` ZIP package directly.

## Product model and constraints

### Why Office.js

VSTO/COM is Windows-only and cannot satisfy the macOS requirement. The add-in should use Office.js with a Word task pane, which is supported by Word desktop on macOS and Windows.

### OOXML boundary

The add-in works with OOXML exposed by Word APIs (selection, range, or body), rather than treating the current document as an unrestricted writable ZIP package. Depending on the API and object used, the returned content may be a fragment or Flat OPC-compatible OOXML rather than the literal `word/document.xml` package entry.

This is intentional: it allows Word to validate and integrate inserted content. UI language must call this **OOXML/XML from Word**, not promise raw package-part access.

### API compatibility

Before implementing each feature, verify its Word API requirement set and support in the targeted Word for Mac releases. The add-in should:

- declare the minimum required `WordApi` requirement set in its manifest;
- use `Office.context.requirements.isSetSupported(...)` at startup;
- disable unavailable controls with a helpful explanation;
- test specifically in current Word for Mac desktop, not only Word on the web or Windows.

## Primary workflows

### 1. Explore current selection

1. User selects text, a table, or another supported range in Word.
2. User opens DocXray from the ribbon.
3. The task pane loads OOXML for the selection.
4. XML is pretty-printed and displayed in the editor.
5. The inspector reports scope, paragraph styles, fonts, colours, and table information where available.

This is the safest initial workflow because applying replacement XML has a narrow, explicit scope.

### 2. Explore document body

1. User selects **Document body** in the add-in scope control.
2. The add-in retrieves OOXML for the body.
3. The XML editor and preview update.
4. Applying changes replaces/inserts only the supported body-level OOXML scope after confirmation.

### 3. Edit and apply

1. User changes XML.
2. The add-in validates that it is well-formed XML and runs basic OOXML safety checks.
3. User selects **Review changes**.
4. The add-in presents a text diff and warns that Word may reject unsupported OOXML.
5. User selects **Apply to Word**.
6. The add-in uses the appropriate Word OOXML insertion API for the selected scope.
7. On success, it refreshes XML from Word so the editor reflects Word's canonicalized result.

### 4. Refresh / discard

- **Refresh from Word** retrieves the current scope again and asks before discarding unsaved editor changes.
- **Discard changes** restores the most recently retrieved XML snapshot.
- The add-in never silently applies XML merely because the document is saved.

## User interface

### Ribbon commands

- **Open DocXray** — opens/focuses the task pane.
- Optional: **Inspect selection in DocXray** — opens the pane and immediately loads the current selection.

### Task pane layout

```text
┌─────────────────────────────────────────────┐
│ DocXray   Scope: [Selection v]  Refresh     │
│ Status: Loaded from Word · no local changes │
├───────────────────────┬─────────────────────┤
│ XML editor            │ Preview / Inspector │
│                       │                     │
│ formatted OOXML       │ rendered structure  │
│                       │ styles and props    │
├───────────────────────┴─────────────────────┤
│ Format  Validate  Review changes  Apply      │
└─────────────────────────────────────────────┘
```

### Editor

Use Monaco Editor (or CodeMirror if bundle size becomes a concern) configured for XML.

Required capabilities:

- XML syntax highlighting and bracket matching
- line numbers, minimap optional
- pretty-format command
- search/replace
- dirty state
- validation diagnostics in the gutter
- keyboard-accessible controls
- preserve user formatting in the editor; formatting is explicit, never automatic while typing

### Preview and inspector

The preview is supplementary; Word is the final visual authority.

Initial preview support:

- paragraphs, lists, and page breaks
- built-in and identifiable custom heading styles
- run-level bold/italic/underline, font family, size, colour, highlight, superscript/subscript
- tables, grid columns, cell shading, spans, and basic alignment
- paragraph alignment, indentation, and spacing

Inspector fields:

- loaded scope and object type
- OOXML size and last refresh time
- selection paragraph/style IDs where available
- detected tables, paragraphs, runs, and validation errors
- API capability/support status

## Technical architecture

```text
Word for Mac / Windows desktop
  │
  ├── Office Add-in manifest
  │     └── Ribbon command + task pane URL
  │
  └── Task-pane web application
        ├── Office.js adapter
        │     ├── obtain current selection/body OOXML
        │     └── insert validated OOXML into the chosen scope
        ├── XML service
        │     ├── pretty printer
        │     ├── well-formedness parser
        │     ├── OOXML guardrails
        │     └── diff generation
        ├── Monaco XML editor
        ├── Preview renderer
        └── UI state / undo snapshots
```

### Suggested stack

- TypeScript
- Office.js (`@types/office-js`)
- React + Vite, or a lightweight TypeScript UI if preferred
- Monaco Editor
- XML parser/formatter that does not alter text-node content
- a small client-side diff library
- HTTPS development server for sideloading on macOS

No backend is required for the core product. A backend should only be introduced for optional telemetry, centralized deployment, or shared policy/configuration.

## XML handling

### Pretty formatting

- Format only on explicit user action or initial load.
- Preserve text-node content and XML attributes exactly.
- Do not use a simplistic `>\s+<` replacement as the only transformation, since whitespace can be meaningful in some XML contexts.
- Retain a raw loaded snapshot and a formatted editor representation separately if necessary.

### Validation layers

1. **Well-formed XML:** parse in the browser and report line/column where possible.
2. **Scope validation:** require a supported root/fragment form for the active insertion API.
3. **Safety checks:** reject or prominently warn on DTDs, external entities, unsupported package references, malformed namespace declarations, and content outside the chosen scope.
4. **Word apply result:** catch Office.js/Word insertion errors and leave the editor content intact.

### Apply model

Start with selection/range replacement only. The exact insertion/replacement mechanism must be proven with real Word for Mac tests before exposing document-body apply.

After every successful apply:

- call `context.sync()`;
- re-read OOXML from the same scope;
- replace the editor contents with Word's accepted/canonicalized representation;
- record an activity message.

## Delivery phases

### Phase 0 — Technical spike

**Outcome:** verify the actual Word for Mac APIs and document the supported XML shapes.

- Scaffold a minimal Office.js Word add-in.
- Sideload it in current Word for Mac.
- Retrieve OOXML for selection, range, and body where supported.
- Test insertion of harmless edited paragraphs, styled runs, and simple tables.
- Determine whether returned data is fragment OOXML, Flat OPC, or another representation for each scope.
- Record required Word API set/version and failure behavior.
- Decide the narrow, safe first-release apply scope.

**Exit criteria:** a selection can be read, prettified, changed, and safely applied on Word for Mac.

### Phase 1 — Explorer MVP

- Create the add-in manifest and task-pane shell.
- Add scope picker: Selection and Document body (read-only body initially if needed).
- Implement fetch/refresh OOXML.
- Add Monaco XML editor.
- Implement explicit pretty formatting and well-formed XML validation.
- Add dirty-state protection and raw snapshot/discard behavior.
- Add macOS sideloading/developer documentation.

**Exit criteria:** users can reliably inspect and copy formatted OOXML from Word for Mac.

### Phase 2 — Safe editing MVP

- Add selection-only apply flow.
- Add pre-apply XML validation and confirmation.
- Add text diff against last loaded snapshot.
- Add robust error messages and post-apply refresh.
- Test normal undo behavior after an apply operation.

**Exit criteria:** supported selection XML edits can be made and reverted safely in Word for Mac.

### Phase 3 — Preview and inspector

- Port/adapt the DocXray browser preview renderer.
- Build an OOXML style/format inspector.
- Improve custom-heading recognition and direct/style-inherited formatting where source XML provides it.
- Add table grid, span, shading, and alignment rendering.
- Clearly label preview limitations.

### Phase 4 — Broader scope and hardening

- Evaluate body-level apply based on Phase 0/2 findings.
- Add larger-document performance handling and cancellation.
- Add accessibility review, keyboard navigation, high-contrast support, and localization-ready strings.
- Add telemetry only with explicit consent, if desired.
- Prepare production manifest, HTTPS hosting, signing/deployment documentation, and release process.

## Test plan

### Environment matrix

At minimum:

- Current Microsoft 365 Word for Mac desktop
- One supported older Word for Mac version, if product support requires it
- Current Microsoft 365 Word for Windows desktop
- Word on the web as an informational compatibility check, not the primary target

### Document fixtures

Maintain fixtures containing:

- plain paragraphs and whitespace-sensitive text
- built-in and custom heading styles
- direct and theme-based colours
- custom fonts, highlighting, and emphasis
- aligned/indented/spaced paragraphs
- simple and merged-cell tables
- nested tables, lists, hyperlinks, fields, tracked changes, headers/footers, images, and content controls
- malformed XML test cases for validation UI

### Key assertions

- Loading does not change the document.
- Formatting XML does not change text-node meaning.
- Refresh warns before replacing dirty editor content.
- Invalid XML cannot be applied.
- A failed Word insertion does not corrupt the current document or erase local edits.
- Successful apply re-reads and displays Word's actual accepted OOXML.
- Selection edits stay inside the intended scope.
- The task pane is usable and visually correct on macOS Word.

## Risks and mitigations

| Risk | Mitigation |
| --- | --- |
| Word for Mac exposes different OOXML capability/behavior than Windows | Make the Mac technical spike the first milestone; gate features on verified capability. |
| Word rejects valid-looking OOXML fragments | Restrict first release to tested scope/root shapes; show validation and preserve editor content on failure. |
| Users expect package-level `document.xml` editing | Use precise terminology in UI/docs and link advanced package editing to the VS Code extension. |
| Preview differs from Word | Label it as structural preview; always provide refresh and rely on Word as visual authority. |
| Applying XML could lose content | Default to selection scope, require review/confirmation, retain snapshots, and encourage normal Word save/undo. |
| Macro/add-in deployment friction on Mac | Use Office Add-in sideloading and a documented HTTPS dev setup; avoid VBA as the primary product. |

## Repository bootstrap

Create this repository as a standalone Office Add-in project rather than extending the VS Code extension codebase directly.

Suggested initial structure:

```text
word-docXray/
  PLAN.md
  README.md
  package.json
  manifest.xml
  src/
    taskpane/
    office/
    xml/
    preview/
  public/
  test/
    fixtures/
  docs/
```

The existing DocXray preview code may be reused conceptually, but it should be refactored into testable browser-side modules rather than copied as an inline webview script.
