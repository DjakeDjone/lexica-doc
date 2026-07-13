# wors

Minimal desktop document editor built with Rust + `eframe/egui`.

## Install

### Prebuilt Binaries

The GitHub workflow builds Linux and Windows desktop binaries on each push to `main`,
for pull requests, and when run manually.

Install the latest successful `main` build without downloading the artifact manually:

Linux:

```bash
curl -fsSL https://raw.githubusercontent.com/DjakeDjone/lexica-doc/main/install-prebuilt.sh | bash
```

Windows (PowerShell):

```powershell
irm https://raw.githubusercontent.com/DjakeDjone/lexica-doc/main/install-prebuilt.ps1 | iex
```

The prebuilt installer downloads the latest non-expired GitHub Actions artifact for your
platform and installs `wors` to a local user binary directory. GitHub requires
authenticated artifact downloads, so run `gh auth login` first or set `GH_TOKEN` /
`GITHUB_TOKEN`.

To install manually instead:

1. Open the latest successful **Build desktop app** run in GitHub Actions.
2. Download the artifact for your platform:
   - `wors-linux-x86_64`
   - `wors-windows-x86_64`
3. Extract the downloaded `.zip` archive.

Linux:

```bash
unzip wors-linux-x86_64.zip -d wors-linux-x86_64
chmod +x wors-linux-x86_64/wors
mkdir -p "$HOME/.local/bin"
mv wors-linux-x86_64/wors "$HOME/.local/bin/wors"
```

Windows (PowerShell):

```powershell
Expand-Archive .\wors-windows-x86_64.zip -DestinationPath .
.\wors-windows-x86_64\wors.exe
```

Move `wors.exe` to a directory on your `PATH` if you want to run it from any terminal.

### Build From Source Installer

Linux:

```bash
curl -fsSL https://raw.githubusercontent.com/DjakeDjone/lexica-doc/main/install.sh | bash
```

The installer downloads the source and installs `wors` to `$HOME/.local/bin`.
It also registers a desktop launcher (`wors.desktop`) so the app appears in your system menu.

Windows (PowerShell):

```powershell
irm https://raw.githubusercontent.com/DjakeDjone/lexica-doc/main/install.ps1 | iex
```

The Windows installer downloads the source, builds it with Cargo, installs `wors.exe` to `%USERPROFILE%\.cargo\bin`, adds that directory to your user `PATH` if needed, and creates a Start Menu shortcut.

## Run

```bash
wors
```

## Run In The Browser

Install Trunk if needed:

```bash
cargo install trunk
rustup target add wasm32-unknown-unknown
```

Serve the editor locally:

```bash
trunk serve --open
```

The web build uses the same editor UI as the desktop app. Desktop-only integrations such as native file dialogs and the local LanguageTool process are disabled in the browser build.

## Missing Microsoft Word Compatibility Features

The editor can already create and edit basic formatted documents, tables,
images, sections, headers, footers, and `.docx` files. Full Microsoft Word
compatibility still needs these features:

### File Format And Round-Tripping

- [ ] Preserve every supported WordprocessingML part when opening and saving
      `.docx`, including unknown parts that the editor does not understand yet.
- [ ] Round-trip custom document properties, core metadata, app metadata, and
      document statistics.
- [ ] Preserve themes, color schemes, font schemes, style inheritance, latent
      styles, style aliases, and document defaults.
- [ ] Import and export `.doc`, `.dot`, `.dotx`, `.dotm`, `.docm`, `.rtf`, and
      Word XML formats.
- [ ] Preserve embedded fonts, font substitutions, complex script fonts, East
      Asian fonts, and per-language font fallback.
- [ ] Preserve custom XML parts, content type overrides, package relationships,
      embedded OLE objects, and ActiveX controls.
- [ ] Support strict and transitional OpenXML documents, compatibility settings,
      and documents created by older Word versions.
- [ ] Validate exported documents against OpenXML schemas and Word's repair
      behavior.

### Text And Character Formatting

- [ ] Add subscript, superscript, small caps, all caps, hidden text, emboss,
      imprint, outline, shadow, character scale, character spacing, kerning, and
      text effects.
- [ ] Support advanced underline styles, underline colors, double
      strikethrough, and richer highlight/shading colors.
- [ ] Support OpenType features, ligatures, number forms, stylistic sets, and
      typography options.
- [ ] Add right-to-left, bidirectional, vertical, East Asian, and complex script
      text layout.
- [ ] Add language tagging per run and proofing language controls.
- [ ] Add symbols, special characters, equations, and mathematical layout.
- [ ] Add fields inside text runs, including hyperlinks, cross-references,
      formulas, dates, page references, and document information.

### Paragraph Formatting

- [x] Add first-line, hanging, left, and right indents.
- [ ] Add tab stops with left, center, right, decimal, bar, and leader styles.
- [ ] Add paragraph borders, paragraph shading, drop caps, and text boxes.
- [ ] Add keep with next, keep lines together, widow/orphan control, suppress
      line numbers, and page break control.
- [ ] Add outline levels, collapsed headings, and full heading style behavior.
- [ ] Add paragraph style creation, style editing, style galleries, and style
      inspector behavior.
- [ ] Add multilevel lists, custom numbering formats, restarts, legal numbering,
      list style inheritance, and picture bullets.

### Page Layout And Sections

- [ ] Add all Word page sizes, custom paper trays, gutter margins, mirrored
      margins, book fold, and multiple-pages-per-sheet settings.
- [ ] Add continuous, next page, odd page, and even page section break types.
- [ ] Add columns, column widths, column spacing, separators, and column breaks.
- [ ] Add line numbering, hyphenation, page borders, page color, and watermark
      support.
- [ ] Add footnote and endnote layout, numbering, continuation separators, and
      placement options.
- [ ] Add master-level layout fidelity for pagination, line breaking, widow and
      orphan handling, and Word-compatible measurement rounding.

### Tables

- [ ] Add nested tables.
- [ ] Add table styles, banded rows/columns, header rows, total rows, and style
      priority handling.
- [ ] Add cell margins, cell spacing, preferred widths, exact row heights,
      autofit modes, and table positioning.
- [ ] Add vertical alignment, text direction, cell shading, diagonal borders,
      per-side border styles, and border conflict resolution.
- [ ] Add row splitting across pages, repeating header rows, keep-with-next
      behavior, and table captions.
- [ ] Add formulas, sorting, repeating rows, and richer table editing commands.

### Images, Shapes, And Drawing

- [ ] Add Word shapes, SmartArt, charts, WordArt, icons, 3D models, and
      screenshots.
- [ ] Add text boxes and linked text boxes.
- [ ] Add grouped objects, object rotation, cropping, flip, effects, artistic
      effects, corrections, recolor, and picture styles.
- [ ] Add exact DrawingML import/export for anchors, wrapping polygons, relative
      positions, z-order, clipping, transforms, and object locks.
- [ ] Add captions, alt text editing, accessibility metadata, and decorative
      image flags.
- [ ] Add embedded and linked external media, including videos and audio.

### Headers, Footers, And Fields

- [ ] Add full header/footer galleries and built-in layouts.
- [ ] Add independent first, even, odd, and section-linked header/footer editing
      in the UI.
- [ ] Add all Word page-number formats and numbering restart rules.
- [ ] Add field updating, field code display, field locking, and nested fields.
- [ ] Add date/time, document property, filename, author, and other built-in
      header/footer fields.

### References And Long Documents

- [ ] Add table of contents generation and updating.
- [ ] Add footnotes, endnotes, citations, bibliography, sources, and citation
      styles.
- [ ] Add captions, table of figures, table of authorities, index, bookmarks,
      and cross-references.
- [ ] Add outline view, navigation pane, heading navigation, page thumbnails,
      and document map behavior.
- [ ] Add master documents, subdocuments, and large-document performance
      features.

### Review, Proofing, And Collaboration

- [ ] Add tracked changes for insertions, deletions, formatting changes, moves,
      authors, timestamps, accept/reject, and change display modes.
- [ ] Add comments, threaded comments, resolved comments, mentions, and comment
      anchors.
- [ ] Add compare, combine, protect document, restrict editing, and document
      inspector workflows.
- [ ] Add spelling, grammar, thesaurus, autocorrect, smart quotes, autoformat,
      readability, translation, and accessibility checker parity.
- [ ] Add real-time coauthoring, presence, conflict handling, version history,
      and cloud save integration.

### Mail Merge, Forms, And Automation

- [ ] Add mail merge documents, recipients, rules, labels, envelopes, preview,
      and merge output.
- [ ] Add legacy form fields, content controls, checkboxes, date pickers,
      repeating sections, and form protection.
- [ ] Add macros, VBA projects, macro-enabled document preservation, and macro
      security handling.
- [ ] Add templates, building blocks, Quick Parts, AutoText, themes, and global
      template behavior.
- [ ] Add add-ins, COM/custom task pane equivalents, and automation APIs.

### Editing UI And Commands

- [ ] Add find, replace, go to, advanced search options, and wildcard search.
- [ ] Add clipboard integration for rich Word content, paste options, paste
      special, and format painter.
- [ ] Add full keyboard shortcut parity with Word.
- [ ] Add ruler controls for margins, indents, and tab stops.
- [ ] Add status bar parity, selection statistics, page/section indicators, and
      accessibility labels.
- [ ] Add print preview, printer settings, page setup dialogs, and direct
      printing.
- [ ] Add command-line file opening and file association behavior.
- [ ] Add autosave, autorecover, backup files, recent document pinning, and
      unsaved-change recovery.

### Export And Interoperability

- [ ] Match Word PDF export for pagination, fonts, images, tagged PDF,
      bookmarks, links, comments, and accessibility metadata.
- [ ] Preserve hyperlinks, bookmarks, comments, revisions, fields, metadata, and
      layout when exporting to HTML, PDF, ODT, Markdown, and plain text where
      each format supports them.
- [ ] Add import/export test fixtures that compare against real Word output for
      complex documents.
