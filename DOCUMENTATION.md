# Word Document Viewer Plugin — Technical Documentation

## 1. Plugin Scope

### 1.1 Purpose & Core Goals
The **Word Document Viewer** plugin is a lightweight, view-only extension designed specifically for Obsidian Desktop (macOS / Windows / Linux). It enables users to view Microsoft Word (`.docx`) files directly inside standard Obsidian workspace leaves with accurate rendering fidelity, responsive full-width layout, and authentic paper document styling—without requiring external desktop applications, complex toolbars, or editing frameworks.

### 1.2 In-Scope Functional Requirements
- **Native `.docx` Viewing**: Renders `.docx` OpenXML binary files within native Obsidian `FileView` leaves.
- **Full-Width Responsive Layout**: Eliminates artificial fixed A4 width containers, allowing documents to scale fluidly across all Obsidian pane sizes.
- **Uniform Gaps**: Maintains equal outer margin gaps (16px / `1rem`) on top, right, bottom, and left surrounding the paper document card.
- **Full Content Width**: Forces block-level elements (headings, shaded background callouts, list boxes, tables, paragraphs) to stretch 100% across the document paper, eliminating right-side whitespace gaps.
- **Formatting Fidelity**: Preserves OOXML paragraph spacing (`w:spacing`), line height, font styles, headings, tables, embedded images, and usable hyperlinks.
- **Native Interactions**: Supports browser-native text selection (`user-select: text`), text copying, link navigation, and vertical scrolling.
- **Race Condition Protection**: Uses a render token counter (`renderToken`) to guard against rapid file switching race conditions.
- **Targeted Bullet Normalization**: Normalizes specific Word Private Use Area bullet glyphs (`\uF0B7`, `\uF0A7`, `\uF0D8`) to standard Unicode bullet points (`•`) without corrupting checkboxes or Wingdings arrows.
- **Memory Safety**: Tracks all `blob:` URLs generated for embedded document images and revokes them (`URL.revokeObjectURL`) on view unload or file change.
- **Zero Dependency Bloat**: Uses only `docx-preview` (0.3.7) and `jszip` (3.10.1). Total production bundle size is ~290 KB (`main.js`).

### 1.3 Out-of-Scope / Non-Goals
- **Legacy Binary `.doc` Support**: Legacy Word 97-2003 binary `.doc` files (Compound File Binary Format / OLE2) are explicitly excluded to avoid adding heavy binary parsers.
- **Document Editing**: The viewer is strictly read-only; no inline text modification or document saving capabilities.
- **Complex UI Clutter**: No search overlay bars, sidebars, outline trees, export menus, or unnecessary settings.

---

## 2. Architecture & Design Principles

### 2.1 Component Architecture
The plugin follows a modular, standard Obsidian API architecture:

```
                          ┌───────────────────────────┐
                          │ WordDocumentViewerPlugin  │
                          │     (extends Plugin)      │
                          └─────────────┬─────────────┘
                                        │
                         Registers view & .docx extension
                                        │
                                        ▼
                          ┌───────────────────────────┐
                          │    DocxDocumentView       │
                          │    (extends FileView)     │
                          └─────────────┬─────────────┘
                                        │
                  Reads ArrayBuffer via vault.readBinary(file)
                                        │
                                        ▼
                          ┌───────────────────────────┐
                          │   docx-preview Engine     │
                          │   renderAsync() API       │
                          └─────────────┬─────────────┘
                                        │
                 Descendant Normalization & Bullet Glyph Pass
                                        │
                                        ▼
                          ┌───────────────────────────┐
                          │      styles.css           │
                          │   (Full-width paper CSS)  │
                          └───────────────────────────┘
```

### 2.2 Integration & View Lifecycle
1. **View Registration**: `WordDocumentViewerPlugin` registers `DocxDocumentView` (`WORD_DOCX_VIEW_TYPE = "word-docx-view"`) and binds `.docx` files via `this.registerExtensions(["docx"], WORD_DOCX_VIEW_TYPE)`.
2. **File Loading & Race Prevention**: When a `.docx` file is selected, `DocxDocumentView.onLoadFile(file)` increments `renderToken` and reads the binary content. If a new file is opened before reading completes, the stale task aborts safely.
3. **DOM Rendering**: `renderAsync()` parses the `ArrayBuffer` into structural HTML.
4. **Descendant Width Normalization**: `normalizeRenderedElementWidths()` targets nested `section.docx` descendant blocks (`p`, `div`, `table`, `article`) to strip hardcoded inline widths and right margins.
5. **Targeted Bullet Glyph Repair**: `fixBulletGlyphs()` uses `root.ownerDocument.createTreeWalker()` to replace targeted Private Use Area bullet symbols (`\uF0B7`, `\uF0A7`, `\uF0D8`) with standard Unicode bullets (`•`).
6. **Resource & Leaf Cleanup**: `clearDocumentState()` revokes all retained `blob:` image URLs and increments `renderToken`. On plugin unload (`onunload()`), `app.workspace.detachLeavesOfType(WORD_DOCX_VIEW_TYPE)` cleanly detaches open leaves.

---

## 3. Code Samples & Reference Implementation

### 3.1 Plugin Entry Point — `src/main.ts`
```typescript
import { Plugin } from "obsidian";
import { DocxDocumentView, WORD_DOCX_VIEW_TYPE } from "./DocxDocumentView";

export default class WordDocumentViewerPlugin extends Plugin {
  async onload(): Promise<void> {
    this.registerView(
      WORD_DOCX_VIEW_TYPE,
      (leaf) => new DocxDocumentView(leaf)
    );

    this.registerExtensions(["docx"], WORD_DOCX_VIEW_TYPE);
  }

  onunload(): void {
    this.app.workspace.detachLeavesOfType(WORD_DOCX_VIEW_TYPE);
  }
}
```

### 3.2 View & Rendering Controller — `src/DocxDocumentView.ts`
```typescript
import { FileView, WorkspaceLeaf, TFile } from "obsidian";
import { renderAsync, type Options } from "docx-preview";

export const WORD_DOCX_VIEW_TYPE = "word-docx-view";

export class DocxDocumentView extends FileView {
  private activeBlobUrls: Set<string> = new Set();
  private documentContainerEl: HTMLElement;
  private renderToken = 0;

  private static readonly RENDER_OPTIONS: Partial<Options> = {
    className: "docx-render",
    inWrapper: true,
    ignoreWidth: true,
    ignoreHeight: true,
    ignoreFonts: false,
    breakPages: true,
    useBase64URL: false,
    trimXmlDeclaration: true,
  };

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
    this.documentContainerEl = this.contentEl.createDiv({
      cls: "word-doc-viewer-container",
    });
  }

  getViewType(): string {
    return WORD_DOCX_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.file ? this.file.basename : "Word Document";
  }

  getIcon(): string {
    return "file-text";
  }

  async onLoadFile(file: TFile): Promise<void> {
    await super.onLoadFile(file);
    await this.renderDocument(file);
  }

  async onUnloadFile(file: TFile): Promise<void> {
    this.clearDocumentState();
    await super.onUnloadFile(file);
  }

  async onClose(): Promise<void> {
    this.clearDocumentState();
    await super.onClose();
  }

  private async renderDocument(file: TFile): Promise<void> {
    this.clearDocumentState();
    const token = ++this.renderToken;

    const loadingEl = this.documentContainerEl.createDiv({
      cls: "word-doc-viewer-status word-doc-viewer-loading",
      text: `Loading "${file.name}"...`,
    });

    try {
      const arrayBuffer = await this.app.vault.readBinary(file);
      if (token !== this.renderToken) return;

      loadingEl.remove();

      const renderTarget = this.documentContainerEl.createDiv({
        cls: "word-doc-viewer-content",
      });

      await renderAsync(arrayBuffer, renderTarget, undefined, DocxDocumentView.RENDER_OPTIONS);
      this.trackEmbeddedBlobUrls(renderTarget);
      if (token !== this.renderToken) return;

      this.normalizeRenderedElementWidths(renderTarget);
      this.fixBulletGlyphs(renderTarget);
    } catch (error) {
      if (token !== this.renderToken) return;
      this.documentContainerEl.empty();
      const errorEl = this.documentContainerEl.createDiv({
        cls: "word-doc-viewer-status word-doc-viewer-error",
      });
      errorEl.createEl("h4", { text: "Unable to load document" });
      errorEl.createEl("p", {
        text: error instanceof Error ? error.message : String(error),
      });
    }
  }

  private normalizeRenderedElementWidths(root: HTMLElement): void {
    // 1. Normalize block-level containers (paragraphs, divs, tables, articles)
    const blockElements = root.querySelectorAll<HTMLElement>(
      "section.docx p, section.docx div, section.docx table, section.docx article"
    );
    blockElements.forEach((el) => {
      if (el.style.width && el.tagName !== "IMG") {
        el.style.width = "100%";
      }
      if (el.style.marginRight) {
        el.style.marginRight = "0px";
      }
    });

    // 2. Strip fixed inline pixel widths from table cells & columns (td, th, col, colgroup)
    // allowing browser table-layout: auto to calculate fluid column widths on mobile viewports.
    const tableCellElements = root.querySelectorAll<HTMLElement>(
      "section.docx td, section.docx th, section.docx col, section.docx colgroup"
    );
    tableCellElements.forEach((el) => {
      if (el.style.width) {
        el.style.removeProperty("width");
      }
    });
  }

  private fixBulletGlyphs(root: HTMLElement): void {
    const doc = root.ownerDocument || document;
    const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    let node: Node | null;
    while ((node = walker.nextNode())) {
      if (node.nodeValue && /\uF0B7|\uF0A7|\uF0D8/.test(node.nodeValue)) {
        node.nodeValue = node.nodeValue.replace(/\uF0B7|\uF0A7|\uF0D8/g, "•");
      }
    }
  }

  private trackEmbeddedBlobUrls(root: HTMLElement): void {
    const imgElements = root.querySelectorAll<HTMLImageElement>("img[src^='blob:']");
    imgElements.forEach((img) => {
      if (img.src) {
        this.activeBlobUrls.add(img.src);
      }
    });
  }

  private clearDocumentState(): void {
    this.renderToken++;
    for (const url of this.activeBlobUrls) {
      URL.revokeObjectURL(url);
    }
    this.activeBlobUrls.clear();
    this.documentContainerEl.empty();
  }
}
```

### 3.3 Style Sheet — `styles.css`
```css
.word-doc-viewer-container {
  display: flex;
  flex-direction: column;
  height: 100%;
  width: 100%;
  overflow-y: auto;
  overflow-x: hidden;
  background-color: var(--background-primary);
  box-sizing: border-box;
}

.word-doc-viewer-status {
  padding: 2rem;
  text-align: center;
  color: var(--text-muted);
  font-size: var(--font-ui-medium);
}

.word-doc-viewer-error {
  color: var(--text-error);
}

.word-doc-viewer-error h4 {
  margin-top: 0;
  color: var(--text-error);
}

.word-doc-viewer-content {
  flex: 1;
  width: 100%;
  padding: 1rem;
  box-sizing: border-box;
}

.word-doc-viewer-content .docx-wrapper {
  background: transparent;
  padding: 0;
  width: 100%;
  display: flex;
  flex-direction: column;
  align-items: stretch;
  gap: 1rem;
  box-sizing: border-box;
}

.word-doc-viewer-content .docx-wrapper > section.docx {
  width: 100%;
  box-sizing: border-box;
  background-color: #ffffff;
  color: #111111;
  border-radius: 4px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
  padding: 2.5rem 3rem;
  margin: 0;
  min-height: auto;
}

.word-doc-viewer-content .docx-render section.docx > * {
  width: 100% !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
  margin-right: 0 !important;
}

.word-doc-viewer-content .docx-render p,
.word-doc-viewer-content .docx-render div,
.word-doc-viewer-content .docx-render article,
.word-doc-viewer-content .docx-render table {
  width: 100% !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
  margin-right: 0 !important;
}

.word-doc-viewer-content .docx-render,
.word-doc-viewer-content .docx-render * {
  -webkit-user-select: text;
  user-select: text;
}

.word-doc-viewer-content .docx-render img {
  max-width: 100% !important;
  width: auto !important;
  height: auto !important;
  display: inline-block;
}

.word-doc-viewer-content .docx-render table {
  max-width: 100% !important;
  border-collapse: collapse;
  margin: 0.8em 0;
}

.word-doc-viewer-content .docx-render p {
  line-height: inherit;
}
```

---

## 4. Future Possibilities

While the current version prioritizes a minimal, lightweight, zero-bloat architecture, future non-breaking enhancements could include:

1. **Page Number Indicator**: A subtle, floating page counter (e.g., `Page 1 of 5`) at the bottom corner of the view leaf that updates dynamically during scroll events.
2. **Native Obsidian Printing**: Integration with Obsidian's internal print service (`window.print()` scoped to the document node) for exporting to PDF or paper without capturing the Obsidian app UI.
3. **Optional Dark Mode Paper Theme Setting**: A simple user toggle in plugin settings to choose between:
   - *Authentic Paper* (white page card, default).
   - *Obsidian Theme Adaptive* (page background adapts to dark/light mode CSS variables).
4. **Large Document Virtualization**: For extremely large `.docx` files (500+ pages / >50MB), applying `content-visibility: auto` to individual page sections to optimize DOM rendering performance.
