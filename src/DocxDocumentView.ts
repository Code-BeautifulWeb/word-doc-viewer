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
    const token = ++this.renderToken;
    this.clearDocumentState();

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
      if (token !== this.renderToken) return;

      this.normalizeRenderedElementWidths(renderTarget);
      this.fixBulletGlyphs(renderTarget);
      this.trackEmbeddedBlobUrls(renderTarget);
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
