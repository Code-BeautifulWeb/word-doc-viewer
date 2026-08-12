import { FileView, WorkspaceLeaf, TFile } from "obsidian";
import { renderAsync, type Options } from "docx-preview";

export const WORD_DOCX_VIEW_TYPE = "word-docx-view";

/**
 * View-only document leaf for rendering Microsoft Word (.docx) files natively inside Obsidian.
 */
export class DocxDocumentView extends FileView {
  private activeBlobUrls: Set<string> = new Set();
  private documentContainerEl: HTMLElement;

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

    const loadingEl = this.documentContainerEl.createDiv({
      cls: "word-doc-viewer-status word-doc-viewer-loading",
      text: `Loading "${file.name}"...`,
    });

    try {
      const arrayBuffer = await this.app.vault.readBinary(file);
      loadingEl.remove();

      const renderTarget = this.documentContainerEl.createDiv({
        cls: "word-doc-viewer-content",
      });

      const renderOptions: Partial<Options> = {
        className: "docx-render",
        inWrapper: true,
        ignoreWidth: true,
        ignoreHeight: true,
        ignoreFonts: false,
        breakPages: true,
        useBase64URL: false,
        trimXmlDeclaration: true,
      };

      await renderAsync(arrayBuffer, renderTarget, undefined, renderOptions);

      this.trackEmbeddedBlobUrls(renderTarget);
    } catch (error) {
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

  private trackEmbeddedBlobUrls(root: HTMLElement): void {
    const imgElements = root.querySelectorAll<HTMLImageElement>("img[src^='blob:']");
    imgElements.forEach((img) => {
      if (img.src) {
        this.activeBlobUrls.add(img.src);
      }
    });
  }

  private clearDocumentState(): void {
    for (const url of this.activeBlobUrls) {
      URL.revokeObjectURL(url);
    }
    this.activeBlobUrls.clear();
    this.documentContainerEl.empty();
  }
}
