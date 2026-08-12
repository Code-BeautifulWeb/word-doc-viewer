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
