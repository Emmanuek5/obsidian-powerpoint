import { Plugin } from 'obsidian';
import { PptxView, PPTX_VIEW_TYPE } from './PptxView';
import { cleanupOldCache } from './converter';

export default class PowerPointPlugin extends Plugin {
  onload() {
    console.debug('[PPTX Plugin] Loading PowerPoint Viewer plugin');

    // Clean up old cached PDFs (older than 7 days)
    cleanupOldCache(7);

    this.registerView(
      PPTX_VIEW_TYPE,
      (leaf) => {
        console.debug('[PPTX Plugin] Creating new PptxView');
        return new PptxView(leaf);
      }
    );

    console.debug('[PPTX Plugin] Registering .pptx and .ppt extensions');
    this.registerExtensions(['pptx', 'ppt'], PPTX_VIEW_TYPE);

    this.addCommand({
      id: 'open-pptx-file',
      name: 'Open PowerPoint file',
      callback: () => {
        const file = this.app.workspace.getActiveFile();
        console.debug('[PPTX Plugin] Command triggered, active file:', file?.path);
        if (file && (file.extension === 'pptx' || file.extension === 'ppt')) {
          const leaf = this.app.workspace.getLeaf('tab');
          leaf.openFile(file, { active: true });
        }
      }
    });
  }

  onunload() {
    console.debug('Unloading PowerPoint Viewer plugin');
  }
}
