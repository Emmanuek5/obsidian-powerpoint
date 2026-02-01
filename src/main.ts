import { Plugin, Notice } from 'obsidian';
import { PptxView, PPTX_VIEW_TYPE } from './PptxView';
import { cleanupOldCache, clearCache, getCacheStats } from './converter';

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
      name: 'Open powerpoint file',
      callback: () => {
        const file = this.app.workspace.getActiveFile();
        console.debug('[PPTX Plugin] Command triggered, active file:', file?.path);
        if (file && (file.extension === 'pptx' || file.extension === 'ppt')) {
          const leaf = this.app.workspace.getLeaf('tab');
          void leaf.openFile(file, { active: true });
        }
      }
    });

    this.addCommand({
      id: 'clear-pptx-cache',
      name: 'Clear powerpoint cache',
      callback: () => {
        const stats = getCacheStats();
        const sizeMB = (stats.totalSize / (1024 * 1024)).toFixed(2);
        clearCache();
        new Notice(`Cleared PowerPoint cache: ${stats.count} files (${sizeMB} MB)`);
        console.debug('[PPTX Plugin] Cache cleared:', stats);
      }
    });
  }

  onunload() {
    console.debug('Unloading PowerPoint Viewer plugin');
  }
}
