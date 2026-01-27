import { Plugin, TFile } from 'obsidian';
import { PptxView, PPTX_VIEW_TYPE } from './PptxView';

export default class PowerPointPlugin extends Plugin {
  async onload() {
    console.log('[PPTX Plugin] Loading PowerPoint Viewer plugin');

    this.registerView(
      PPTX_VIEW_TYPE,
      (leaf) => {
        console.log('[PPTX Plugin] Creating new PptxView');
        return new PptxView(leaf);
      }
    );

    console.log('[PPTX Plugin] Registering .pptx and .ppt extensions');
    this.registerExtensions(['pptx', 'ppt'], PPTX_VIEW_TYPE);

    this.addCommand({
      id: 'open-pptx-file',
      name: 'Open PowerPoint file',
      callback: () => {
        const file = this.app.workspace.getActiveFile();
        console.log('[PPTX Plugin] Command triggered, active file:', file?.path);
        if (file && (file.extension === 'pptx' || file.extension === 'ppt')) {
          const leaf = this.app.workspace.getLeaf('tab');
          leaf.openFile(file, { active: true });
        }
      }
    });
  }

  onunload() {
    console.log('Unloading PowerPoint Viewer plugin');
    this.app.workspace.detachLeavesOfType(PPTX_VIEW_TYPE);
  }
}
