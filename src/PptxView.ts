import { FileView, WorkspaceLeaf, TFile } from 'obsidian';
import * as pdfjsLib from 'pdfjs-dist';
import { convertPptxToPdf, cleanupPdf } from './converter';

export const PPTX_VIEW_TYPE = 'pptx-view';

// Set worker path for PDF.js - use jsdelivr CDN matching installed version
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@5.4.530/build/pdf.worker.min.mjs';

export class PptxView extends FileView {
  private currentSlide: number = 0;
  private totalSlides: number = 0;
  private zoomLevel: number = 1.0;
  private pdfDoc: pdfjsLib.PDFDocumentProxy | null = null;
  private pageCanvases: HTMLCanvasElement[] = [];
  private currentPdfPath: string | null = null;
  private slideCounter: HTMLElement | null = null;
  private prevButton: HTMLButtonElement | null = null;
  private nextButton: HTMLButtonElement | null = null;
  private contentContainer: HTMLElement | null = null;
  private sidebar: HTMLElement | null = null;
  private mainContent: HTMLElement | null = null;
  private thumbnailContainer: HTMLElement | null = null;
  private mainCanvas: HTMLCanvasElement | null = null;

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
  }

  canAcceptExtension(extension: string): boolean {
    return extension === 'pptx' || extension === 'ppt';
  }

  getViewType(): string {
    return PPTX_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.file?.basename || 'PowerPoint Viewer';
  }

  getIcon(): string {
    return 'presentation';
  }

  async onOpen(): Promise<void> {
    const container = this.containerEl.children[1];
    container.empty();
    container.addClass('pptx-view-container');

    this.createLayout(container);
    this.registerKeyboardHandlers();
  }

  async onLoadFile(file: TFile): Promise<void> {
    await this.loadPptxFile(file);
  }

  async onUnloadFile(file: TFile): Promise<void> {
    this.cleanup();
  }

  private cleanup(): void {
    if (this.currentPdfPath) {
      cleanupPdf(this.currentPdfPath);
      this.currentPdfPath = null;
    }
    this.pdfDoc = null;
    this.pageCanvases = [];
  }

  private createLayout(container: Element): void {
    const layoutContainer = container.createDiv({ cls: 'pptx-layout' });

    // Left sidebar with thumbnails
    this.sidebar = layoutContainer.createDiv({ cls: 'pptx-sidebar' });
    this.thumbnailContainer = this.sidebar.createDiv({ cls: 'pptx-thumbnails' });

    // Right side: toolbar + content
    this.mainContent = layoutContainer.createDiv({ cls: 'pptx-main-content' });
    
    // Top toolbar
    this.createToolbar();
    
    // Main slide content
    this.contentContainer = this.mainContent.createDiv({ cls: 'pptx-content' });
  }

  private createToolbar(): void {
    if (!this.mainContent) return;

    const toolbar = this.mainContent.createDiv({ cls: 'pptx-toolbar' });

    // Left side: navigation
    const navGroup = toolbar.createDiv({ cls: 'pptx-toolbar-group' });
    
    this.prevButton = navGroup.createEl('button', { cls: 'pptx-toolbar-btn', attr: { 'aria-label': 'Previous slide' } });
    this.prevButton.innerHTML = '‹';
    this.prevButton.addEventListener('click', () => this.previousSlide());

    this.slideCounter = navGroup.createDiv({ cls: 'pptx-page-counter', text: '1 of 1' });

    this.nextButton = navGroup.createEl('button', { cls: 'pptx-toolbar-btn', attr: { 'aria-label': 'Next slide' } });
    this.nextButton.innerHTML = '›';
    this.nextButton.addEventListener('click', () => this.nextSlide());

    // Right side: zoom
    const zoomGroup = toolbar.createDiv({ cls: 'pptx-toolbar-group' });

    const zoomOutBtn = zoomGroup.createEl('button', { cls: 'pptx-toolbar-btn', attr: { 'aria-label': 'Zoom out' } });
    zoomOutBtn.innerHTML = '−';
    zoomOutBtn.addEventListener('click', () => this.zoomOut());

    const zoomInBtn = zoomGroup.createEl('button', { cls: 'pptx-toolbar-btn', attr: { 'aria-label': 'Zoom in' } });
    zoomInBtn.innerHTML = '+';
    zoomInBtn.addEventListener('click', () => this.zoomIn());

    this.updateSidebarState();
  }

  private registerKeyboardHandlers(): void {
    this.registerDomEvent(document, 'keydown', (evt: KeyboardEvent) => {
      if (!this.containerEl.isShown()) return;

      switch (evt.key) {
        case 'ArrowUp':
          evt.preventDefault();
          this.previousSlide();
          break;
        case 'ArrowDown':
          evt.preventDefault();
          this.nextSlide();
          break;
        case '+':
        case '=':
          evt.preventDefault();
          this.zoomIn();
          break;
        case '-':
        case '_':
          evt.preventDefault();
          this.zoomOut();
          break;
      }
    });
  }

  async loadPptxFile(file: TFile): Promise<void> {
    this.currentSlide = 0;
    this.zoomLevel = 1.0;

    if (!this.contentContainer) return;

    this.contentContainer.empty();
    this.contentContainer.createDiv({ cls: 'pptx-loading', text: 'Loading presentation...' });

    try {
      // Get the absolute path of the PPTX file
      const vaultPath = (this.app.vault.adapter as any).basePath;
      const pptxPath = `${vaultPath}/${file.path}`;

      // Convert PPTX to PDF (with caching)
      const result = await convertPptxToPdf(pptxPath);

      if (!result.success || !result.pdfPath) {
        this.contentContainer.empty();
        this.contentContainer.createDiv({ 
          cls: 'pptx-error pptx-install-message', 
          text: result.error || 'Failed to convert presentation' 
        });
        return;
      }

      // Log cache status
      if (result.fromCache) {
        console.log('[PPTX View] Loaded from cache:', result.pdfPath);
      } else {
        console.log('[PPTX View] Converted and cached:', result.pdfPath);
      }

      this.currentPdfPath = result.pdfPath;

      // Load the PDF
      await this.loadPdf(result.pdfPath);
    } catch (error: any) {
      this.contentContainer.empty();
      this.contentContainer.createDiv({ 
        cls: 'pptx-error', 
        text: `Failed to load presentation: ${error.message}` 
      });
    }
  }

  private async loadPdf(pdfPath: string): Promise<void> {
    if (!this.contentContainer) return;

    try {
      // Read PDF as binary data (can't use file:// URLs in browser)
      const fs = require('fs');
      const pdfData = fs.readFileSync(pdfPath);
      const pdfUint8Array = new Uint8Array(pdfData);
      
      // Load PDF document from binary data
      this.pdfDoc = await pdfjsLib.getDocument({ data: pdfUint8Array }).promise;
      this.totalSlides = this.pdfDoc.numPages;

      this.contentContainer.empty();
      
      // Create main canvas for current slide
      this.mainCanvas = document.createElement('canvas');
      this.mainCanvas.className = 'pptx-main-canvas';
      this.contentContainer.appendChild(this.mainCanvas);

      // Render first slide
      await this.renderSlide(0);
      
      // Create thumbnails
      await this.createThumbnails();
      
      this.updateSidebarState();
    } catch (error: any) {
      this.contentContainer.empty();
      this.contentContainer.createDiv({ 
        cls: 'pptx-error', 
        text: `Failed to render PDF: ${error.message}` 
      });
    }
  }

  private async renderSlide(pageNum: number): Promise<void> {
    if (!this.pdfDoc || !this.mainCanvas) return;

    const page = await this.pdfDoc.getPage(pageNum + 1); // PDF pages are 1-indexed
    const scale = 2.0 * this.zoomLevel; // Base scale for quality
    const viewport = page.getViewport({ scale });

    this.mainCanvas.width = viewport.width;
    this.mainCanvas.height = viewport.height;
    this.mainCanvas.style.width = `${viewport.width / 2}px`;
    this.mainCanvas.style.height = `${viewport.height / 2}px`;

    const ctx = this.mainCanvas.getContext('2d');
    if (!ctx) return;

    await page.render({
      canvasContext: ctx,
      viewport: viewport,
      canvas: this.mainCanvas
    } as any).promise;
  }

  private async createThumbnails(): Promise<void> {
    if (!this.thumbnailContainer || !this.pdfDoc) return;
    
    this.thumbnailContainer.empty();
    this.pageCanvases = [];
    
    for (let i = 0; i < this.totalSlides; i++) {
      const thumbnail = this.thumbnailContainer.createDiv({ cls: 'pptx-thumbnail' });
      if (i === this.currentSlide) {
        thumbnail.addClass('active');
      }
      
      const preview = thumbnail.createDiv({ cls: 'pptx-thumbnail-preview' });
      
      // Create thumbnail canvas
      const thumbCanvas = document.createElement('canvas');
      preview.appendChild(thumbCanvas);
      this.pageCanvases.push(thumbCanvas);
      
      // Render thumbnail
      await this.renderThumbnail(i, thumbCanvas);
      
      // Slide number overlay
      const slideNum = thumbnail.createDiv({ cls: 'pptx-thumbnail-number' });
      slideNum.setText(`${i + 1}`);
      
      const slideIndex = i;
      thumbnail.addEventListener('click', () => {
        this.goToSlide(slideIndex);
      });
    }
  }

  private async renderThumbnail(pageNum: number, canvas: HTMLCanvasElement): Promise<void> {
    if (!this.pdfDoc) return;

    const page = await this.pdfDoc.getPage(pageNum + 1);
    const targetWidth = 120;
    const scale = (targetWidth / page.getViewport({ scale: 1 }).width) * 2; // 2x for clarity
    const viewport = page.getViewport({ scale });

    canvas.width = viewport.width;
    canvas.height = viewport.height;
    canvas.style.width = `${viewport.width / 2}px`;
    canvas.style.height = `${viewport.height / 2}px`;

    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    await page.render({
      canvasContext: ctx,
      viewport: viewport,
      canvas: canvas
    } as any).promise;
  }

  private async goToSlide(index: number): Promise<void> {
    if (index >= 0 && index < this.totalSlides) {
      this.currentSlide = index;
      await this.renderSlide(index);
      this.updateSidebarState();
      this.updateThumbnailSelection();
    }
  }

  private updateThumbnailSelection(): void {
    if (!this.thumbnailContainer) return;
    
    const thumbnails = this.thumbnailContainer.querySelectorAll('.pptx-thumbnail');
    thumbnails.forEach((thumb, index) => {
      if (index === this.currentSlide) {
        thumb.addClass('active');
      } else {
        thumb.removeClass('active');
      }
    });

    // Scroll thumbnail into view
    const activeThumbnail = thumbnails[this.currentSlide];
    if (activeThumbnail) {
      activeThumbnail.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }

  private async applyZoom(): Promise<void> {
    await this.renderSlide(this.currentSlide);
  }

  private previousSlide(): void {
    if (this.currentSlide > 0) {
      this.goToSlide(this.currentSlide - 1);
    }
  }

  private nextSlide(): void {
    if (this.currentSlide < this.totalSlides - 1) {
      this.goToSlide(this.currentSlide + 1);
    }
  }

  private zoomIn(): void {
    this.zoomLevel = Math.min(this.zoomLevel + 0.25, 3.0);
    this.applyZoom();
  }

  private zoomOut(): void {
    this.zoomLevel = Math.max(this.zoomLevel - 0.25, 0.5);
    this.applyZoom();
  }

  private updateSidebarState(): void {
    if (this.slideCounter) {
      this.slideCounter.setText(`${this.currentSlide + 1} / ${this.totalSlides}`);
    }

    if (this.prevButton) {
      this.prevButton.disabled = this.currentSlide === 0;
    }

    if (this.nextButton) {
      this.nextButton.disabled = this.currentSlide >= this.totalSlides - 1;
    }
  }

  async onClose(): Promise<void> {
    this.cleanup();
    this.file = null;
  }
}
