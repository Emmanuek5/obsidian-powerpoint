import { exec } from 'child_process';
import { promisify } from 'util';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import * as crypto from 'crypto';

const execAsync = promisify(exec);

export interface ConversionResult {
  success: boolean;
  pdfPath?: string;
  iframeUrl?: string;
  error?: string;
  fromCache?: boolean;
}

// Cache directory for converted PDFs
const CACHE_DIR = path.join(os.tmpdir(), 'obsidian-pptx-cache');

/**
 * Initialize cache directory
 */
function initCacheDir(): void {
  if (!fs.existsSync(CACHE_DIR)) {
    fs.mkdirSync(CACHE_DIR, { recursive: true });
    console.debug('[PPTX Converter] Created cache directory:', CACHE_DIR);
  }
}

/**
 * Calculate MD5 hash of a file for cache key
 */
function getFileHash(filePath: string): string {
  const fileBuffer = fs.readFileSync(filePath);
  const hashSum = crypto.createHash('md5');
  hashSum.update(fileBuffer);
  return hashSum.digest('hex');
}

/**
 * Get cached PDF path if it exists and is valid
 */
function getCachedPdf(pptxPath: string, fileHash: string): string | null {
  initCacheDir();
  
  const ext = path.extname(pptxPath);
  const baseName = path.basename(pptxPath, ext);
  const cachedPdfPath = path.join(CACHE_DIR, `${baseName}_${fileHash}.pdf`);
  
  if (fs.existsSync(cachedPdfPath)) {
    console.debug('[PPTX Converter] Found cached PDF:', cachedPdfPath);
    return cachedPdfPath;
  }
  
  return null;
}

/**
 * Main conversion function with caching support
 */
export async function convertPptxToPdf(pptxPath: string): Promise<ConversionResult> {
  try {
    // Calculate file hash for cache key
    const fileHash = getFileHash(pptxPath);
    
    // Check if we have a cached version
    const cachedPdf = getCachedPdf(pptxPath, fileHash);
    if (cachedPdf) {
      return { 
        success: true, 
        pdfPath: cachedPdf,
        fromCache: true 
      };
    }
    
    // No cache, proceed with conversion
    console.debug('[PPTX Converter] No cache found, converting...');
    initCacheDir();
    
    const ext = path.extname(pptxPath);
    const baseName = path.basename(pptxPath, ext);
    const outputPdfPath = path.join(CACHE_DIR, `${baseName}_${fileHash}.pdf`);
    
    // Try LibreOffice conversion (desktop)
    const libreOfficeResult = await tryLibreOfficeConversion(pptxPath, CACHE_DIR, baseName, outputPdfPath, fileHash);
    if (libreOfficeResult.success) {
      return libreOfficeResult;
    }
    
    // Fallback to Office Online iframe (mobile)
    console.debug('[PPTX Converter] LibreOffice not available, trying iframe...');
    return tryIframeConversion(pptxPath);
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    console.error('[PPTX Converter] Error during conversion:', error);
    return {
      success: false,
      error: `Conversion failed: ${errorMessage}`
    };
  }
}

async function tryLibreOfficeConversion(
  pptxPath: string,
  cacheDir: string,
  baseName: string,
  outputPdfPath: string,
  fileHash: string
): Promise<ConversionResult> {
  const libreOfficePaths = getLibreOfficePaths();
  
  for (const loPath of libreOfficePaths) {
    try {
      // Test if LibreOffice is available (silent on Windows)
      await execAsync(`"${loPath}" --version`, { 
        windowsHide: true 
      });
      
      console.debug(`[PPTX Converter] Using LibreOffice at: ${loPath}`);
      
      // Method 1: Try print-to-file first (better overflow handling)
      // This method uses the print subsystem which can better handle content
      // that extends beyond slide boundaries
      // Note: LibreOffice 25.8+ uses --print-to-file --outdir syntax
      
      try {
        console.debug('[PPTX Converter] Trying print-to-file method (better overflow handling)...');
        // Correct syntax: --print-to-file --outdir {dir} {file}
        // Output will be named same as input but with .pdf extension
        const printCommand = `"${loPath}" --headless --print-to-file --outdir "${cacheDir}" "${pptxPath}"`;
        await execAsync(printCommand, { 
          windowsHide: true,
          timeout: 90000
        });
        
        // LibreOffice names the output file same as input but with .pdf
        const printOutputPath = path.join(cacheDir, `${baseName}.pdf`);
        
        if (fs.existsSync(printOutputPath)) {
          fs.renameSync(printOutputPath, outputPdfPath);
          console.debug('[PPTX Converter] Print-to-file conversion successful');
          return { 
            success: true, 
            pdfPath: outputPdfPath,
            fromCache: false 
          };
        }
      } catch {
        console.debug('[PPTX Converter] Print method failed, trying convert method...');
      }
      
      // Method 2: Fall back to convert-to with maximum quality settings
      // Using impress_pdf_Export filter with options to preserve content:
      // - ExportNotesPages=false: Don't export notes
      // - Quality=100: Maximum quality
      // - ReduceImageResolution=false: Keep full image resolution
      // - MaxImageResolution=600: High DPI for sharp images
      // - EmbedStandardFonts=true: Embed all fonts to prevent reflow
      // - UseLosslessCompression=true: No data loss
      // - IsSkipEmptyPages=false: Keep all pages
      // - ExportBookmarks=true: Preserve structure
      // - OpenBookmarkLevels=-1: All bookmarks expanded
      const filterData = [
        'ExportNotesPages=false',
        'Quality=100',
        'ReduceImageResolution=false',
        'MaxImageResolution=600',
        'EmbedStandardFonts=true',
        'UseLosslessCompression=true',
        'IsSkipEmptyPages=false',
        'ExportBookmarks=true',
        'OpenBookmarkLevels=-1',
        'ExportFormFields=true',
        'SelectPdfVersion=1'  // PDF 1.5 for better compatibility
      ].join(':');
      
      const command = `"${loPath}" --headless --convert-to "pdf:impress_pdf_Export:${filterData}" --outdir "${cacheDir}" "${pptxPath}"`;

      console.debug('[PPTX Converter] Converting with enhanced PDF export settings...');
      await execAsync(command, { 
        windowsHide: true,
        timeout: 90000  // 90 second timeout for large presentations
      });
      
      // LibreOffice generates PDF with original basename
      const generatedPdf = path.join(cacheDir, `${baseName}.pdf`);
      
      if (fs.existsSync(generatedPdf)) {
        // Rename to include hash for caching
        fs.renameSync(generatedPdf, outputPdfPath);
        console.debug('[PPTX Converter] Conversion successful, cached at:', outputPdfPath);
        return {
          success: true,
          pdfPath: outputPdfPath,
          fromCache: false
        };
      }
    } catch (e) {
      console.error(`[PPTX Converter] Failed with ${loPath}:`, e instanceof Error ? e.message : String(e));
      continue;
    }
  }
  
  return { success: false };
}

function tryIframeConversion(pptxPath: string): ConversionResult {
  // Office Online iframe requires a publicly accessible URL
  // For local files, we'd need to serve them via HTTP
  // For now, return error explaining the limitation
  return {
    success: false,
    error: `LibreOffice not found.\n\nDesktop: Install LibreOffice for offline conversion:\n  brew install --cask libreoffice\n\nMobile/Android: Local file viewing not supported yet.\n\nTo view PPTX files on mobile, you can:\n1. Open the file directly in a PPTX viewer app\n2. Upload to a cloud service and view online\n\nWe're working on mobile support!`
  };
}

function getLibreOfficePaths(): string[] {
  const platform = process.platform;
  
  if (platform === 'darwin') {
    return [
      '/Applications/LibreOffice.app/Contents/MacOS/soffice',
      '/usr/local/bin/soffice',
      'soffice'
    ];
  } else if (platform === 'win32') {
    return [
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
      'soffice'
    ];
  } else {
    return [
      '/usr/bin/soffice',
      '/usr/bin/libreoffice',
      'soffice',
      'libreoffice'
    ];
  }
}

/**
 * Clean up old cached PDFs (older than specified days)
 */
export function cleanupOldCache(daysOld: number = 7): void {
  try {
    if (!fs.existsSync(CACHE_DIR)) {
      return;
    }
    
    const now = Date.now();
    const maxAge = daysOld * 24 * 60 * 60 * 1000; // Convert days to milliseconds
    
    const files = fs.readdirSync(CACHE_DIR);
    let deletedCount = 0;
    
    for (const file of files) {
      const filePath = path.join(CACHE_DIR, file);
      const stats = fs.statSync(filePath);
      
      if (now - stats.mtimeMs > maxAge) {
        fs.unlinkSync(filePath);
        deletedCount++;
      }
    }
    
    if (deletedCount > 0) {
      console.debug(`[PPTX Converter] Cleaned up ${deletedCount} old cached PDF(s)`);
    }
  } catch (e) {
    console.error('[PPTX Converter] Error cleaning cache:', e);
  }
}

/**
 * Clear entire cache directory
 */
export function clearCache(): void {
  try {
    if (fs.existsSync(CACHE_DIR)) {
      const files = fs.readdirSync(CACHE_DIR);
      for (const file of files) {
        fs.unlinkSync(path.join(CACHE_DIR, file));
      }
      console.debug('[PPTX Converter] Cache cleared');
    }
  } catch (e) {
    console.error('[PPTX Converter] Error clearing cache:', e);
  }
}

/**
 * Get cache statistics
 */
export function getCacheStats(): { count: number; totalSize: number; path: string } {
  try {
    if (!fs.existsSync(CACHE_DIR)) {
      return { count: 0, totalSize: 0, path: CACHE_DIR };
    }
    
    const files = fs.readdirSync(CACHE_DIR);
    let totalSize = 0;
    
    for (const file of files) {
      const stats = fs.statSync(path.join(CACHE_DIR, file));
      totalSize += stats.size;
    }
    
    return {
      count: files.length,
      totalSize,
      path: CACHE_DIR
    };
  } catch {
    return { count: 0, totalSize: 0, path: CACHE_DIR };
  }
}

/**
 * Legacy cleanup function - now just cleans old cache
 * @deprecated Use cleanupOldCache instead
 */
export function cleanupPdf(pdfPath: string): void {
  // Don't delete cached PDFs anymore - they're managed by cleanupOldCache
  // This function is kept for backward compatibility
  console.debug('[PPTX Converter] PDF cleanup is now handled by cache management');
}
