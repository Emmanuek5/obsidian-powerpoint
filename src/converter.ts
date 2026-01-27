import { exec } from 'child_process';
import { promisify } from 'util';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';

const execAsync = promisify(exec);

export interface ConversionResult {
  success: boolean;
  pdfPath?: string;
  iframeUrl?: string;
  error?: string;
}

export async function convertPptxToPdf(pptxPath: string): Promise<ConversionResult> {
  const tempDir = os.tmpdir();
  const ext = path.extname(pptxPath);
  const baseName = path.basename(pptxPath, ext);
  const outputPdfPath = path.join(tempDir, `${baseName}_${Date.now()}.pdf`);
  
  // Try LibreOffice first (desktop)
  const libreOfficeResult = await tryLibreOfficeConversion(pptxPath, tempDir, baseName, outputPdfPath);
  if (libreOfficeResult.success) {
    return libreOfficeResult;
  }
  
  // Fallback to Office Online iframe (mobile)
  console.log('[PPTX Converter] LibreOffice not available, trying iframe...');
  return await tryIframeConversion(pptxPath);
}

async function tryLibreOfficeConversion(
  pptxPath: string,
  tempDir: string,
  baseName: string,
  outputPdfPath: string
): Promise<ConversionResult> {
  const libreOfficePaths = getLibreOfficePaths();
  
  for (const loPath of libreOfficePaths) {
    try {
      await execAsync(`"${loPath}" --version`);
      
      const command = `"${loPath}" --headless --convert-to pdf --outdir "${tempDir}" "${pptxPath}"`;
      await execAsync(command);
      
      const generatedPdf = path.join(tempDir, `${baseName}.pdf`);
      
      if (fs.existsSync(generatedPdf)) {
        fs.renameSync(generatedPdf, outputPdfPath);
        return { success: true, pdfPath: outputPdfPath };
      }
    } catch (e) {
      continue;
    }
  }
  
  return { success: false };
}

async function tryIframeConversion(pptxPath: string): Promise<ConversionResult> {
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

export function cleanupPdf(pdfPath: string): void {
  try {
    if (fs.existsSync(pdfPath)) {
      fs.unlinkSync(pdfPath);
    }
  } catch (e) {
    // Ignore cleanup errors
  }
}
