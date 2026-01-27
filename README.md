# Obsidian PowerPoint Viewer Plugin

A powerful Obsidian community plugin that allows you to open and read PowerPoint files (`.pptx` and `.ppt`) directly inside Obsidian with perfect rendering fidelity.

## Features

- **Custom View for PowerPoint Files**: Opens `.pptx` and `.ppt` files inside Obsidian instead of externally
- **Perfect Rendering**: Uses LibreOffice to convert presentations to PDF for pixel-perfect display
- **Intelligent PDF Caching**: Automatically caches converted PDFs based on file hash - subsequent opens are instant!
- **Silent Conversion (Windows)**: No more CMD popup windows during conversion
- **Thumbnail Sidebar**: Slim sidebar with compact slide thumbnails for quick navigation
- **Navigation Toolbar**:
  - Previous/Next slide buttons
  - Slide counter (e.g., "3 / 20")
  - Zoom in/out controls
- **Keyboard Navigation**:
  - `↑` / `↓` : Previous/Next slide
  - `+` / `-` : Zoom in/out
- **Desktop Support**: Works on macOS, Windows, and Linux with LibreOffice installed
- **Clean UI**: Matches Obsidian's design language (inspired by PDF viewer)
- **Automatic Cache Cleanup**: Old cached PDFs (>7 days) are automatically removed

## Requirements

### Desktop (Required for conversion)

**LibreOffice** must be installed for PowerPoint to PDF conversion:

- **macOS**: `brew install --cask libreoffice`
- **Windows**: Download from [libreoffice.org](https://www.libreoffice.org/)
- **Linux**: `sudo apt install libreoffice`

### Development Prerequisites

- Node.js or Bun installed
- Obsidian installed

### Install Dependencies

```bash
bun install
# or
npm install
```

### Build the Plugin

```bash
bun run build
# or
npm run build
```

This creates `main.js` which bundles all required libraries locally.

### Development Mode (Watch)

```bash
bun run dev
# or
npm run dev
```

## Installation in Obsidian

### Method 1: Manual Installation (Development)

1. Build the plugin (see above)
2. Copy these files to your Obsidian vault's plugins folder:
   ```
   <vault>/.obsidian/plugins/obsidian-powerpoint/
   ├── main.js
   ├── manifest.json
   └── styles.css
   ```
3. Open Obsidian Settings → Community Plugins
4. Disable Safe Mode (if enabled)
5. Reload Obsidian or click "Reload plugins"
6. Enable "PowerPoint Viewer" in the plugin list

### Method 2: Symlink (For Development)

```bash
# Navigate to your vault's plugins directory
cd /path/to/your/vault/.obsidian/plugins/

# Create a symlink to this project
ln -s /Users/andrewemmanuel/Documents/Codes/obsidian-powerpoint obsidian-powerpoint

# Reload Obsidian
```

## Usage

1. Install LibreOffice (see Requirements above)
2. Place a `.pptx` or `.ppt` file in your Obsidian vault
3. Click on the file in the file explorer
4. **First time**: The plugin will convert it to PDF (may take a few seconds) and cache it
5. **Subsequent opens**: The cached PDF is loaded instantly - no reconversion needed!
6. The presentation will open with perfect rendering
7. Use the thumbnail sidebar, toolbar buttons, or keyboard shortcuts to navigate

### Keyboard Shortcuts

- **Up Arrow** (`↑`): Previous slide
- **Down Arrow** (`↓`): Next slide
- **Plus** (`+`): Zoom in
- **Minus** (`-`): Zoom out

### PDF Caching System

The plugin uses an intelligent caching system to dramatically improve performance:

- **Hash-based Caching**: Each PowerPoint file is hashed (MD5), and the converted PDF is stored with this hash
- **Instant Loading**: If you open the same file again, the cached PDF is used - no reconversion needed
- **Smart Invalidation**: If you modify the PowerPoint file, the hash changes and a new PDF is generated
- **Cache Location**: PDFs are stored in your system's temp directory under `obsidian-pptx-cache/`
- **Automatic Cleanup**: PDFs older than 7 days are automatically removed when the plugin loads
- **No Manual Management**: The cache is completely transparent - you don't need to do anything!

**Performance Impact**:

- First open: 3-10 seconds (depending on presentation size)
- Subsequent opens: <1 second (instant from cache)

## Testing

1. Download or create a sample `.pptx` file
2. Add it to your Obsidian vault
3. Click on the file to open it
4. Verify:
   - Slides render correctly
   - Navigation buttons work
   - Keyboard shortcuts work
   - Zoom controls work
   - Slide counter updates

## Project Structure

```
obsidian-powerpoint/
├── src/
│   ├── main.ts          # Plugin entry point
│   ├── PptxView.ts      # Custom view implementation
│   └── converter.ts     # PPTX/PPT to PDF conversion
├── manifest.json        # Plugin metadata
├── styles.css           # Plugin styles
├── esbuild.config.mjs   # Build configuration
├── package.json         # Dependencies
└── tsconfig.json        # TypeScript configuration
```

## Architecture

### Conversion Pipeline

1. **File Loading**: User opens `.pptx` or `.ppt` file
2. **LibreOffice Conversion**: File is converted to PDF using LibreOffice headless mode
3. **PDF Rendering**: PDF is rendered using PDF.js with canvas-based rendering
4. **Display**: Slides shown with thumbnail sidebar and navigation controls

### Components

- **`main.ts`**: Registers the custom view and file extension handlers (`.pptx`, `.ppt`)
- **`converter.ts`**: Handles PowerPoint to PDF conversion:
  - Tries LibreOffice conversion (desktop)
  - Provides helpful error messages for mobile
  - Cleans up temporary PDF files
- **`PptxView.ts`**: Implements the slide viewer with:
  - PDF.js rendering on canvas elements
  - Thumbnail sidebar with slide previews
  - Toolbar with navigation controls
  - Keyboard event handlers
  - Zoom functionality
- **`styles.css`**: Obsidian-themed styling matching the PDF viewer design

## Future Enhancements

- Mobile/Android support with web-based conversion
- Slide annotations and highlights
- Export slides as images
- Text extraction from slides
- Note linking to specific slides
- Presentation mode (fullscreen)

## Technical Details

- **Language**: TypeScript
- **Platform**: Obsidian Plugin API
- **Conversion**: LibreOffice (headless mode)
- **Renderer**: PDF.js (canvas-based)
- **Build Tool**: esbuild
- **Target**: ES6, works in Electron environment

## Troubleshooting

### Plugin doesn't appear in Obsidian

- Ensure all three files (`main.js`, `manifest.json`, `styles.css`) are in the plugins folder
- Check Obsidian console (Ctrl+Shift+I / Cmd+Option+I) for errors
- Verify Safe Mode is disabled

### PowerPoint file doesn't open

**Error: "LibreOffice not found"**

- Install LibreOffice (see Requirements section)
- Verify LibreOffice is in your system PATH
- On macOS: Check `/Applications/LibreOffice.app/Contents/MacOS/soffice` exists
- On Windows: Check `C:\Program Files\LibreOffice\program\soffice.exe` exists

**Error: "Failed to render PDF"**

- Check console for detailed error messages
- Ensure the PowerPoint file is not corrupted
- Try with a simple test presentation first
- Verify LibreOffice can open the file directly

**Conversion is slow**

- First conversion creates a temporary PDF (takes a few seconds)
- Subsequent views should be faster
- Large presentations with many images may take longer

### Build fails

- Ensure dependencies are installed: `bun install`
- Check Node.js/Bun version compatibility
- Clear `node_modules` and reinstall

### Mobile/Android Support

Currently, the plugin requires LibreOffice for conversion, which is only available on desktop platforms. Mobile support is planned for future releases using web-based conversion services.

## License

MIT
