/**
 * Setup script — copies the SheetJS-compatible library into lib/ for the
 * Chrome extension. Run: npm install && npm run setup
 *
 * We use `xlsx-js-style`, a SheetJS fork that adds support for writing
 * cell styles (background color, font color, etc.).
 * The read-side API is identical to the upstream SheetJS Community
 * Edition, so parser and exporter code is unchanged.
 */

const fs = require('fs');
const path = require('path');

const LIB_DIR = path.join(__dirname, 'lib');
// xlsx-js-style ships a single minified bundle; we copy it to the same
// path the side panel loads (lib/xlsx.full.min.js) so the existing
// <script src="../lib/xlsx.full.min.js"> tag continues to work.
const SOURCE = path.join(__dirname, 'node_modules', 'xlsx-js-style', 'dist', 'xlsx.min.js');
const DEST = path.join(LIB_DIR, 'xlsx.full.min.js');
const LUCIDE_SOURCE = path.join(__dirname, 'node_modules', 'lucide', 'dist', 'umd', 'lucide.js');
const LUCIDE_DEST = path.join(LIB_DIR, 'lucide.js');

if (!fs.existsSync(SOURCE)) {
  console.error('xlsx-js-style not found. Run "npm install" first.');
  process.exit(1);
}

if (!fs.existsSync(LIB_DIR)) {
  fs.mkdirSync(LIB_DIR, { recursive: true });
}

fs.copyFileSync(SOURCE, DEST);
console.log(`Copied xlsx-js-style to ${DEST}`);
console.log('Excel support (.xlsx/.xls) with cell formatting is now enabled.');

if (fs.existsSync(LUCIDE_SOURCE)) {
  fs.copyFileSync(LUCIDE_SOURCE, LUCIDE_DEST);
  console.log(`Copied Lucide icons to ${LUCIDE_DEST}`);
} else {
  console.warn('Lucide not found — icons will not be available. Run "npm install" first.');
}
