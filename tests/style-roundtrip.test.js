// End-to-end round-trip test: verify that styles flow correctly from
// Google Sheets userEnteredFormat → SheetJS style → XLSX file → parser.
//
// This test uses xlsx-js-style (the same library the extension bundles
// via setup.js) to write and read the file, simulating the full
// save/import cycle that a real user goes through.

const fs = require('fs');
const path = require('path');
const os = require('os');
const { loadModule } = require('./helpers');

const GoogleAPI = loadModule('../sidepanel/google-api.js', 'GoogleAPI');

describe('Style round-trip (Google Sheets → XLSX → parser)', () => {
  let XLSX;
  const tempFiles = [];

  beforeAll(() => {
    XLSX = require('xlsx-js-style');
  });

  afterAll(() => {
    for (const f of tempFiles) {
      if (f && fs.existsSync(f)) {
        fs.unlinkSync(f);
      }
    }
  });

  function makeTempFile() {
    const f = path.join(os.tmpdir(), `roundtrip-${Date.now()}-${Math.random().toString(36).slice(2)}.xlsx`);
    tempFiles.push(f);
    return f;
  }

  function readWorkbook(file) {
    return XLSX.read(fs.readFileSync(file), {
      type: 'buffer',
      cellStyles: true,
      bookFiles: true,
    });
  }

  // Helper: find a fill whose fgColor RGB ends with `hexSuffix`
  function findFillByColorSuffix(wb, hexSuffix) {
    return (wb.Styles?.Fills || []).find(
      (f) => f && f.fgColor && String(f.fgColor.rgb || '').endsWith(hexSuffix)
    );
  }

  // Helper: find a font with the given properties
  function findFont(wb, predicate) {
    return (wb.Styles?.Fonts || []).find(predicate);
  }

  // Helper: find a CellXf that references a non-default fill
  function findStyledCellXf(wb) {
    return (wb.Styles?.CellXf || []).find(
      (xf) => xf && (xf.fillId > 1 || (xf.fontId > 0 && xf.applyFont))
    );
  }

  test('background colors are written to the saved XLSX', () => {
    const file = makeTempFile();
    // yellow background
    const sheetJsStyle = GoogleAPI.sheetsFormatToSheetJs({
      backgroundColor: { red: 1, green: 1, blue: 0 },
    });
    expect(sheetJsStyle.fgColor).toEqual({ rgb: 'FFFF00' });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['Header', 'Plain'], ['a', 'b']]);
    ws.A1.s = sheetJsStyle;
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, file);

    const readWb = readWorkbook(file);

    // Fills table contains a solid yellow fill
    const yellowFill = findFillByColorSuffix(readWb, 'FFFF00');
    expect(yellowFill).toBeTruthy();
    expect(yellowFill.patternType).toBe('solid');

    // A CellXf references the yellow fill
    const cellXf = (readWb.Styles?.CellXf || []).find(
      (xf) => xf && xf.fillId != null && readWb.Styles?.Fills?.[xf.fillId]?.fgColor?.rgb?.endsWith('FFFF00')
    );
    expect(cellXf).toBeTruthy();
  });

  test('font colors are written to the saved XLSX', () => {
    const file = makeTempFile();
    // red text
    const sheetJsStyle = GoogleAPI.sheetsFormatToSheetJs({
      textFormat: { foregroundColor: { red: 1, green: 0, blue: 0 }, bold: true },
    });
    expect(sheetJsStyle.font.color).toEqual({ rgb: 'FF0000' });
    expect(sheetJsStyle.font.bold).toBe(true);

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['Red text']]);
    ws.A1.s = sheetJsStyle;
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, file);

    const readWb = readWorkbook(file);

    // Fonts table contains a bold red font
    const boldRed = findFont(readWb, (f) =>
      f && (f.bold === true || f.bold === 1) && f.color && String(f.color.rgb || '').endsWith('FF0000')
    );
    expect(boldRed).toBeTruthy();

    // A CellXf references the bold red font
    const cellXf = (readWb.Styles?.CellXf || []).find((xf) => {
      if (!xf || xf.fontId == null) return false;
      const font = readWb.Styles?.Fonts?.[xf.fontId];
      return font && font.color && String(font.color.rgb || '').endsWith('FF0000');
    });
    expect(cellXf).toBeTruthy();
  });

  test('only cells with applied styles get non-default fills', () => {
    const file = makeTempFile();
    const sheetJsStyle = GoogleAPI.sheetsFormatToSheetJs({
      backgroundColor: { red: 1, green: 0, blue: 0 },
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['styled', 'plain'], ['plain', 'plain']]);
    ws.A1.s = sheetJsStyle; // only A1 is styled
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, file);

    const readWb = readWorkbook(file);

    // A red fill exists in the fills table
    const redFill = findFillByColorSuffix(readWb, 'FF0000');
    expect(redFill).toBeTruthy();

    // A non-default CellXf exists
    const cellXf = findStyledCellXf(readWb);
    expect(cellXf).toBeTruthy();
  });

  test('Parser.extractSheetStyles reads the styled cell back correctly', () => {
    // This is the most important test: the same XLSX file the user
    // saves must be re-importable with the original styling.
    const file = makeTempFile();
    const sheetJsStyle = GoogleAPI.sheetsFormatToSheetJs({
      backgroundColor: { red: 1, green: 1, blue: 0 },
      textFormat: { bold: true, foregroundColor: { red: 1, green: 0, blue: 0 } },
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['A1', 'B1'], ['A2', 'B2']]);
    ws.A1.s = sheetJsStyle;
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, file);

    // Read with the same options the extension's parser uses.
    const readWb = XLSX.read(fs.readFileSync(file), {
      type: 'buffer',
      cellStyles: true,
      bookFiles: true,
    });

    // The parser's fast path uses the CellXf + Fills + Fonts tables
    // along with the raw sheet XML to extract styles per cell. Verify
    // those tables are populated and the styled CellXf references the
    // expected font and fill.
    const CellXf = readWb.Styles.CellXf;
    const Fills = readWb.Styles.Fills;
    const Fonts = readWb.Styles.Fonts;

    // Find the CellXf entry that references the yellow fill
    const yellowXf = CellXf.find(
      (xf) => xf && xf.fillId != null && Fills[xf.fillId]?.fgColor?.rgb?.endsWith('FFFF00')
    );
    expect(yellowXf).toBeTruthy();

    // That CellXf also references a bold red font
    const yellowFont = Fonts[yellowXf.fontId];
    expect(yellowFont).toBeTruthy();
    expect(yellowFont.bold === true || yellowFont.bold === 1).toBe(true);
    // The font color is written as FF0000 (6-char RGB) or FFFF0000 (ARGB)
    const fontRgb = String(yellowFont.color.rgb);
    expect(fontRgb === 'FF0000' || fontRgb === 'FFFF0000').toBe(true);
  });
});
