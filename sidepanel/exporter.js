/**
 * Export utilities for converting sheet data arrays to file formats.
 * Supports CSV, TSV, and XLSX (via SheetJS).
 */

// eslint-disable-next-line no-unused-vars
const Exporter = (() => {
  'use strict';

  /**
   * Convert a 2D data array to a CSV string.
   * Handles quoting for fields containing commas, quotes, or newlines.
   * @param {string[][]} data
   * @returns {string}
   */
  function toCsv(data) {
    return toDelimited(data, ',');
  }

  /**
   * Convert a 2D data array to a TSV string.
   * @param {string[][]} data
   * @returns {string}
   */
  function toTsv(data) {
    return toDelimited(data, '\t');
  }

  /**
   * Convert a 2D data array to a delimited string.
   * @param {string[][]} data
   * @param {string} delimiter
   * @returns {string}
   */
  function toDelimited(data, delimiter) {
    if (!data || data.length === 0) return '';

    return data
      .map((row) =>
        (row || [])
          .map((cell) => {
            const value = cell == null ? '' : String(cell);
            // Quote if contains delimiter, quote, or newline
            if (
              value.includes(delimiter) ||
              value.includes('"') ||
              value.includes('\n') ||
              value.includes('\r')
            ) {
              return '"' + value.replace(/"/g, '""') + '"';
            }
            return value;
          })
          .join(delimiter)
      )
      .join('\r\n');
  }

  /**
   * Convert a 2D data array to an XLSX Blob using SheetJS.
   * When each sheet has a parallel `styles` 2D array (in the shape produced
   * by Parser.extractSheetStyles), those cell-level styles are applied so
   * the file retains its colors, fonts, and other formatting on re-import.
   * @param {Array<{ name: string, data: string[][], styles?: object[][] }>} sheetsData
   * @returns {Blob}
   */
  function toXlsx(sheetsData) {
    if (typeof XLSX === 'undefined') {
      throw new Error('SheetJS (XLSX) library not loaded — cannot export as .xlsx');
    }

    const workbook = XLSX.utils.book_new();

    for (const sheet of sheetsData) {
      const ws = XLSX.utils.aoa_to_sheet(sheet.data || []);

      if (Array.isArray(sheet.styles) && sheet.styles.length > 0) {
        applyCellStyles(ws, sheet.styles);
      }

      XLSX.utils.book_append_sheet(workbook, ws, sheet.name || 'Sheet1');
    }

    const wbOut = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
      cellStyles: true,
    });
    return new Blob([wbOut], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }

  /**
   * Apply a 2D array of SheetJS cell-style objects to a worksheet's cells.
   * Styles parallel to the sheet's data (styles[row][col]) are written as
   * `cell.s` so SheetJS serialises them into the workbook's Styles table.
   * Skips cells that are missing or have no style to apply.
   */
  function applyCellStyles(ws, stylesGrid) {
    if (!ws || !ws['!ref']) return;
    const range = XLSX.utils.decode_range(ws['!ref']);

    const rowCount = stylesGrid.length;
    for (let r = 0; r <= range.e.r && r < rowCount; r++) {
      const styleRow = stylesGrid[r];
      if (!Array.isArray(styleRow)) continue;
      for (let c = 0; c <= range.e.c && c < styleRow.length; c++) {
        const cellStyle = styleRow[c];
        if (!cellStyle || typeof cellStyle !== 'object') continue;
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (cell) cell.s = cellStyle;
      }
    }
  }

  /**
   * Convert sheet data to a Blob in the requested format.
   * @param {Array<{ name: string, data: string[][] }>} sheetsData
   * @param {string} format - 'csv', 'tsv', or 'xlsx'
   * @returns {Blob}
   */
  function toBlob(sheetsData, format) {
    const fmt = (format || 'csv').toLowerCase();

    if (fmt === 'xlsx') {
      return toXlsx(sheetsData);
    }

    // For CSV/TSV, use only the first sheet's data
    const data = sheetsData[0]?.data || [];

    if (fmt === 'tsv') {
      const content = toTsv(data);
      return new Blob([content], { type: 'text/tab-separated-values' });
    }

    // Default: CSV
    const content = toCsv(data);
    return new Blob([content], { type: 'text/csv' });
  }

  /**
   * Derive the output filename given the original name and target format.
   * If the original is "data.csv" and target is "xlsx", returns "data.xlsx".
   * @param {string} originalName
   * @param {string} targetFormat
   * @returns {string}
   */
  function deriveFileName(originalName, targetFormat) {
    const base = (originalName || 'export').replace(/\.[^.]+$/, '');
    const ext = (targetFormat || 'csv').toLowerCase();
    return `${base}.${ext}`;
  }

  return {
    toCsv,
    toTsv,
    toXlsx,
    toBlob,
    deriveFileName,
  };
})();
