const { loadModule } = require('./helpers');

const Exporter = loadModule('../sidepanel/exporter.js', 'Exporter');

describe('Exporter', () => {
  let originalXLSX;

  beforeEach(() => {
    originalXLSX = global.XLSX;
  });

  afterEach(() => {
    if (originalXLSX === undefined) {
      delete global.XLSX;
    } else {
      global.XLSX = originalXLSX;
    }
  });

  describe('deriveFileName', () => {
    test('swaps extension to target format', () => {
      expect(Exporter.deriveFileName('data.csv', 'xlsx')).toBe('data.xlsx');
      expect(Exporter.deriveFileName('sales.tsv', 'csv')).toBe('sales.csv');
    });

    test('falls back to default name when no original', () => {
      expect(Exporter.deriveFileName('', 'csv')).toBe('export.csv');
    });
  });

  describe('toCsv', () => {
    test('returns empty string for empty data', () => {
      expect(Exporter.toCsv([])).toBe('');
      expect(Exporter.toCsv(null)).toBe('');
    });

    test('joins rows with CRLF and cells with commas', () => {
      expect(Exporter.toCsv([['a', 'b'], ['1', '2']])).toBe('a,b\r\n1,2');
    });

    test('quotes values that contain delimiters, quotes, or newlines', () => {
      expect(Exporter.toCsv([['has,comma']])).toBe('"has,comma"');
      expect(Exporter.toCsv([['has"quote']])).toBe('"has""quote"');
      expect(Exporter.toCsv([['has\nnewline']])).toBe('"has\nnewline"');
    });

    test('handles null and undefined cells as empty strings', () => {
      expect(Exporter.toCsv([[null, undefined, 'x']])).toBe(',,x');
    });
  });

  describe('toTsv', () => {
    test('uses tab delimiter', () => {
      expect(Exporter.toTsv([['a', 'b'], ['1', '2']])).toBe('a\tb\r\n1\t2');
    });
  });

  describe('toBlob', () => {
    test('returns CSV blob for csv format with text/csv MIME type', () => {
      const blob = Exporter.toBlob(
        [{ name: 'S', data: [['a', 'b']] }],
        'csv'
      );
      expect(blob).toBeInstanceOf(Blob);
      expect(blob.type).toBe('text/csv');
    });

    test('returns TSV blob for tsv format with text/tab-separated-values MIME', () => {
      const blob = Exporter.toBlob(
        [{ name: 'S', data: [['a', 'b']] }],
        'tsv'
      );
      expect(blob).toBeInstanceOf(Blob);
      expect(blob.type).toBe('text/tab-separated-values');
    });

    test('returns xlsx blob via toXlsx for xlsx format', () => {
      const saved = global.XLSX;
      global.XLSX = {
        utils: {
          book_new: jest.fn(() => ({})),
          aoa_to_sheet: jest.fn(() => ({ '!ref': 'A1' })),
          book_append_sheet: jest.fn(),
          decode_range: jest.fn(() => ({ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } })),
          encode_cell: jest.fn(() => 'A1'),
        },
        write: jest.fn(() => new Uint8Array([1, 2, 3])),
      };
      try {
        const blob = Exporter.toBlob(
          [{ name: 'S', data: [['a', 'b']] }],
          'xlsx'
        );
        expect(blob).toBeInstanceOf(Blob);
        expect(blob.type).toBe(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        expect(global.XLSX.write).toHaveBeenCalled();
      } finally {
        if (saved === undefined) {
          delete global.XLSX;
        } else {
          global.XLSX = saved;
        }
      }
    });
  });

  describe('toXlsx', () => {
    function makeSheet(name, data, styles) {
      const sheet = { name, data };
      if (styles) sheet.styles = styles;
      return sheet;
    }

    test('throws when SheetJS is not loaded', () => {
      const saved = global.XLSX;
      delete global.XLSX;
      try {
        expect(() => Exporter.toXlsx([{ name: 'S', data: [] }])).toThrow(/SheetJS/);
      } finally {
        global.XLSX = saved;
      }
    });

    test('writes workbook with one sheet per entry', async () => {
      const appended = [];
      global.XLSX = {
        utils: {
          book_new: jest.fn(() => ({})),
          aoa_to_sheet: jest.fn(() => ({ '!ref': 'A1:B2' })),
          book_append_sheet: jest.fn((wb, ws, name) => {
            appended.push({ wb, ws, name });
          }),
          decode_range: jest.fn(() => ({ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } })),
          encode_cell: jest.fn(({ r, c }) => {
            const col = String.fromCharCode(65 + c);
            return `${col}${r + 1}`;
          }),
        },
        write: jest.fn(() => new Uint8Array([1, 2, 3])),
      };

      await Exporter.toXlsx([
        makeSheet('A', [['a', 'b']]),
        makeSheet('B', [['1']]),
      ]);

      expect(appended).toHaveLength(2);
      expect(appended[0].name).toBe('A');
      expect(appended[1].name).toBe('B');
      expect(global.XLSX.write).toHaveBeenCalledWith(
        expect.anything(),
        expect.objectContaining({ bookType: 'xlsx', cellStyles: true })
      );
    });

    test('writes with cellStyles option even without styles', async () => {
      global.XLSX = {
        utils: {
          book_new: jest.fn(() => ({})),
          aoa_to_sheet: jest.fn(() => ({ '!ref': 'A1' })),
          book_append_sheet: jest.fn(),
          decode_range: jest.fn(() => ({ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } })),
          encode_cell: jest.fn(() => 'A1'),
        },
        write: jest.fn(() => new Uint8Array()),
      };

      await Exporter.toXlsx([makeSheet('S', [['x']])]);

      const writeOpts = global.XLSX.write.mock.calls[0][1];
      expect(writeOpts.cellStyles).toBe(true);
    });

    test('applies styles to cells when present', async () => {
      const sheet = { '!ref': 'A1:B2' };
      global.XLSX = {
        utils: {
          book_new: jest.fn(() => ({})),
          aoa_to_sheet: jest.fn(() => sheet),
          book_append_sheet: jest.fn(),
          decode_range: jest.fn(() => ({ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } })),
          encode_cell: jest.fn(({ r, c }) => {
            const col = String.fromCharCode(65 + c);
            return `${col}${r + 1}`;
          }),
        },
        write: jest.fn(() => new Uint8Array()),
      };

      // Pre-populate cells as aoa_to_sheet would
      sheet.A1 = { v: 'Name' };
      sheet.B1 = { v: 'Age' };
      sheet.A2 = { v: 'Alice' };
      sheet.B2 = { v: 30 };

      const styles = [
        [
          { fgColor: { rgb: 'FF0000' }, font: { bold: true } },
          null,
        ],
        [null, { font: { color: { rgb: '0000FF' } } }],
      ];

      await Exporter.toXlsx([makeSheet('S', [['Name', 'Age'], ['Alice', 30]], styles)]);

      expect(sheet.A1.s).toEqual({ fgColor: { rgb: 'FF0000' }, font: { bold: true } });
      expect(sheet.B1.s).toBeUndefined();
      expect(sheet.A2.s).toBeUndefined();
      expect(sheet.B2.s).toEqual({ font: { color: { rgb: '0000FF' } } });
    });

    test('skips styles for cells that are missing in the worksheet', async () => {
      // Worksheet only has one column, but styles grid has two columns
      const sheet = { '!ref': 'A1:A2' };
      global.XLSX = {
        utils: {
          book_new: jest.fn(() => ({})),
          aoa_to_sheet: jest.fn(() => sheet),
          book_append_sheet: jest.fn(),
          decode_range: jest.fn(() => ({ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } })),
          encode_cell: jest.fn(({ r, c }) => {
            const col = String.fromCharCode(65 + c);
            return `${col}${r + 1}`;
          }),
        },
        write: jest.fn(() => new Uint8Array()),
      };

      sheet.A1 = { v: 'a' };
      sheet.A2 = { v: 'b' };

      const styles = [
        [{ fgColor: { rgb: 'FF0000' } }, { fgColor: { rgb: '00FF00' } }],
        [null, null],
      ];

      await Exporter.toXlsx([makeSheet('S', [['a'], ['b']], styles)]);

      expect(sheet.A1.s).toEqual({ fgColor: { rgb: 'FF0000' } });
      expect(sheet.A2.s).toBeUndefined();
    });

    test('ignores non-object style entries', async () => {
      const sheet = { '!ref': 'A1' };
      global.XLSX = {
        utils: {
          book_new: jest.fn(() => ({})),
          aoa_to_sheet: jest.fn(() => sheet),
          book_append_sheet: jest.fn(),
          decode_range: jest.fn(() => ({ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } })),
          encode_cell: jest.fn(() => 'A1'),
        },
        write: jest.fn(() => new Uint8Array()),
      };

      sheet.A1 = { v: 'x' };

      await Exporter.toXlsx([
        makeSheet('S', [['x']], [[null, 'invalid', undefined, { fgColor: { rgb: 'FF0000' } }]]),
      ]);

      // Only valid style objects are applied. The first 3 cells in the row
      // don't exist (range is A1 only), so they're skipped silently.
      expect(sheet.A1.s).toBeUndefined();
    });
  });
});
