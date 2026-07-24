const { loadModule } = require('./helpers');

const GoogleAPI = loadModule('../sidepanel/google-api.js', 'GoogleAPI');

describe('GoogleAPI', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    global.fetch = jest.fn();
  });

  // ================================================================
  //  sheetJsToSheetsFormat (pure function)
  // ================================================================

  describe('sheetJsToSheetsFormat', () => {
    test('returns null for null input', () => {
      expect(GoogleAPI.sheetJsToSheetsFormat(null)).toBeNull();
    });

    test('returns null for undefined input', () => {
      expect(GoogleAPI.sheetJsToSheetsFormat(undefined)).toBeNull();
    });

    test('returns null for non-object input', () => {
      expect(GoogleAPI.sheetJsToSheetsFormat('string')).toBeNull();
      expect(GoogleAPI.sheetJsToSheetsFormat(42)).toBeNull();
    });

    test('returns null for empty style object', () => {
      expect(GoogleAPI.sheetJsToSheetsFormat({})).toBeNull();
    });

    // Background color
    test('converts fgColor.rgb to backgroundColor', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ fgColor: { rgb: 'FF0000' } });
      expect(result.backgroundColor).toEqual({ red: 1, green: 0, blue: 0 });
    });

    test('converts fill.fgColor.rgb to backgroundColor', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        fill: { fgColor: { rgb: '00FF00' } },
      });
      expect(result.backgroundColor).toEqual({ red: 0, green: 1, blue: 0 });
    });

    test('handles ARGB hex (8 chars, skips alpha)', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ fgColor: { rgb: 'FF0000FF' } });
      expect(result.backgroundColor).toEqual({ red: 0, green: 0, blue: 1 });
    });

    // Text format
    test('converts bold', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ font: { bold: true } });
      expect(result.textFormat.bold).toBe(true);
    });

    test('converts italic', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ font: { italic: true } });
      expect(result.textFormat.italic).toBe(true);
    });

    test('converts underline', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ font: { underline: true } });
      expect(result.textFormat.underline).toBe(true);
    });

    test('converts strikethrough (strike)', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ font: { strike: true } });
      expect(result.textFormat.strikethrough).toBe(true);
    });

    test('converts strikethrough (strikethrough)', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ font: { strikethrough: true } });
      expect(result.textFormat.strikethrough).toBe(true);
    });

    test('converts font size', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ font: { sz: 14 } });
      expect(result.textFormat.fontSize).toBe(14);
    });

    test('converts font family', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ font: { name: 'Arial' } });
      expect(result.textFormat.fontFamily).toBe('Arial');
    });

    test('converts font color', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        font: { color: { rgb: '0000FF' } },
      });
      expect(result.textFormat.foregroundColor).toEqual({ red: 0, green: 0, blue: 1 });
    });

    test('converts theme font color', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        font: { color: { theme: 1 } },
      });
      expect(result.textFormat.foregroundColor).toEqual({ red: 0, green: 0, blue: 0 });
    });

    test('uses workbook theme colors when provided', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat(
        { font: { color: { theme: 4 } } },
        ['FFFFFF', '000000', 'FFFFFF', '000000', '4285F4']
      );
      expect(result.textFormat.foregroundColor).toEqual({
        red: 66 / 255,
        green: 133 / 255,
        blue: 244 / 255,
      });
    });

    test('converts tinted theme font color', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        font: { color: { theme: 1, tint: 0.5 } },
      });
      expect(result.textFormat.foregroundColor.red).toBeCloseTo(0.5, 2);
      expect(result.textFormat.foregroundColor.green).toBeCloseTo(0.5, 2);
      expect(result.textFormat.foregroundColor.blue).toBeCloseTo(0.5, 2);
    });

    test('converts indexed font color', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        font: { color: { indexed: 4 } },
      });
      expect(result.textFormat.foregroundColor).toEqual({ red: 0, green: 0, blue: 1 });
    });

    test('falls back to default theme colors when themeColors is null', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat(
        { font: { color: { theme: 1 } } },
        null
      );
      expect(result.textFormat.foregroundColor).toEqual({ red: 0, green: 0, blue: 0 });
    });

    test('converts textColor as fallback', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ textColor: 'FF0000' });
      expect(result.textFormat.foregroundColor).toEqual({ red: 1, green: 0, blue: 0 });
    });

    // Bold/italic from root level (without font wrapper)
    test('handles bold at root level', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ bold: true });
      expect(result.textFormat.bold).toBe(true);
    });

    // Alignment
    test('converts horizontal alignment', () => {
      expect(
        GoogleAPI.sheetJsToSheetsFormat({ alignment: { horizontal: 'left' } })
          .horizontalAlignment
      ).toBe('LEFT');
      expect(
        GoogleAPI.sheetJsToSheetsFormat({ alignment: { horizontal: 'center' } })
          .horizontalAlignment
      ).toBe('CENTER');
      expect(
        GoogleAPI.sheetJsToSheetsFormat({ alignment: { horizontal: 'right' } })
          .horizontalAlignment
      ).toBe('RIGHT');
      expect(
        GoogleAPI.sheetJsToSheetsFormat({ alignment: { horizontal: 'justify' } })
          .horizontalAlignment
      ).toBe('LEFT');
    });

    test('converts vertical alignment', () => {
      expect(
        GoogleAPI.sheetJsToSheetsFormat({ alignment: { vertical: 'top' } })
          .verticalAlignment
      ).toBe('TOP');
      expect(
        GoogleAPI.sheetJsToSheetsFormat({ alignment: { vertical: 'center' } })
          .verticalAlignment
      ).toBe('MIDDLE');
      expect(
        GoogleAPI.sheetJsToSheetsFormat({ alignment: { vertical: 'bottom' } })
          .verticalAlignment
      ).toBe('BOTTOM');
    });

    test('converts wrap text', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        alignment: { wrapText: true },
      });
      expect(result.wrapStrategy).toBe('WRAP');
    });

    // Borders
    test('converts border styles', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        border: {
          top: { style: 'thin' },
          bottom: { style: 'medium' },
          left: { style: 'dashed' },
          right: { style: 'double' },
        },
      });
      expect(result.borders.top.style).toBe('SOLID');
      expect(result.borders.bottom.style).toBe('SOLID_MEDIUM');
      expect(result.borders.left.style).toBe('DASHED');
      expect(result.borders.right.style).toBe('DOUBLE');
    });

    test('converts border colors', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        border: {
          top: { style: 'thin', color: { rgb: 'FF0000' } },
        },
      });
      expect(result.borders.top.color).toEqual({ red: 1, green: 0, blue: 0 });
    });

    test('handles hair border style as DOTTED', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        border: { top: { style: 'hair' } },
      });
      expect(result.borders.top.style).toBe('DOTTED');
    });

    // Number format
    test('converts number format', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ numFmt: '#,##0.00' });
      expect(result.numberFormat).toEqual({
        type: 'NUMBER',
        pattern: '#,##0.00',
      });
    });

    test('converts z property as number format', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ z: '0.00%' });
      expect(result.numberFormat).toEqual({
        type: 'NUMBER',
        pattern: '0.00%',
      });
    });

    test('ignores "General" number format', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ numFmt: 'General' });
      expect(result).toBeNull();
    });

    // Combined formats
    test('handles multiple format properties together', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({
        fgColor: { rgb: 'FFFF00' },
        font: { bold: true, sz: 12 },
        alignment: { horizontal: 'center' },
      });
      expect(result.backgroundColor).toEqual({ red: 1, green: 1, blue: 0 });
      expect(result.textFormat.bold).toBe(true);
      expect(result.textFormat.fontSize).toBe(12);
      expect(result.horizontalAlignment).toBe('CENTER');
    });

    // Invalid hex
    test('ignores short/invalid hex colors', () => {
      const result = GoogleAPI.sheetJsToSheetsFormat({ fgColor: { rgb: 'abc' } });
      expect(result).toBeNull();
    });
  });

  // ================================================================
  //  sheetsFormatToSheetJs (inverse direction — used by save flow)
  // ================================================================

  describe('sheetsFormatToSheetJs', () => {
    test('returns null for null input', () => {
      expect(GoogleAPI.sheetsFormatToSheetJs(null)).toBeNull();
    });

    test('returns null for undefined input', () => {
      expect(GoogleAPI.sheetsFormatToSheetJs(undefined)).toBeNull();
    });

    test('returns null for non-object input', () => {
      expect(GoogleAPI.sheetsFormatToSheetJs('string')).toBeNull();
      expect(GoogleAPI.sheetsFormatToSheetJs(42)).toBeNull();
    });

    test('returns null for empty format object', () => {
      expect(GoogleAPI.sheetsFormatToSheetJs({})).toBeNull();
    });

    test('converts backgroundColor to fgColor + fill', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        backgroundColor: { red: 1, green: 0, blue: 0 },
      });
      expect(result.fgColor).toEqual({ rgb: 'FF0000' });
      expect(result.fill).toEqual({
        patternType: 'solid',
        fgColor: { rgb: 'FF0000' },
      });
    });

    test('converts partial backgroundColor to hex', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        backgroundColor: { red: 0, green: 0, blue: 1 },
      });
      expect(result.fgColor).toEqual({ rgb: '0000FF' });
    });

    test('converts textFormat properties to font', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        textFormat: {
          bold: true,
          italic: true,
          underline: true,
          strikethrough: true,
          fontSize: 14,
          fontFamily: 'Arial',
          foregroundColor: { red: 0, green: 0, blue: 1 },
        },
      });
      expect(result.font.bold).toBe(true);
      expect(result.font.italic).toBe(true);
      expect(result.font.underline).toBe(true);
      expect(result.font.strike).toBe(true);
      expect(result.font.sz).toBe(14);
      expect(result.font.name).toBe('Arial');
      expect(result.font.color).toEqual({ rgb: '0000FF' });
    });

    test('ignores zero/invalid font size', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        textFormat: { fontSize: 0 },
      });
      expect(result).toBeNull();
    });

    test('converts horizontal alignment', () => {
      expect(
        GoogleAPI.sheetsFormatToSheetJs({ horizontalAlignment: 'LEFT' }).alignment
          .horizontal
      ).toBe('left');
      expect(
        GoogleAPI.sheetsFormatToSheetJs({ horizontalAlignment: 'CENTER' }).alignment
          .horizontal
      ).toBe('center');
      expect(
        GoogleAPI.sheetsFormatToSheetJs({ horizontalAlignment: 'RIGHT' }).alignment
          .horizontal
      ).toBe('right');
    });

    test('converts vertical alignment', () => {
      expect(
        GoogleAPI.sheetsFormatToSheetJs({ verticalAlignment: 'TOP' }).alignment.vertical
      ).toBe('top');
      expect(
        GoogleAPI.sheetsFormatToSheetJs({ verticalAlignment: 'MIDDLE' }).alignment.vertical
      ).toBe('center');
      expect(
        GoogleAPI.sheetsFormatToSheetJs({ verticalAlignment: 'BOTTOM' }).alignment.vertical
      ).toBe('bottom');
    });

    test('converts WRAP wrapStrategy to wrapText', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({ wrapStrategy: 'WRAP' });
      expect(result.alignment.wrapText).toBe(true);
    });

    test('converts WRAP_AND_CLIP wrapStrategy to wrapText', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({ wrapStrategy: 'WRAP_AND_CLIP' });
      expect(result.alignment.wrapText).toBe(true);
    });

    test('does not set wrapText for OVERFLOW_CELL wrapStrategy', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({ wrapStrategy: 'OVERFLOW_CELL' });
      expect(result).toBeNull();
    });

    test('converts border styles to SheetJS names', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        borders: {
          top: { style: 'SOLID', color: { red: 1, green: 0, blue: 0 } },
          bottom: { style: 'SOLID_MEDIUM' },
          left: { style: 'DASHED' },
          right: { style: 'DOUBLE' },
        },
      });
      expect(result.border.top.style).toBe('thin');
      expect(result.border.top.color).toEqual({ rgb: 'FF0000' });
      expect(result.border.bottom.style).toBe('medium');
      expect(result.border.left.style).toBe('dashed');
      expect(result.border.right.style).toBe('double');
    });

    test('ignores borders with unsupported style', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        borders: { top: { style: 'SOLID' }, bottom: { style: 'UNKNOWN_STYLE' } },
      });
      expect(result.border.top).toBeDefined();
      expect(result.border.bottom).toBeUndefined();
    });

    test('converts number format pattern', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        numberFormat: { type: 'NUMBER', pattern: '0.00%' },
      });
      expect(result.numFmt).toBe('0.00%');
    });

    test('ignores General number format', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        numberFormat: { type: 'NUMBER', pattern: 'General' },
      });
      expect(result).toBeNull();
    });

    test('ignores GENERAL number format', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        numberFormat: { type: 'NUMBER', pattern: 'GENERAL' },
      });
      expect(result).toBeNull();
    });

    test('handles combined format properties', () => {
      const result = GoogleAPI.sheetsFormatToSheetJs({
        backgroundColor: { red: 1, green: 1, blue: 0 },
        textFormat: { bold: true, fontSize: 12 },
        horizontalAlignment: 'CENTER',
      });
      expect(result.fgColor).toEqual({ rgb: 'FFFF00' });
      expect(result.font.bold).toBe(true);
      expect(result.font.sz).toBe(12);
      expect(result.alignment.horizontal).toBe('center');
    });

    test('preserves round-trip with sheetJsToSheetsFormat for color/bold', () => {
      // Start from a SheetJS style and convert it to Sheets format,
      // then convert back. The result should contain the same
      // semantic values.
      const original = { fgColor: { rgb: 'FF8800' }, font: { bold: true, sz: 13 } };
      const sheets = GoogleAPI.sheetJsToSheetsFormat(original);
      const back = GoogleAPI.sheetsFormatToSheetJs(sheets);

      expect(back.fgColor).toEqual({ rgb: 'FF8800' });
      expect(back.font.bold).toBe(true);
      expect(back.font.sz).toBe(13);
    });
  });

  // ================================================================
  //  sheetsFormatGridToSheetJs (2D grid wrapper)
  // ================================================================

  describe('sheetsFormatGridToSheetJs', () => {
    test('returns null for non-array input', () => {
      expect(GoogleAPI.sheetsFormatGridToSheetJs(null)).toBeNull();
      expect(GoogleAPI.sheetsFormatGridToSheetJs(undefined)).toBeNull();
    });

    test('returns null for empty array', () => {
      expect(GoogleAPI.sheetsFormatGridToSheetJs([])).toBeNull();
    });

    test('converts a 2D grid of formats', () => {
      const grid = [
        [
          { backgroundColor: { red: 1, green: 0, blue: 0 } },
          { textFormat: { bold: true } },
        ],
        [null, { backgroundColor: { red: 0, green: 1, blue: 0 } }],
      ];
      const result = GoogleAPI.sheetsFormatGridToSheetJs(grid);
      expect(result).toHaveLength(2);
      expect(result[0][0].fgColor).toEqual({ rgb: 'FF0000' });
      expect(result[0][1].font.bold).toBe(true);
      expect(result[1][0]).toBeNull();
      expect(result[1][1].fgColor).toEqual({ rgb: '00FF00' });
    });

    test('preserves null cells in the grid', () => {
      const grid = [
        [null, null],
        [null, null],
      ];
      const result = GoogleAPI.sheetsFormatGridToSheetJs(grid);
      expect(result[0][0]).toBeNull();
      expect(result[1][1]).toBeNull();
    });
  });

  // ================================================================
  //  getToken
  // ================================================================

  describe('getToken', () => {
    test('returns token from chrome.identity', async () => {
      const token = await GoogleAPI.getToken();
      expect(token).toBe('mock-token');
      expect(chrome.identity.getAuthToken).toHaveBeenCalledWith({ interactive: true });
    });
  });

  // ================================================================
  //  revokeToken
  // ================================================================

  describe('revokeToken', () => {
    test('revokes and removes cached token', async () => {
      global.fetch.mockResolvedValue({});
      await GoogleAPI.revokeToken();

      expect(chrome.identity.getAuthToken).toHaveBeenCalledWith({ interactive: false });
      expect(chrome.identity.removeCachedAuthToken).toHaveBeenCalledWith({
        token: 'mock-token',
      });
      expect(global.fetch).toHaveBeenCalledWith(
        expect.stringContaining('accounts.google.com/o/oauth2/revoke')
      );
    });

    test('does not throw when no cached token', async () => {
      chrome.identity.getAuthToken.mockRejectedValueOnce(new Error('No token'));
      await expect(GoogleAPI.revokeToken()).resolves.toBeUndefined();
    });
  });

  // ================================================================
  //  createSpreadsheet
  // ================================================================

  describe('createSpreadsheet', () => {
    function mockFetchSequence(...responses) {
      for (const resp of responses) {
        global.fetch.mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(resp),
        });
      }
    }

    test('creates spreadsheet and returns id and url', async () => {
      mockFetchSequence(
        // Create spreadsheet
        {
          spreadsheetId: 'sheet-123',
          spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/sheet-123/edit',
          sheets: [{ properties: { sheetId: 0, title: 'Data' } }],
        },
        // Write values
        {},
        // Format (auto-resize)
        {}
      );

      const result = await GoogleAPI.createSpreadsheet('Test', [
        { name: 'Data', data: [['a', 'b'], ['1', '2']] },
      ]);

      expect(result.id).toBe('sheet-123');
      expect(result.url).toBe('https://docs.google.com/spreadsheets/d/sheet-123/edit');
    });

    test('sends correct data to Sheets API', async () => {
      mockFetchSequence(
        {
          spreadsheetId: 's1',
          spreadsheetUrl: 'url',
          sheets: [{ properties: { sheetId: 0, title: 'Sheet1' } }],
        },
        {},
        {}
      );

      await GoogleAPI.createSpreadsheet('My Sheet', [
        { name: 'Sheet1', data: [['Name'], ['Alice']] },
      ]);

      // First call creates the spreadsheet
      const createCall = global.fetch.mock.calls[0];
      expect(createCall[0]).toContain('sheets.googleapis.com');
      const createBody = JSON.parse(createCall[1].body);
      expect(createBody.properties.title).toBe('My Sheet');

      // Second call writes values
      const writeCall = global.fetch.mock.calls[1];
      expect(writeCall[0]).toContain('values:batchUpdate');
    });

    test('writes literal strings with RAW input option', async () => {
      mockFetchSequence(
        {
          spreadsheetId: 's1',
          spreadsheetUrl: 'url',
          sheets: [{ properties: { sheetId: 0, title: 'Sheet1' } }],
        },
        {},
        {}
      );

      await GoogleAPI.createSpreadsheet('My Sheet', [
        { name: 'Sheet1', data: [['Phone'], ['+1-551-848-4656x482']] },
      ]);

      const writeBody = JSON.parse(global.fetch.mock.calls[1][1].body);
      expect(writeBody.valueInputOption).toBe('RAW');
      expect(writeBody.data[0].values[1][0]).toBe('+1-551-848-4656x482');
    });

    test('chunks large value writes into multiple requests', async () => {
      const rows = Array.from({ length: 600 }, (_, index) => [String(index), 'value']);

      mockFetchSequence(
        {
          spreadsheetId: 's1',
          spreadsheetUrl: 'url',
          sheets: [{ properties: { sheetId: 0, title: 'Sheet1' } }],
        },
        {},
        {},
        {}
      );

      await GoogleAPI.createSpreadsheet('Chunked', [
        { name: 'Sheet1', data: rows.map((row) => Array.from({ length: 100 }, () => row[1])) },
      ]);

      const valueCalls = global.fetch.mock.calls.filter((call) =>
        call[0].includes('values:batchUpdate')
      );

      expect(valueCalls).toHaveLength(2);
      expect(JSON.parse(valueCalls[0][1].body).data[0].range).toBe("'Sheet1'!A1");
      expect(JSON.parse(valueCalls[1][1].body).data[0].range).toBe("'Sheet1'!A501");
    });

    test('uses cleaned bounds when tight grid is enabled', async () => {
      mockFetchSequence(
        {
          spreadsheetId: 's1',
          spreadsheetUrl: 'url',
          sheets: [{ properties: { sheetId: 0, title: 'Sheet1' } }],
        },
        {},
        {}
      );

      await GoogleAPI.createSpreadsheet(
        'My Sheet',
        [{ name: 'Sheet1', data: [['Name'], ['Alice']] }],
        { tightGrid: true }
      );

      const createBody = JSON.parse(global.fetch.mock.calls[0][1].body);
      expect(createBody.sheets[0].properties.gridProperties).toEqual({
        rowCount: 2,
        columnCount: 1,
      });
    });

    test('sizes default grid to fit large chunked writes', async () => {
      const rows = Array.from({ length: 2446 }, (_, rowIndex) =>
        Array.from({ length: 30 }, (_, colIndex) => `r${rowIndex}c${colIndex}`)
      );

      mockFetchSequence(
        {
          spreadsheetId: 's1',
          spreadsheetUrl: 'url',
          sheets: [{ properties: { sheetId: 0, title: 'Sheet1' } }],
        },
        {},
        {},
        {}
      );

      await GoogleAPI.createSpreadsheet('Large Sheet', [
        { name: 'Sheet1', data: rows },
      ]);

      const createBody = JSON.parse(global.fetch.mock.calls[0][1].body);
      expect(createBody.sheets[0].properties.gridProperties).toEqual({
        rowCount: 2446,
        columnCount: 30,
      });

      const valueCalls = global.fetch.mock.calls.filter((call) =>
        call[0].includes('values:batchUpdate')
      );
      expect(valueCalls).toHaveLength(2);
      expect(JSON.parse(valueCalls[1][1].body).data[0].range).toBe("'Sheet1'!A1667");
    });

    test('handles empty sheet data', async () => {
      mockFetchSequence(
        {
          spreadsheetId: 's1',
          spreadsheetUrl: 'url',
          sheets: [{ properties: { sheetId: 0, title: 'Empty' } }],
        },
        // No write call for empty data
        {} // Format
      );

      const result = await GoogleAPI.createSpreadsheet('Empty', [
        { name: 'Empty', data: [] },
      ]);
      expect(result.id).toBe('s1');
    });

    test('deduplicates sheet names', async () => {
      mockFetchSequence(
        {
          spreadsheetId: 's1',
          spreadsheetUrl: 'url',
          sheets: [
            { properties: { sheetId: 0, title: 'Data' } },
            { properties: { sheetId: 1, title: 'Data (1)' } },
          ],
        },
        {},
        {}
      );

      await GoogleAPI.createSpreadsheet('Test', [
        { name: 'Data', data: [['A']] },
        { name: 'Data', data: [['B']] },
      ]);

      const body = JSON.parse(global.fetch.mock.calls[0][1].body);
      const sheetNames = body.sheets.map((s) => s.properties.title);
      expect(sheetNames[0]).toBe('Data');
      expect(sheetNames[1]).toBe('Data (1)');
    });

    test('handles API errors', async () => {
      global.fetch.mockResolvedValueOnce({
        ok: false,
        status: 403,
        statusText: 'Forbidden',
        json: () => Promise.resolve({ error: { message: 'Insufficient permissions' } }),
      });

      await expect(
        GoogleAPI.createSpreadsheet('Test', [{ name: 'S', data: [['A']] }])
      ).rejects.toThrow('Google API 403');
    });
  });

  // ================================================================
  //  uploadFileToDrive
  // ================================================================

  describe('uploadFileToDrive', () => {
    test('uploads file and returns id and url', async () => {
      global.fetch.mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ id: 'drive-456' }),
      });

      const file = new File(['data'], 'test.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const result = await GoogleAPI.uploadFileToDrive(file, 'My Sheet');

      expect(result.id).toBe('drive-456');
      expect(result.url).toBe('https://docs.google.com/spreadsheets/d/drive-456/edit');
    });

    test('sends multipart upload to Drive API', async () => {
      global.fetch.mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ id: 'id' }),
      });

      const file = new File(['csv,data'], 'test.csv');
      await GoogleAPI.uploadFileToDrive(file, 'Title');

      const call = global.fetch.mock.calls[0];
      expect(call[0]).toContain('upload/drive/v3/files');
      expect(call[0]).toContain('uploadType=multipart');
      expect(call[1].method).toBe('POST');
    });

    test('throws on upload failure', async () => {
      global.fetch.mockResolvedValueOnce({
        ok: false,
        status: 500,
        statusText: 'Internal Server Error',
        json: () => Promise.resolve({ error: { message: 'Server error' } }),
      });

      const file = new File(['data'], 'test.csv');
      await expect(GoogleAPI.uploadFileToDrive(file, 'T')).rejects.toThrow(
        'Drive upload failed (500)'
      );
    });
  });

  // ---- CellData grid helpers for cleanUploadedSheet tests ----

  function cellS(s, fmt) {
    const c = { userEnteredValue: { stringValue: String(s) } };
    if (fmt) c.effectiveFormat = { numberFormat: { type: fmt } };
    return c;
  }
  function cellF(f) {
    return { userEnteredValue: { formulaValue: f } };
  }
  function cellN(n, fmt) {
    const c = { userEnteredValue: { numberValue: n } };
    if (fmt) c.effectiveFormat = { numberFormat: { type: fmt } };
    return c;
  }
  function cellB(b) {
    return { userEnteredValue: { boolValue: b } };
  }
  function emptyCell() { return {}; }
  function gridData(rows) {
    return { sheets: [{ data: [{ rowData: rows.map((r) => ({ values: r })) }] }] };
  }

  // ================================================================
  //  cleanUploadedSheet
  // ================================================================

  describe('cleanUploadedSheet', () => {
    test('shrinks trailing rows and columns when empty cleanup is enabled', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 1000, columnCount: 26 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [
              ['Name', 'Age', ''],
              ['Alice', '30', ''],
              ['', '', ''],
            ],
          }),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: true,
        removeEmptyColumns: true,
        removeDuplicates: false,
        trim: false,
        fixNumbers: false,
        normalizeHeaders: false,
      });

      const deleteBody = JSON.parse(global.fetch.mock.calls[2][1].body);
      expect(deleteBody.requests.some((request) => request.deleteDimension)).toBe(true);

      const resizeBody = JSON.parse(global.fetch.mock.calls[3][1].body);
      expect(resizeBody.requests[0].updateSheetProperties.properties.gridProperties).toEqual({
        rowCount: 2,
        columnCount: 2,
      });
    });

    test('normalizes headers to title case in value updates', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 2 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [
              ['FIRST NAME', 'eMAIL ADDRESS'],
              ['Alice', 'alice@example.com'],
            ],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('FIRST NAME'), cellS('eMAIL ADDRESS')],
            [cellS('Alice'), cellS('alice@example.com')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: false,
        normalizeHeaders: true,
      });

      const updateBody = JSON.parse(global.fetch.mock.calls[3][1].body);
      expect(updateBody.valueInputOption).toBe('RAW');
      expect(updateBody.data).toEqual([
        { range: "'Sheet1'!A1", values: [['First Name']] },
        { range: "'Sheet1'!B1", values: [['Email Address']] },
      ]);
    });

    test('skips formula cells identified through formulaValue', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 3, columnCount: 2 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [
              ['Name', 'Total'],
              ['Alice', '1,234'],
              ['Bob', 'test'],
            ],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('Name'), cellS('Total')],
            [cellS('Alice'), cellF('=SUM(B2:B2)')],
            [cellS('Bob'), cellF('=A3&" test"')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: true,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      const allCalls = global.fetch.mock.calls;
      expect(allCalls.length).toBe(3); // info, FORMATTED_VALUE read, grid read — no writes
    });

    test('treats literal text beginning with equals as a string, not a formula', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Input'], ['  =not-a-formula']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          // Stored as stringValue, not formulaValue — Sheets treats it as literal text
          json: () => Promise.resolve(gridData([
            [cellS('Input')],
            [cellS('  =not-a-formula')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: true,
        fixNumbers: false,
        normalizeHeaders: false,
      });

      // The =text is a stringValue, not a formulaValue — trim operates on it normally
      const updateBody = JSON.parse(global.fetch.mock.calls[3][1].body);
      expect(updateBody.data).toEqual([
        { range: "'Sheet1'!A2", values: [['=not-a-formula']] },
      ]);
    });

    test('does not convert numeric-looking string in a TEXT-formatted cell', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Code'], ['1,234']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          // numberFormat.type === 'TEXT' — fixNumbers must skip this cell
          json: () => Promise.resolve(gridData([
            [cellS('Code')],
            [cellS('1,234', 'TEXT')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      expect(global.fetch.mock.calls.length).toBe(3); // no value writes
    });

    test('converts eligible numeric-looking string into number', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Amount'], ['1,234']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          // stringValue with no TEXT format — eligible for conversion
          json: () => Promise.resolve(gridData([
            [cellS('Amount')],
            [cellS('1,234')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      const updateBody = JSON.parse(global.fetch.mock.calls[3][1].body);
      expect(updateBody.valueInputOption).toBe('USER_ENTERED');
      expect(updateBody.data).toEqual([
        { range: "'Sheet1'!A2", values: [[1234]] },
      ]);
    });

    test('leaves actual numbers untouched', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Value'], ['42']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          // numberValue — already a number, skip
          json: () => Promise.resolve(gridData([
            [cellS('Value')],
            [cellN(42)],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      expect(global.fetch.mock.calls.length).toBe(3);
    });

    test('leaves DATE cells untouched', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Date'], ['2024-01-15']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          // numberValue with DATE format type
          json: () => Promise.resolve(gridData([
            [cellS('Date')],
            [cellN(45306, 'DATE')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      expect(global.fetch.mock.calls.length).toBe(3);
    });

    test('leaves TIME cells untouched', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Time'], ['14:30:00']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('Time')],
            [cellN(0.6041666666666666, 'TIME')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      expect(global.fetch.mock.calls.length).toBe(3);
    });

    test('leaves DATE_TIME cells untouched', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Timestamp'], ['2024-01-15 14:30:00']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('Timestamp')],
            [cellN(45306.604166666664, 'DATE_TIME')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      expect(global.fetch.mock.calls.length).toBe(3);
    });

    test('leaves boolean cells untouched', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Active'], ['TRUE']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('Active')],
            [cellB(true)],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      expect(global.fetch.mock.calls.length).toBe(3);
    });

    test('preserves leading-zero identifier as string', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Code'], ['00123']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('Code')],
            [cellS('00123')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      expect(global.fetch.mock.calls.length).toBe(3);
    });

    test('preserves leading-zero identifier but cleans commas', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Account'], ['0,012,345']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('Account')],
            [cellS('0,012,345')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      const updateBody = JSON.parse(global.fetch.mock.calls[3][1].body);
      expect(updateBody.valueInputOption).toBe('RAW');
      expect(updateBody.data).toEqual([
        { range: "'Sheet1'!A2", values: [['0012345']] },
      ]);
    });

    test('normalizes headers but skips formula cell in header row', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 3 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [
              ['first name', '=TODAY()', 'eMAIL ADDRESS'],
              ['Alice', '2024-01-15', 'alice@example.com'],
            ],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('first name'), cellF('=TODAY()'), cellS('eMAIL ADDRESS')],
            [cellS('Alice'), cellS('2024-01-15'), cellS('alice@example.com')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: false,
        fixNumbers: false,
        normalizeHeaders: true,
      });

      const updateBody = JSON.parse(global.fetch.mock.calls[3][1].body);
      expect(updateBody.valueInputOption).toBe('RAW');
      expect(updateBody.data).toEqual([
        { range: "'Sheet1'!A1", values: [['First Name']] },
        { range: "'Sheet1'!C1", values: [['Email Address']] },
      ]);
    });

    test('fixNumbers + trim chain correctly on eligible string', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 2, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Price'], ['  1,234.56  ']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('Price')],
            [cellS('  1,234.56  ')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: true,
        fixNumbers: true,
        normalizeHeaders: false,
      });

      const updateBody = JSON.parse(global.fetch.mock.calls[3][1].body);
      expect(updateBody.valueInputOption).toBe('USER_ENTERED');
      expect(updateBody.data).toEqual([
        { range: "'Sheet1'!A2", values: [[1234.56]] },
      ]);
    });

    test('trim updates string cells but not formulas or numbers', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 3, columnCount: 1 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['  Name  '], ['=A1'], ['   text   ']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('  Name  ')],
            [cellF('=A1')],
            [cellS('   text   ')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: true,
        fixNumbers: false,
        normalizeHeaders: false,
      });

      const updateBody = JSON.parse(global.fetch.mock.calls[3][1].body);
      expect(updateBody.valueInputOption).toBe('RAW');
      expect(updateBody.data).toEqual([
        { range: "'Sheet1'!A1", values: [['Name']] },
        { range: "'Sheet1'!A3", values: [['text']] },
      ]);
    });

    test('grid-data request uses an explicit bounded A1 range', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Sheet1',
                gridProperties: { rowCount: 1000, columnCount: 26 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [
              ['A', 'B', 'C'],
              ['1', '2', '3'],
            ],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('A'), cellS('B'), cellS('C')],
            [cellS('1'), cellS('2'), cellS('3')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: true,
        fixNumbers: false,
        normalizeHeaders: false,
      });

      // Call 2 is the grid-data request — inspect its URL
      const gridUrl = global.fetch.mock.calls[2][0];
      expect(gridUrl).toContain('spreadsheets/sheet-id');
      expect(gridUrl).toContain('fields=');
      expect(gridUrl).not.toContain('includeGridData');
      // 2 rows × 3 cols → range should be 'Sheet1'!A1:C2
      expect(gridUrl).toContain(encodeURIComponent("'Sheet1'!A1:C2"));
    });

    test('grid-data range escapes sheet name containing spaces', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'My Data',
                gridProperties: { rowCount: 100, columnCount: 10 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['X'], ['y']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('X')],
            [cellS('y')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: true,
        fixNumbers: false,
        normalizeHeaders: false,
      });

      const gridUrl = global.fetch.mock.calls[2][0];
      // Sheet name with spaces: 'My Data'!A1:A2
      expect(gridUrl).toContain(encodeURIComponent("'My Data'!A1:A2"));
    });

    test('grid-data range escapes sheet name containing an apostrophe', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: "John's Sheet",
                gridProperties: { rowCount: 50, columnCount: 5 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            values: [['Data'], ['1']],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('Data')],
            [cellS('1')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: true,
        fixNumbers: false,
        normalizeHeaders: false,
      });

      const gridUrl = global.fetch.mock.calls[2][0];
      // escapeSheetName wraps in single quotes and doubles internal apostrophes
      const expectedA1 = "'John''s Sheet'!A1:A2";
      expect(gridUrl).toContain(encodeURIComponent(expectedA1));
    });

    test('grid-data range uses max column count from uneven row lengths', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Ragged',
                gridProperties: { rowCount: 10, columnCount: 10 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          // Row 0 has 2 cols, row 1 has 5, row 2 has 3 — max is 5
          json: () => Promise.resolve({
            values: [
              ['A', 'B'],
              ['1', '2', '3', '4', '5'],
              ['x', 'y', 'z'],
            ],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(gridData([
            [cellS('A'), cellS('B')],
            [cellS('1'), cellS('2'), cellS('3'), cellS('4'), cellS('5')],
            [cellS('x'), cellS('y'), cellS('z')],
          ])),
        })
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) });

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: true,
        fixNumbers: false,
        normalizeHeaders: false,
      });

      const gridUrl = global.fetch.mock.calls[2][0];
      // 3 rows × max 5 cols → 'Ragged'!A1:E3
      expect(gridUrl).toContain(encodeURIComponent("'Ragged'!A1:E3"));
    });

    test('effectively empty sheet does not issue a grid-data request', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({
            sheets: [{
              properties: {
                sheetId: 0,
                title: 'Empty',
                gridProperties: { rowCount: 100, columnCount: 10 },
              },
            }],
          }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({ values: [] }),
        });
      // No grid-data mock needed — request should not be issued

      await GoogleAPI.cleanUploadedSheet('sheet-id', {
        removeEmptyRows: false,
        removeEmptyColumns: false,
        removeDuplicates: false,
        trim: true,
        fixNumbers: true,
        normalizeHeaders: true,
      });

      // Only 2 calls: getSpreadsheetInfo + FORMATTED_VALUE read, then skip
      expect(global.fetch.mock.calls.length).toBe(2);
    });
  });

  // ================================================================
  //  applyFormatting
  // ================================================================

  describe('applyFormatting', () => {
    test('does nothing when blocks are empty', async () => {
      await GoogleAPI.applyFormatting('id', []);
      expect(global.fetch).not.toHaveBeenCalled();
    });

    test('does nothing when blocks are null', async () => {
      await GoogleAPI.applyFormatting('id', null);
      expect(global.fetch).not.toHaveBeenCalled();
    });

    test('sends updateCells requests for formatting blocks', async () => {
      // Mock: getSheetInfo + batchUpdate
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () =>
            Promise.resolve({
              sheets: [{ properties: { sheetId: 0 } }],
            }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({}),
        });

      await GoogleAPI.applyFormatting('sheet-id', [
        {
          startRow: 0,
          rows: [[{ backgroundColor: { red: 1, green: 0, blue: 0 } }]],
        },
      ]);

      // Second call is batchUpdate with formatting
      const batchCall = global.fetch.mock.calls[1];
      const body = JSON.parse(batchCall[1].body);
      expect(body.requests[0].updateCells.fields).toBe('userEnteredFormat');
    });
  });

  // ================================================================
  //  formatUploadedSheet
  // ================================================================

  describe('formatUploadedSheet', () => {
    test('auto-resizes columns', async () => {
      global.fetch
        .mockResolvedValueOnce({
          ok: true,
          json: () =>
            Promise.resolve({
              sheets: [
                {
                  properties: {
                    sheetId: 0,
                    gridProperties: { columnCount: 5 },
                  },
                },
              ],
            }),
        })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve({}),
        });

      await GoogleAPI.formatUploadedSheet('sheet-id');

      const body = JSON.parse(global.fetch.mock.calls[1][1].body);
      expect(body.requests[0].autoResizeDimensions).toBeDefined();
      expect(body.requests[0].autoResizeDimensions.dimensions.endIndex).toBe(5);
    });
  });
});
