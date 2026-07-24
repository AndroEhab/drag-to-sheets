const { loadModule } = require('./helpers');

const Merger = loadModule('../sidepanel/merger.js', 'Merger');

/** Helper to wrap data in the parsedFiles format that Merger.merge expects */
function makeParsedFile(headers, ...rows) {
  return { sheets: [{ name: 'Sheet', data: [headers, ...rows] }] };
}

describe('Merger', () => {
  // ---- Empty / edge cases ----

  describe('edge cases', () => {
    test('returns empty merged sheet for empty input', () => {
      const result = Merger.merge([]);
      expect(result.sheets).toHaveLength(1);
      expect(result.sheets[0].name).toBe('Merged');
      expect(result.sheets[0].data).toEqual([]);
      expect(result.sourceMap).toEqual([]);
    });

    test('handles files with empty sheets', () => {
      const result = Merger.merge([{ sheets: [{ name: 'Empty', data: [] }] }]);
      expect(result.sheets[0].data).toEqual([]);
    });

    test('filters out files with no data', () => {
      const result = Merger.merge([
        { sheets: [{ name: 'A', data: [] }] },
        makeParsedFile(['Name'], ['Alice']),
      ]);
      expect(result.sheets[0].data).toEqual([['Name'], ['Alice']]);
    });
  });

  // ---- Single file ----

  describe('single file', () => {
    test('returns the file data as merged', () => {
      const file = makeParsedFile(['Name', 'Age'], ['Alice', '30'], ['Bob', '25']);
      const result = Merger.merge([file]);

      expect(result.sheets[0].name).toBe('Merged');
      expect(result.sheets[0].data).toEqual([
        ['Name', 'Age'],
        ['Alice', '30'],
        ['Bob', '25'],
      ]);
    });

    test('creates identity sourceMap for single file', () => {
      const file = makeParsedFile(['A', 'B'], ['1', '2']);
      const result = Merger.merge([file], { includeSourceMap: true });

      expect(result.sourceMap).toHaveLength(2);
      expect(result.sourceMap[0]).toEqual({
        fileIndex: 0,
        sourceRow: 0,
        colMap: [0, 1],
      });
      expect(result.sourceMap[1]).toEqual({
        fileIndex: 0,
        sourceRow: 1,
        colMap: [0, 1],
      });
    });
  });

  // ---- Two files with same headers ----

  describe('same headers', () => {
    test('concatenates rows from both files', () => {
      const file1 = makeParsedFile(['Name', 'Age'], ['Alice', '30']);
      const file2 = makeParsedFile(['Name', 'Age'], ['Bob', '25']);
      const result = Merger.merge([file1, file2]);

      expect(result.sheets[0].data).toEqual([
        ['Name', 'Age'],
        ['Alice', '30'],
        ['Bob', '25'],
      ]);
    });

    test('uses single unified header row', () => {
      const file1 = makeParsedFile(['A', 'B'], ['1', '2']);
      const file2 = makeParsedFile(['A', 'B'], ['3', '4']);
      const result = Merger.merge([file1, file2]);

      // Only one header row
      expect(result.sheets[0].data).toHaveLength(3);
      expect(result.sheets[0].data[0]).toEqual(['A', 'B']);
    });
  });

  // ---- Two files with different headers ----

  describe('different headers', () => {
    test('creates unified header with all columns', () => {
      const file1 = makeParsedFile(['Name', 'Age'], ['Alice', '30']);
      const file2 = makeParsedFile(['Name', 'City'], ['Bob', 'NYC']);
      const result = Merger.merge([file1, file2]);

      expect(result.sheets[0].data[0]).toEqual(['Name', 'Age', 'City']);
    });

    test('fills missing columns with empty strings', () => {
      const file1 = makeParsedFile(['Name', 'Age'], ['Alice', '30']);
      const file2 = makeParsedFile(['Name', 'City'], ['Bob', 'NYC']);
      const result = Merger.merge([file1, file2]);

      // Alice has no City
      expect(result.sheets[0].data[1]).toEqual(['Alice', '30', '']);
      // Bob has no Age
      expect(result.sheets[0].data[2]).toEqual(['Bob', '', 'NYC']);
    });

    test('handles completely non-overlapping headers', () => {
      const file1 = makeParsedFile(['A', 'B'], ['1', '2']);
      const file2 = makeParsedFile(['C', 'D'], ['3', '4']);
      const result = Merger.merge([file1, file2]);

      expect(result.sheets[0].data[0]).toEqual(['A', 'B', 'C', 'D']);
      expect(result.sheets[0].data[1]).toEqual(['1', '2', '', '']);
      expect(result.sheets[0].data[2]).toEqual(['', '', '3', '4']);
    });
  });

  // ---- Case-insensitive matching ----

  describe('case-insensitive header matching', () => {
    test('matches headers regardless of case', () => {
      const file1 = makeParsedFile(['Name', 'AGE'], ['Alice', '30']);
      const file2 = makeParsedFile(['name', 'age'], ['Bob', '25']);
      const result = Merger.merge([file1, file2]);

      // Should use first-seen display names
      expect(result.sheets[0].data[0]).toEqual(['Name', 'AGE']);
      expect(result.sheets[0].data[1]).toEqual(['Alice', '30']);
      expect(result.sheets[0].data[2]).toEqual(['Bob', '25']);
    });

    test('matches headers with extra whitespace', () => {
      const file1 = makeParsedFile(['First  Name'], ['Alice']);
      const file2 = makeParsedFile(['first name'], ['Bob']);
      const result = Merger.merge([file1, file2]);

      // Unified header uses first-seen display name
      expect(result.sheets[0].data[0]).toEqual(['First  Name']);
      expect(result.sheets[0].data).toHaveLength(3);
    });
  });

  // ---- sourceMap ----

  describe('sourceMap', () => {
    test('tracks file origins for each row', () => {
      const file1 = makeParsedFile(['A'], ['1']);
      const file2 = makeParsedFile(['A'], ['2']);
      const result = Merger.merge([file1, file2], { includeSourceMap: true });

      // Header row comes from file 0
      expect(result.sourceMap[0].fileIndex).toBe(0);
      // First data row from file 0
      expect(result.sourceMap[1].fileIndex).toBe(0);
      expect(result.sourceMap[1].sourceRow).toBe(1);
      // Second data row from file 1
      expect(result.sourceMap[2].fileIndex).toBe(1);
      expect(result.sourceMap[2].sourceRow).toBe(1);
    });

    test('preserves original file index when files are filtered', () => {
      const empty = { sheets: [{ name: 'E', data: [] }] };
      const file = makeParsedFile(['A'], ['1']);
      const result = Merger.merge([empty, file], { includeSourceMap: true });

      // The file with data was at index 1 in the original array
      expect(result.sourceMap[0].fileIndex).toBe(1);
    });

    test('includes column mapping in sourceMap', () => {
      const file1 = makeParsedFile(['A', 'B'], ['1', '2']);
      const file2 = makeParsedFile(['B', 'C'], ['3', '4']);
      const result = Merger.merge([file1, file2], { includeSourceMap: true });

      // file2's colMap: B→1, C→2 (in unified order A,B,C)
      const file2Map = result.sourceMap[2].colMap;
      expect(file2Map).toEqual([1, 2]);
    });
  });

  // ---- Multiple files ----

  describe('multiple files', () => {
    test('merges three files correctly', () => {
      const f1 = makeParsedFile(['Name'], ['Alice']);
      const f2 = makeParsedFile(['Name'], ['Bob']);
      const f3 = makeParsedFile(['Name'], ['Charlie']);
      const result = Merger.merge([f1, f2, f3]);

      expect(result.sheets[0].data).toEqual([
        ['Name'],
        ['Alice'],
        ['Bob'],
        ['Charlie'],
      ]);
    });

    test('includes all rows from multi-row files', () => {
      const f1 = makeParsedFile(['X'], ['1'], ['2']);
      const f2 = makeParsedFile(['X'], ['3'], ['4'], ['5']);
      const result = Merger.merge([f1, f2]);

      expect(result.sheets[0].data).toHaveLength(6); // header + 5 data rows
    });

    test('uses first sheet from each file', () => {
      const file = {
        sheets: [
          { name: 'First', data: [['A'], ['1']] },
          { name: 'Second', data: [['B'], ['2']] },
        ],
      };
      const result = Merger.merge([file]);

      // Only uses the first sheet
      expect(result.sheets[0].data).toEqual([['A'], ['1']]);
    });
  });

  // ---- Preserves first-seen header names ----

  describe('header display names', () => {
    test('uses first-seen header display name', () => {
      const f1 = makeParsedFile(['First Name'], ['Alice']);
      const f2 = makeParsedFile(['first name'], ['Bob']);
      const result = Merger.merge([f1, f2]);

      expect(result.sheets[0].data[0]).toEqual(['First Name']);
    });
  });

  // ---- Smart mapping (fuzzy) ----

  describe('smart mapping', () => {
    test('matches underscored and spaced headers', () => {
      const f1 = makeParsedFile(['first_name'], ['Alice']);
      const f2 = makeParsedFile(['First Name'], ['Bob']);
      const result = Merger.merge([f1, f2], { smartMapping: true });

      expect(result.sheets[0].data[0]).toEqual(['first_name']);
      expect(result.sheets[0].data[1]).toEqual(['Alice']);
      expect(result.sheets[0].data[2]).toEqual(['Bob']);
    });

    test('matches hyphenated and spaced headers', () => {
      const f1 = makeParsedFile(['last-name'], ['Smith']);
      const f2 = makeParsedFile(['Last Name'], ['Jones']);
      const result = Merger.merge([f1, f2], { smartMapping: true });

      expect(result.sheets[0].data[0]).toEqual(['last-name']);
      expect(result.sheets[0].data).toHaveLength(3);
    });

    test('matches singular and plural headers (trailing s)', () => {
      const f1 = makeParsedFile(['email'], ['a@b.com']);
      const f2 = makeParsedFile(['emails'], ['c@d.com']);
      const result = Merger.merge([f1, f2], { smartMapping: true });

      expect(result.sheets[0].data[0]).toEqual(['email']);
      expect(result.sheets[0].data[1]).toEqual(['a@b.com']);
      expect(result.sheets[0].data[2]).toEqual(['c@d.com']);
    });

    test('matches ies/y plural forms', () => {
      const f1 = makeParsedFile(['category'], ['A']);
      const f2 = makeParsedFile(['categories'], ['B']);
      const result = Merger.merge([f1, f2], { smartMapping: true });

      expect(result.sheets[0].data[0]).toEqual(['category']);
      expect(result.sheets[0].data).toHaveLength(3);
    });

    test('matches es plural forms (boxes/box)', () => {
      const f1 = makeParsedFile(['box'], ['X']);
      const f2 = makeParsedFile(['boxes'], ['Y']);
      const result = Merger.merge([f1, f2], { smartMapping: true });

      expect(result.sheets[0].data[0]).toEqual(['box']);
      expect(result.sheets[0].data).toHaveLength(3);
    });

    test('does not strip s from short keys (3 chars or fewer)', () => {
      const f1 = makeParsedFile(['bus'], ['route1']);
      const f2 = makeParsedFile(['bu'], ['route2']);
      const result = Merger.merge([f1, f2], { smartMapping: true });

      // "bus" should NOT become "bu" — they remain separate columns
      expect(result.sheets[0].data[0]).toEqual(['bus', 'bu']);
    });

    test('does not false-match words ending in ss', () => {
      const f1 = makeParsedFile(['class'], ['A']);
      const f2 = makeParsedFile(['clas'], ['B']);
      const result = Merger.merge([f1, f2], { smartMapping: true });

      // "class" ends in ss → /[^s]s$/ does not match → kept as "class"
      expect(result.sheets[0].data[0]).toEqual(['class', 'clas']);
    });

    test('without smartMapping, underscored and spaced headers stay separate', () => {
      const f1 = makeParsedFile(['first_name'], ['Alice']);
      const f2 = makeParsedFile(['First Name'], ['Bob']);
      const result = Merger.merge([f1, f2]);

      expect(result.sheets[0].data[0]).toEqual(['first_name', 'First Name']);
    });
  });

  // ---- Custom mappings ----

  describe('custom mappings', () => {
    test('maps source header to target column', () => {
      const f1 = makeParsedFile(['email_address'], ['a@b.com']);
      const f2 = makeParsedFile(['Email'], ['c@d.com']);
      const result = Merger.merge([f1, f2], {
        customMappings: [{ from: 'email_address', to: 'Email' }],
      });

      expect(result.sheets[0].data[0]).toEqual(['Email']);
      expect(result.sheets[0].data[1]).toEqual(['a@b.com']);
      expect(result.sheets[0].data[2]).toEqual(['c@d.com']);
    });

    test('custom mapping target aligns with smart mapping', () => {
      const f1 = makeParsedFile(['email_addr'], ['a@b.com']);
      const f2 = makeParsedFile(['emails'], ['c@d.com']);
      const result = Merger.merge([f1, f2], {
        smartMapping: true,
        customMappings: [{ from: 'email_addr', to: 'emails' }],
      });

      // "email_addr" custom-mapped to keyFn("emails") = "email"
      // "emails" resolves via fuzzy to "email"
      // Both should end up in the same column
      expect(result.sheets[0].data[0]).toEqual(['emails']);
      expect(result.sheets[0].data[1]).toEqual(['a@b.com']);
      expect(result.sheets[0].data[2]).toEqual(['c@d.com']);
    });

    test('custom mapping uses display name of target', () => {
      const f1 = makeParsedFile(['fname', 'Age'], ['Alice', '30']);
      const f2 = makeParsedFile(['Full Name', 'Age'], ['Bob', '25']);
      const result = Merger.merge([f1, f2], {
        customMappings: [{ from: 'fname', to: 'Full Name' }],
      });

      expect(result.sheets[0].data[0]).toEqual(['Full Name', 'Age']);
    });

    test('ignores custom mapping with empty from or to', () => {
      const f1 = makeParsedFile(['A'], ['1']);
      const f2 = makeParsedFile(['B'], ['2']);
      const result = Merger.merge([f1, f2], {
        customMappings: [{ from: '', to: 'B' }, { from: 'A', to: '' }],
      });

      expect(result.sheets[0].data[0]).toEqual(['A', 'B']);
    });

    test('ignores custom mapping where from equals to (after normalize)', () => {
      const f1 = makeParsedFile(['Name'], ['Alice']);
      const f2 = makeParsedFile(['Name'], ['Bob']);
      const result = Merger.merge([f1, f2], {
        customMappings: [{ from: 'Name', to: '  name  ' }],
      });

      expect(result.sheets[0].data[0]).toEqual(['Name']);
      expect(result.sheets[0].data).toHaveLength(3);
    });

    test('multiple custom mappings to same target', () => {
      const f1 = makeParsedFile(['email_addr'], ['a@b.com']);
      const f2 = makeParsedFile(['e-mail'], ['c@d.com']);
      const f3 = makeParsedFile(['Email'], ['e@f.com']);
      const result = Merger.merge([f1, f2, f3], {
        customMappings: [
          { from: 'email_addr', to: 'Email' },
          { from: 'e-mail', to: 'Email' },
        ],
      });

      expect(result.sheets[0].data[0]).toEqual(['Email']);
      expect(result.sheets[0].data).toHaveLength(4);
    });

    test('does not collapse from+to columns in the same file', () => {
      // File 1 has BOTH the "from" and "to" headers — they must stay separate
      const f1 = makeParsedFile(
        ['email_address', 'Email', 'Age'],
        ['a@b', 'c@d', '30']
      );
      const f2 = makeParsedFile(['Email', 'Age'], ['e@f', '25']);
      const result = Merger.merge([f1, f2], {
        customMappings: [{ from: 'email_address', to: 'Email' }],
      });

      // email_address keeps its own column; Email keeps its own column
      expect(result.sheets[0].data[0]).toEqual(['email_address', 'Email', 'Age']);
      expect(result.sheets[0].data[1]).toEqual(['a@b', 'c@d', '30']);
      expect(result.sheets[0].data[2]).toEqual(['', 'e@f', '25']);
    });

    test('redirects from header when to header is only in other files', () => {
      const f1 = makeParsedFile(['email_address', 'Age'], ['a@b', '30']);
      const f2 = makeParsedFile(['Email', 'Age'], ['e@f', '25']);
      const result = Merger.merge([f1, f2], {
        customMappings: [{ from: 'email_address', to: 'Email' }],
      });

      // email_address merges into Email column since file 1 doesn't have Email
      expect(result.sheets[0].data[0]).toEqual(['Email', 'Age']);
      expect(result.sheets[0].data[1]).toEqual(['a@b', '30']);
      expect(result.sheets[0].data[2]).toEqual(['e@f', '25']);
    });
  });

  // ---- Duplicate headers within a file ----

  describe('duplicate headers in a single file', () => {
    test('keeps first non-empty value when columns share a name', () => {
      const f1 = makeParsedFile(['Name', 'Name'], ['Alice', 'Bob']);
      const f2 = makeParsedFile(['Name'], ['Charlie']);
      const result = Merger.merge([f1, f2]);

      // "Alice" is first non-empty value; "Bob" should not overwrite
      expect(result.sheets[0].data[0]).toEqual(['Name']);
      expect(result.sheets[0].data[1]).toEqual(['Alice']);
    });

    test('uses second value when first is empty', () => {
      const f1 = makeParsedFile(['Name', 'Name'], ['', 'Bob']);
      const f2 = makeParsedFile(['Name'], ['Charlie']);
      const result = Merger.merge([f1, f2]);

      expect(result.sheets[0].data[1]).toEqual(['Bob']);
    });
  });

  // ---- Empty header fallback ----

  describe('empty header cells', () => {
    test('assigns positional fallback name for empty headers', () => {
      const f1 = makeParsedFile(['Name', '', 'Age'], ['Alice', 'x', '30']);
      const f2 = makeParsedFile(['Name', 'Age'], ['Bob', '25']);
      const result = Merger.merge([f1, f2]);

      expect(result.sheets[0].data[0]).toEqual(['Name', 'Column 2', 'Age']);
    });

    test('whitespace-only headers also get fallback names', () => {
      const f1 = makeParsedFile(['A', '   '], ['1', '2']);
      const f2 = makeParsedFile(['A', 'B'], ['3', '4']);
      const result = Merger.merge([f1, f2]);

      expect(result.sheets[0].data[0]).toEqual(['A', 'Column 2', 'B']);
    });
  });

  // ---- detectMappings ----

  describe('detectMappings', () => {
    test('detects headers that fuzzy-match but not exact-match', () => {
      const f1 = makeParsedFile(['first_name'], ['Alice']);
      const f2 = makeParsedFile(['First Name'], ['Bob']);
      const mappings = Merger.detectMappings([f1, f2]);

      expect(mappings).toHaveLength(1);
      expect(mappings[0].canonical).toBe('first_name');
      expect(mappings[0].variants).toEqual(['first_name', 'First Name']);
    });

    test('does not flag headers that already match by normalizeKey', () => {
      const f1 = makeParsedFile(['Name'], ['Alice']);
      const f2 = makeParsedFile(['name'], ['Bob']);
      const mappings = Merger.detectMappings([f1, f2]);

      expect(mappings).toHaveLength(0);
    });

    test('returns empty for identical headers', () => {
      const f1 = makeParsedFile(['A', 'B'], ['1', '2']);
      const f2 = makeParsedFile(['A', 'B'], ['3', '4']);
      const mappings = Merger.detectMappings([f1, f2]);

      expect(mappings).toHaveLength(0);
    });

    test('groups multiple variants together', () => {
      const f1 = makeParsedFile(['first_name'], ['Alice']);
      const f2 = makeParsedFile(['First Name'], ['Bob']);
      const f3 = makeParsedFile(['first-name'], ['Charlie']);
      const mappings = Merger.detectMappings([f1, f2, f3]);

      expect(mappings).toHaveLength(1);
      expect(mappings[0].variants).toEqual(['first_name', 'First Name', 'first-name']);
    });

    test('handles empty input', () => {
      expect(Merger.detectMappings([])).toEqual([]);
    });

    test('handles files with empty sheets', () => {
      const f1 = { sheets: [{ name: 'E', data: [] }] };
      const f2 = makeParsedFile(['A'], ['1']);
      expect(Merger.detectMappings([f1, f2])).toEqual([]);
    });

    test('detects plural variants', () => {
      const f1 = makeParsedFile(['email'], ['a@b.com']);
      const f2 = makeParsedFile(['emails'], ['c@d.com']);
      const mappings = Merger.detectMappings([f1, f2]);

      expect(mappings).toHaveLength(1);
      expect(mappings[0].variants).toEqual(['email', 'emails']);
    });
  });

  // ---- collectHeadersByFile ----

  describe('collectHeadersByFile', () => {
    test('returns headers grouped by file with file names', () => {
      const f1 = makeParsedFile(['Name', 'Age'], ['Alice', '30']);
      const f2 = makeParsedFile(['Email', 'City'], ['a@b.com', 'NYC']);
      const result = Merger.collectHeadersByFile([f1, f2], ['students.csv', 'contacts.csv']);

      expect(result).toEqual([
        { fileName: 'students.csv', headers: ['Name', 'Age'] },
        { fileName: 'contacts.csv', headers: ['Email', 'City'] },
      ]);
    });

    test('uses fallback file names when not provided', () => {
      const f1 = makeParsedFile(['A'], ['1']);
      const result = Merger.collectHeadersByFile([f1]);

      expect(result).toEqual([
        { fileName: 'File 1', headers: ['A'] },
      ]);
    });

    test('returns empty headers for files with no data', () => {
      const f1 = { sheets: [{ name: 'E', data: [] }] };
      const f2 = makeParsedFile(['X'], ['1']);
      const result = Merger.collectHeadersByFile([f1, f2], ['empty.csv', 'data.csv']);

      expect(result).toEqual([
        { fileName: 'empty.csv', headers: [] },
        { fileName: 'data.csv', headers: ['X'] },
      ]);
    });

    test('skips empty/whitespace headers', () => {
      const f1 = makeParsedFile(['A', '', '  ', 'B'], ['1', '2', '3', '4']);
      const result = Merger.collectHeadersByFile([f1], ['file.csv']);

      expect(result).toEqual([
        { fileName: 'file.csv', headers: ['A', 'B'] },
      ]);
    });

    test('preserves duplicate headers within the same file', () => {
      const f1 = makeParsedFile(['Name', 'Name'], ['Alice', 'Bob']);
      const result = Merger.collectHeadersByFile([f1], ['dup.csv']);

      expect(result).toEqual([
        { fileName: 'dup.csv', headers: ['Name', 'Name'] },
      ]);
    });
  });

  // ---- Master rows untouched during mapping ----

  describe('master file integrity', () => {
    test('master rows are not mutated when secondary headers are mapped to master headers', () => {
      const master = makeParsedFile(['Name', 'Age'], ['Alice', '30'], ['Bob', '25']);
      const secondary = makeParsedFile(['name_of_students', 'Age'], ['Charlie', '20']);
      const result = Merger.merge([master, secondary], {
        customMappings: [{ from: 'name_of_students', to: 'Name' }],
      });

      // Unified header uses master column name
      expect(result.sheets[0].data[0]).toEqual(['Name', 'Age']);
      // Master rows are intact and unmodified
      expect(result.sheets[0].data[1]).toEqual(['Alice', '30']);
      expect(result.sheets[0].data[2]).toEqual(['Bob', '25']);
      // Secondary data is appended, mapped into the Name column
      expect(result.sheets[0].data[3]).toEqual(['Charlie', '20']);
    });

    test('master rows preserve values even when secondary has overlapping mapped column', () => {
      const master = makeParsedFile(['Email', 'Name'], ['a@b.com', 'Alice']);
      const secondary = makeParsedFile(['student_email'], ['c@d.com']);
      const result = Merger.merge([master, secondary], {
        customMappings: [{ from: 'student_email', to: 'Email' }],
      });

      expect(result.sheets[0].data[0]).toEqual(['Email', 'Name']);
      expect(result.sheets[0].data[1]).toEqual(['a@b.com', 'Alice']);
      expect(result.sheets[0].data[2]).toEqual(['c@d.com', '']);
    });
  });

  // ---- collectHeaders ----

  describe('collectHeaders', () => {
    test('collects unique headers from all files', () => {
      const f1 = makeParsedFile(['A', 'B'], ['1', '2']);
      const f2 = makeParsedFile(['B', 'C'], ['3', '4']);
      const headers = Merger.collectHeaders([f1, f2]);

      expect(headers).toEqual(['A', 'B', 'C']);
    });

    test('preserves first-seen display form', () => {
      const f1 = makeParsedFile(['Name'], ['Alice']);
      const f2 = makeParsedFile(['Name'], ['Bob']);
      const headers = Merger.collectHeaders([f1, f2]);

      expect(headers).toEqual(['Name']);
    });

    test('returns empty for no files', () => {
      expect(Merger.collectHeaders([])).toEqual([]);
    });

    test('skips empty/whitespace headers', () => {
      const f1 = makeParsedFile(['A', '', '  ', 'B'], ['1', '2', '3', '4']);
      const headers = Merger.collectHeaders([f1]);

      expect(headers).toEqual(['A', 'B']);
    });

    test('skips files with empty sheets', () => {
      const f1 = { sheets: [{ name: 'E', data: [] }] };
      const f2 = makeParsedFile(['X'], ['1']);
      const headers = Merger.collectHeaders([f1, f2]);

      expect(headers).toEqual(['X']);
    });
  });

  // ---- cellMeta ----

  describe('cellMeta', () => {
    test('merged formulas survive as formulas', () => {
      const f1 = {
        sheets: [{
          name: 'S1',
          data: [['A'], ['val']],
          cellMeta: [[{ type: 'string', value: 'A' }], [{ type: 'formula', value: 'SUM(1,2)' }]],
        }],
      };
      const f2 = {
        sheets: [{
          name: 'S2',
          data: [['A'], ['val2']],
          cellMeta: [[{ type: 'string', value: 'A' }], [{ type: 'formula', value: 'SUM(3,4)' }]],
        }],
      };
      const result = Merger.merge([f1, f2]);
      const meta = result.sheets[0].cellMeta;
      expect(meta[1][0]).toEqual({ type: 'formula', value: 'SUM(1,2)' });
      expect(meta[2][0]).toEqual({ type: 'formula', value: 'SUM(3,4)' });
    });

    test('merged column mappings also map metadata', () => {
      const f1 = {
        sheets: [{
          name: 'S1',
          data: [['Name', 'Age'], ['Alice', '30']],
          cellMeta: [[{ type: 'string', value: 'Name' }, { type: 'string', value: 'Age' }], [{ type: 'string', value: 'Alice' }, { type: 'number', value: 30 }]],
        }],
      };
      const f2 = {
        sheets: [{
          name: 'S2',
          data: [['Age', 'City'], ['25', 'NYC']],
          cellMeta: [[{ type: 'string', value: 'Age' }, { type: 'string', value: 'City' }], [{ type: 'number', value: 25 }, { type: 'string', value: 'NYC' }]],
        }],
      };
      const result = Merger.merge([f1, f2]);
      const meta = result.sheets[0].cellMeta;
      // Unified headers: Name, Age, City
      // f1 row: Alice, 30, empty
      expect(meta[1]).toEqual([
        { type: 'string', value: 'Alice' },
        { type: 'number', value: 30 },
        { type: 'empty' },
      ]);
      // f2 row: empty, 25, NYC
      expect(meta[2]).toEqual([
        { type: 'empty' },
        { type: 'number', value: 25 },
        { type: 'string', value: 'NYC' },
      ]);
    });

    test('generated header cells create string tokens', () => {
      const f1 = {
        sheets: [{
          name: 'S1',
          data: [['Name'], ['Alice']],
          cellMeta: [[{ type: 'string', value: 'Name' }], [{ type: 'string', value: 'Alice' }]],
        }],
      };
      const f2 = {
        sheets: [{
          name: 'S2',
          data: [['Name'], ['Bob']],
          cellMeta: [[{ type: 'string', value: 'Name' }], [{ type: 'string', value: 'Bob' }]],
        }],
      };
      const result = Merger.merge([f1, f2]);
      const meta = result.sheets[0].cellMeta;
      // Unified header row should have string tokens
      expect(meta[0][0]).toEqual({ type: 'string', value: 'Name' });
    });
  });

  // ---- cellMeta synthesis from data values ----

  describe('cellMeta synthesis from data', () => {
    test('merged CSV files get synthesized string tokens', () => {
      const f1 = makeParsedFile(['Name', 'Age'], ['Alice', '30']);
      const f2 = makeParsedFile(['Name', 'Age'], ['Bob', '25']);
      const result = Merger.merge([f1, f2]);
      const meta = result.sheets[0].cellMeta;

      expect(meta).toBeTruthy();
      expect(meta).toHaveLength(3); // header + 2 data rows
      expect(meta[0]).toEqual([
        { type: 'string', value: 'Name' },
        { type: 'string', value: 'Age' },
      ]);
      expect(meta[1]).toEqual([
        { type: 'string', value: 'Alice' },
        { type: 'string', value: '30' },
      ]);
      expect(meta[2]).toEqual([
        { type: 'string', value: 'Bob' },
        { type: 'string', value: '25' },
      ]);
    });

    test('merged mixed CSV + Excel preserves Excel types and CSV string tokens', () => {
      const csvFile = makeParsedFile(['Name', 'Score'], ['Alice', '95']);
      const excelFile = {
        sheets: [{
          name: 'E1',
          data: [['Name', 'Score'], ['Bob', 88]],
          cellMeta: [[{ type: 'string', value: 'Name' }, { type: 'string', value: 'Score' }], [{ type: 'string', value: 'Bob' }, { type: 'number', value: 88 }]],
        }],
      };
      const result = Merger.merge([csvFile, excelFile]);
      const meta = result.sheets[0].cellMeta;

      expect(meta[1][0]).toEqual({ type: 'string', value: 'Alice' });
      expect(meta[1][1]).toEqual({ type: 'string', value: '95' });
      expect(meta[2][0]).toEqual({ type: 'string', value: 'Bob' });
      expect(meta[2][1]).toEqual({ type: 'number', value: 88 });
    });

    test('merged two metadata-less sheets still produces complete cellMeta', () => {
      const f1 = makeParsedFile(['A', 'B'], ['x', '']);
      const f2 = makeParsedFile(['A', 'B'], ['', 'y']);
      const result = Merger.merge([f1, f2]);
      const meta = result.sheets[0].cellMeta;

      expect(meta).toHaveLength(3);
      expect(meta[1][1]).toEqual({ type: 'empty' });
      expect(meta[2][0]).toEqual({ type: 'empty' });
      expect(meta[2][1]).toEqual({ type: 'string', value: 'y' });
    });

    test('cellMeta synthesized from data does not produce empty tokens for non-empty values', () => {
      const f1 = {
        sheets: [{
          name: 'S1',
          data: [['Value'], [42], [true], [0], ['']],
        }],
      };
      const result = Merger.merge([f1]);
      const meta = result.sheets[0].cellMeta;

      expect(meta[1][0]).toEqual({ type: 'number', value: 42 });
      expect(meta[2][0]).toEqual({ type: 'boolean', value: true });
      expect(meta[3][0]).toEqual({ type: 'number', value: 0 });
      expect(meta[4][0]).toEqual({ type: 'empty' });
    });

    test('duplicate headers: formula returning empty string wins mapped position, metadata follows', () => {
      const formula = '=IF(FALSE,"x","")';
      const f1 = {
        sheets: [{
          name: 'S1',
          data: [['Col', 'Col', 'Col'], ['', 'real', 'later literal']],
          cellMeta: [
            [
              { type: 'string', value: 'Col' },
              { type: 'string', value: 'Col' },
              { type: 'string', value: 'Col' },
            ],
            [
              { type: 'formula', value: formula, displayValue: '' },
              { type: 'string', value: 'real' },
              { type: 'string', value: 'later literal' },
            ],
          ],
        }],
      };
      const f2 = {
        sheets: [{
          name: 'S2',
          data: [['Col'], ['other']],
          cellMeta: [[{ type: 'string', value: 'Col' }], [{ type: 'string', value: 'other' }]],
        }],
      };
      const result = Merger.merge([f1, f2]);
      const meta = result.sheets[0].cellMeta;
      // The formula source has an empty cached display value, but its token
      // occupies the destination and prevents later duplicate literals from
      // replacing the paired data/token selection.
      expect(result.sheets[0].data[1][0]).toBe('');
      expect(meta[1][0]).toEqual({ type: 'formula', value: formula, displayValue: '' });
      // f2 row should come through
      expect(meta[2][0]).toEqual({ type: 'string', value: 'other' });
    });
  });
});
