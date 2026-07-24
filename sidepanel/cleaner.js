/**
 * Data cleaning utilities for spreadsheet data.
 * All functions operate on 2D string arrays (first row = headers).
 */

// eslint-disable-next-line no-unused-vars
const Cleaner = (() => {
  'use strict';

  /**
   * Apply all selected cleaning operations in sequence.
   * @param {string[][]} data - 2D array (row 0 = headers)
   * @param {Object} options - Cleaning options
   * @returns {string[][]}
   */
  function apply(data, options, cellMeta) {
    if (!data || data.length === 0) return cellMeta ? { data, cellMeta } : data;

    const hasWork =
      options.trim ||
      options.removeEmptyRows ||
      options.removeEmptyColumns ||
      options.removeDuplicates ||
      options.fixNumbers ||
      options.normalizeHeaders;

    if (!hasWork) return cellMeta ? { data, cellMeta } : data;

    let result = data;
    let meta = cellMeta ? clone2D(cellMeta) : null;

    if (options.trim) {
      result = trimWhitespace(result);
    }
    if (options.removeEmptyRows) {
      result = removeEmptyRows(result);
      if (meta) meta = filterRowsByData(result, data, meta);
      data = result; // update reference for subsequent index tracking
    }
    if (options.removeEmptyColumns) {
      result = removeEmptyColumns(result);
      if (meta) meta = filterColsByData(result, meta);
    }
    if (options.removeDuplicates) {
      result = options.duplicateMode === 'absolute'
        ? removeAbsoluteDuplicates(result)
        : removeDuplicateRows(result);
      if (meta) meta = filterRowsByData(result, data, meta);
    }
    if (options.fixNumbers) {
      result = fixNumberFormatting(result, meta);
    }
    if (options.normalizeHeaders) {
      result = normalizeHeaders(result);
    }

    return cellMeta ? { data: result, cellMeta: meta } : result;
  }

  function clone2D(arr) {
    return arr.map((row) => row.slice());
  }

  function filterRowsByData(newData, oldData, meta) {
    if (!meta) return null;
    const keep = [];
    for (let i = 0, ni = 0; i < oldData.length; i++) {
      if (ni < newData.length && rowsEqual(newData[ni], oldData[i])) {
        keep.push(i);
        ni++;
      }
    }
    return keep.map((idx) => meta[idx]);
  }

  function rowsEqual(a, b) {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }

  function filterColsByData(newData, meta) {
    if (!meta) return null;
    const colCount = newData[0] ? newData[0].length : 0;
    return meta.map((row) => {
      const newRow = new Array(colCount);
      for (let c = 0; c < colCount; c++) newRow[c] = row ? row[c] || { type: 'empty' } : { type: 'empty' };
      return newRow;
    });
  }

  /**
   * Trim leading/trailing whitespace from string cells only.
   * Numbers, booleans, null, and undefined pass through unchanged.
   */
  function trimWhitespace(data) {
    return data.map((row) =>
      row.map((cell) => {
        if (typeof cell === 'string') return cell.trim();
        return cell;
      })
    );
  }

  /**
   * Remove rows where all cells are empty or whitespace.
   * Always keeps the header row (index 0).
   */
  function removeEmptyRows(data) {
    if (data.length === 0) return data;
    return [
      data[0],
      ...data.slice(1).filter((row) =>
        row.some((cell) => {
          if (cell === null || cell === undefined) return false;
          if (typeof cell === 'string') return cell.trim().length > 0;
          return true;
        })
      ),
    ];
  }

  /**
   * Remove columns where all cells (including header) are empty.
   */
  function removeEmptyColumns(data) {
    if (data.length === 0) return data;
    const colCount = data.reduce((max, r) => Math.max(max, r.length), 0);
    const keepCols = [];
    for (let col = 0; col < colCount; col++) {
      if (data.some((row) => {
        const val = row[col];
        if (val === null || val === undefined) return false;
        if (typeof val === 'string') return val.trim().length > 0;
        return true;
      })) keepCols.push(col);
    }
    return data.map((row) => keepCols.map((col) => row[col] ?? ''));
  }

  /**
   * Remove ALL occurrences of any row that appears more than once.
   * Uses token-based comparison via tokenFromValue + rowComparisonKey.
   */
  function removeAbsoluteDuplicates(data) {
    if (data.length <= 1) return data;
    const header = data[0];
    const keys = new Array(data.length);
    const counts = new Map();
    for (let i = 1; i < data.length; i++) {
      const tokens = data[i].map((v) => tokenFromValue(v));
      const key = rowComparisonKey(tokens, false, []);
      keys[i] = key;
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    const result = [header];
    for (let i = 1; i < data.length; i++) {
      if (counts.get(keys[i]) === 1) result.push(data[i]);
    }
    return result;
  }

  function removeDuplicateRows(data) {
    if (data.length <= 1) return data;
    const header = data[0];
    const seen = new Set();
    const result = [header];
    for (let i = 1; i < data.length; i++) {
      const tokens = data[i].map((v) => tokenFromValue(v));
      const key = rowComparisonKey(tokens, false, []);
      if (!seen.has(key)) {
        seen.add(key);
        result.push(data[i]);
      }
    }
    return result;
  }

  /**
   * Convert text-formatted numbers to actual numbers.
   * Skips the header row. Converts all eligible numeric-looking strings.
   * Preserves leading-zero identifiers (postal codes, SKUs, etc.).
   */
  function fixNumberFormatting(data, meta) {
    if (data.length <= 1) return data;

    return data.map((row, rowIndex) => {
      if (rowIndex === 0) return row;
      return row.map((cell, colIndex) => {
        if (typeof cell !== 'string') return cell;
        // Skip TEXT-formatted cells
        if (meta && meta[rowIndex] && meta[rowIndex][colIndex] && meta[rowIndex][colIndex].formatType === 'TEXT') return cell;
        const trimmed = cell.trim();
        if (trimmed === '') return cell;

        const cleaned = trimmed.replace(/[,\s]/g, '');
        if (/^-?\d+(\.\d+)?$/.test(cleaned)) {
          if (cleaned.length > 1 && cleaned.startsWith('0') && !cleaned.startsWith('0.')) {
            return cleaned;
          }
          const num = Number(cleaned);
          if (Number.isFinite(num)) return num;
        }

        return cell;
      });
    });
  }

  /**
   * Normalize header names:
   * - Trim whitespace
   * - Collapse multiple spaces to one
   * - Convert to Title Case
   */
  function normalizeHeaders(data) {
    if (data.length === 0) return data;

    const result = [...data];
    result[0] = data[0].map((header) => {
      if (typeof header !== 'string') return header;
      return header
        .trim()
        .replace(/\s+/g, ' ')
        .toLowerCase()
        .replace(/\b\w/g, (ch) => ch.toUpperCase());
    });

    return result;
  }

  /**
   * Get cleaning statistics (before/after comparison).
   */
  function getStats(original, cleaned) {
    const origRows = original.length;
    const cleanedRows = cleaned.length;
    const origCols = original[0]?.length || 0;
    const cleanedCols = cleaned[0]?.length || 0;

    return {
      rowsRemoved: origRows - cleanedRows,
      colsRemoved: origCols - cleanedCols,
      originalRows: origRows,
      cleanedRows,
      originalCols: origCols,
      cleanedCols,
    };
  }

  // ================================================================
  //  Cell-token helpers — shared by local preview and Google-side
  //  structural planning to guarantee identical row-identity and
  //  emptiness semantics.
  // ================================================================

  /**
   * Build a cell token from a raw 2D-array value (used by the local
   * Cleaner path for structural comparison).
   */
  function tokenFromValue(value) {
    if (value === null || value === undefined || value === '') return { type: 'empty' };
    if (typeof value === 'number') return { type: 'number', value };
    if (typeof value === 'boolean') return { type: 'boolean', value };
    if (typeof value === 'string') return { type: 'string', value };
    return { type: 'string', value: String(value) };
  }

  /**
   * Build a cell token from a Sheets API CellData object (used by
   * the Google-side structural-planning path).
   */
  function tokenFromCellData(cellData) {
    const uev = (cellData && cellData.userEnteredValue) || {};
    const fmtType = cellData && cellData.effectiveFormat && cellData.effectiveFormat.numberFormat
      ? cellData.effectiveFormat.numberFormat.type
      : undefined;

    if (uev.formulaValue !== undefined) return { type: 'formula', value: uev.formulaValue };
    if (fmtType === 'DATE' || fmtType === 'TIME' || fmtType === 'DATE_TIME') return { type: 'date', value: uev.numberValue };
    if (uev.boolValue !== undefined) return { type: 'boolean', value: uev.boolValue };
    if (uev.numberValue !== undefined) return { type: 'number', value: uev.numberValue };
    if (uev.stringValue !== undefined && uev.stringValue !== '') return { type: 'string', value: uev.stringValue };
    return { type: 'empty' };
  }

  /**
   * Returns true when a token represents an effectively empty cell.
   * Whitespace-only strings are treated as empty.
   */
  function isTokenEmpty(token) {
    if (!token || token.type === 'empty') return true;
    if (token.type === 'string' && String(token.value || '').trim().length === 0) return true;
    return false;
  }

  /**
   * Stable comparison key for a single cell token.
   * When `shouldTrim` is true, string values are trimmed before keying.
   */
  function tokenComparisonKey(token, shouldTrim) {
    if (token.type === 'string' && shouldTrim) {
      return `string\x00${String(token.value ?? '').trim()}`;
    }
    return `${token.type}\x00${token.value ?? ''}`;
  }

  /**
   * Build a row-level comparison key from an array of cell tokens.
   * Columns in `excludedCols` are skipped (for empty-column handling).
   */
  function rowComparisonKey(tokens, shouldTrim, excludedCols) {
    const excluded = new Set(excludedCols || []);
    return tokens.reduce((parts, token, idx) => {
      if (excluded.has(idx)) return parts;
      parts.push(tokenComparisonKey(token, shouldTrim));
      return parts;
    }, []).join('\x01');
  }

  return {
    apply,
    trimWhitespace,
    removeEmptyRows,
    removeEmptyColumns,
    removeDuplicateRows,
    removeAbsoluteDuplicates,
    fixNumberFormatting,
    normalizeHeaders,
    getStats,
    // Cell-token helpers
    tokenFromValue,
    tokenFromCellData,
    isTokenEmpty,
    tokenComparisonKey,
    rowComparisonKey,
  };
})();
