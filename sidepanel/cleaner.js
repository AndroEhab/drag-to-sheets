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
   * @param {Object[][]} [cellMeta] - Optional 2D cell token array
   * @returns {string[][]|{data: string[][], cellMeta: Object[][]}}
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
    let meta = cellMeta || null;

    if (options.trim) {
      const trimResult = trimWhitespace(result, meta);
      result = trimResult.data;
      meta = trimResult.cellMeta || null;
    }
    if (options.removeEmptyRows) {
      const opResult = removeEmptyRows(result, meta);
      result = opResult.data;
      meta = opResult.cellMeta || null;
    }
    if (options.removeEmptyColumns) {
      const opResult = removeEmptyColumns(result, meta);
      result = opResult.data;
      meta = opResult.cellMeta || null;
    }
    if (options.removeDuplicates) {
      const opResult = options.duplicateMode === 'absolute'
        ? removeAbsoluteDuplicates(result, meta)
        : removeDuplicateRows(result, meta);
      result = opResult.data;
      meta = opResult.cellMeta || null;
    }
    if (options.fixNumbers) {
      const opResult = fixNumberFormatting(result, meta);
      result = opResult.data;
      meta = opResult.cellMeta || null;
    }
    if (options.normalizeHeaders) {
      result = normalizeHeaders(result);
    }

    return cellMeta ? { data: result, cellMeta: meta } : result;
  }

  /**
   * Trim leading/trailing whitespace from string cells only.
   * Numbers, booleans, null, and undefined pass through unchanged.
   */
  function trimWhitespace(data, meta) {
    const trimmedData = data.map((row, ri) =>
      row.map((cell, ci) => {
        if (typeof cell === 'string') {
          const trimmed = cell.trim();
          if (meta && meta[ri] && meta[ri][ci] && meta[ri][ci].type === 'string') {
            meta[ri][ci].value = trimmed;
          }
          return trimmed;
        }
        return cell;
      })
    );
    return { data: trimmedData, cellMeta: meta || null };
  }

  /**
   * Remove rows where all cells are empty or whitespace.
   * Always keeps the header row (index 0).
   * @returns {{data: string[][]}} or {{data: string[][], cellMeta: Object[][]}}
   */
  function removeEmptyRows(data, meta) {
    if (data.length === 0) return meta ? { data, cellMeta: meta } : { data };

    const keepIndices = [0];
    for (let i = 1; i < data.length; i++) {
      const isEmpty = meta
        ? meta[i].every((token) => isTokenEmpty(token))
        : data[i].every((cell) => {
            if (cell === null || cell === undefined) return true;
            if (typeof cell === 'string') return cell.trim().length === 0;
            return false;
          });
      if (!isEmpty) keepIndices.push(i);
    }

    const result = { data: keepIndices.map((i) => data[i]) };
    if (meta) result.cellMeta = keepIndices.map((i) => meta[i]);
    return result;
  }

  /**
   * Remove columns where all cells (including header) are empty.
   * @returns {{data: string[][]}} or {{data: string[][], cellMeta: Object[][]}}
   */
  function removeEmptyColumns(data, meta) {
    if (data.length === 0) return meta ? { data, cellMeta: meta } : { data };

    const colCount = data.reduce((max, r) => Math.max(max, r.length), 0);
    const keepCols = [];
    for (let col = 0; col < colCount; col++) {
      const hasContent = meta
        ? meta.some((row) => row[col] && !isTokenEmpty(row[col]))
        : data.some((row) => {
            const val = row[col];
            if (val === null || val === undefined) return false;
            if (typeof val === 'string') return val.trim().length > 0;
            return true;
          });
      if (hasContent) keepCols.push(col);
    }

    const result = { data: data.map((row) => keepCols.map((col) => row[col] ?? '')) };
    if (meta) {
      result.cellMeta = meta.map((row) =>
        keepCols.map((col) => (row ? row[col] || { type: 'empty' } : { type: 'empty' }))
      );
    }
    return result;
  }

  /**
   * Remove ALL occurrences of any row that appears more than once.
   * Uses token-based comparison via metadata tokens (or tokenFromValue fallback).
   * @returns {{data: string[][]}} or {{data: string[][], cellMeta: Object[][]}}
   */
  function removeAbsoluteDuplicates(data, meta) {
    if (data.length <= 1) return meta ? { data, cellMeta: meta } : { data };

    const keys = new Array(data.length);
    const counts = new Map();
    for (let i = 1; i < data.length; i++) {
      let tokens;
      if (meta && meta[i]) {
        tokens = meta[i];
      } else {
        tokens = data[i].map((v) => tokenFromValue(v));
      }
      const key = rowComparisonKey(tokens, false, []);
      keys[i] = key;
      counts.set(key, (counts.get(key) || 0) + 1);
    }

    const keepIndices = [0];
    for (let i = 1; i < data.length; i++) {
      if (counts.get(keys[i]) === 1) keepIndices.push(i);
    }

    const result = { data: keepIndices.map((i) => data[i]) };
    if (meta) result.cellMeta = keepIndices.map((i) => meta[i]);
    return result;
  }

  /**
   * Remove duplicate rows keeping first occurrence.
   * Uses token-based comparison via metadata tokens (or tokenFromValue fallback).
   * @returns {{data: string[][]}} or {{data: string[][], cellMeta: Object[][]}}
   */
  function removeDuplicateRows(data, meta) {
    if (data.length <= 1) return meta ? { data, cellMeta: meta } : { data };

    const seen = new Set();
    const keepIndices = [0];
    for (let i = 1; i < data.length; i++) {
      let tokens;
      if (meta && meta[i]) {
        tokens = meta[i];
      } else {
        tokens = data[i].map((v) => tokenFromValue(v));
      }
      const key = rowComparisonKey(tokens, false, []);
      if (!seen.has(key)) {
        seen.add(key);
        keepIndices.push(i);
      }
    }

    const result = { data: keepIndices.map((i) => data[i]) };
    if (meta) result.cellMeta = keepIndices.map((i) => meta[i]);
    return result;
  }

  /**
   * Convert text-formatted numbers to actual numbers.
   * Skips the header row. Converts all eligible numeric-looking strings.
   * Preserves leading-zero identifiers (postal codes, SKUs, etc.).
   * @returns {{data: string[][]}} or {{data: string[][], cellMeta: Object[][]}}
   */
  function fixNumberFormatting(data, meta) {
    if (data.length <= 1) return meta ? { data, cellMeta: meta } : { data };

    let newMeta = meta ? meta.map((row) => [...row]) : null;

    const newData = data.map((row, rowIndex) => {
      if (rowIndex === 0) return row;
      return row.map((cell, colIndex) => {
        if (typeof cell !== 'string') return cell;
        if (meta && meta[rowIndex] && meta[rowIndex][colIndex] && meta[rowIndex][colIndex].formatType === 'TEXT') return cell;
        const trimmed = cell.trim();
        if (trimmed === '') return cell;

        const cleaned = trimmed.replace(/[,\s]/g, '');
        if (/^-?\d+(\.\d+)?$/.test(cleaned)) {
          if (cleaned.length > 1 && cleaned.startsWith('0') && !cleaned.startsWith('0.')) {
            return cleaned;
          }
          const num = Number(cleaned);
          if (Number.isFinite(num)) {
            if (newMeta) newMeta[rowIndex][colIndex] = { type: 'number', value: num };
            return num;
          }
        }

        return cell;
      });
    });

    const result = { data: newData };
    if (meta) result.cellMeta = newMeta;
    return result;
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
    if (token.type === 'formula') {
      return `formula\x00${token.value ?? ''}`;
    }
    if (token.type === 'date') {
      return `date\x00${token.value ?? ''}\x00${token.formatType || 'DATE'}`;
    }
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
