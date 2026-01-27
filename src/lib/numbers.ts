import * as XLSX from "xlsx";
import * as fs from "fs";
import * as path from "path";

// Configure XLSX for Node.js filesystem access
XLSX.set_fs(fs);

// Types
export interface TableRow {
  [key: string]: unknown;
}

export interface TableData {
  headers: string[];
  rows: TableRow[];
  rowCount: number;
  totalRows: number;
  hasMore: boolean;
}

export interface SheetInfo {
  name: string;
  headers: string[];
  rowCount: number;
  columnCount: number;
}

export interface WorkbookInfo {
  sheets: string[];
  activeSheet: SheetInfo;
  sampleRows: TableRow[];
}

export interface WhereCondition {
  [column: string]: unknown | { op: "contains" | "gt" | "lt" | "gte" | "lte" | "ne"; value: unknown };
}

// Read workbook from file
export function readWorkbook(filePath: string): XLSX.WorkBook {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
  return XLSX.readFile(filePath);
}

// Get sheet from workbook (defaults to first sheet)
export function getSheet(wb: XLSX.WorkBook, sheetName?: string): XLSX.WorkSheet {
  const name = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[name];
  if (!ws) {
    throw new Error(`Sheet '${name}' not found. Available: ${wb.SheetNames.join(", ")}`);
  }
  return ws;
}

// Convert sheet to array of objects with headers
export function sheetToTable(ws: XLSX.WorkSheet): { headers: string[]; rows: TableRow[] } {
  // Get all rows as array of arrays
  const aoa: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  if (aoa.length === 0) {
    return { headers: [], rows: [] };
  }

  // First row is headers
  const headers = aoa[0].map((h, i) => (h != null ? String(h) : `Column${i + 1}`));

  // Convert remaining rows to objects
  const rows: TableRow[] = [];
  for (let i = 1; i < aoa.length; i++) {
    const row: TableRow = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = aoa[i]?.[j] ?? null;
    }
    rows.push(row);
  }

  return { headers, rows };
}

// Filter rows by column selection
export function filterColumns(rows: TableRow[], columns: string[]): TableRow[] {
  return rows.map((row) => {
    const filtered: TableRow = {};
    for (const col of columns) {
      filtered[col] = row[col];
    }
    return filtered;
  });
}

// Apply pagination
export function paginate<T>(items: T[], limit?: number, offset?: number): { items: T[]; hasMore: boolean } {
  const start = offset || 0;
  const end = limit ? start + limit : items.length;
  const sliced = items.slice(start, end);
  return {
    items: sliced,
    hasMore: end < items.length,
  };
}

// Check if a row matches a where condition
export function matchesCondition(row: TableRow, where: WhereCondition): boolean {
  for (const [column, condition] of Object.entries(where)) {
    const value = row[column];

    // Handle operator-based conditions
    if (condition && typeof condition === "object" && "op" in condition) {
      const { op, value: condValue } = condition as { op: string; value: unknown };

      switch (op) {
        case "contains":
          if (!String(value).toLowerCase().includes(String(condValue).toLowerCase())) {
            return false;
          }
          break;
        case "gt":
          if (!(Number(value) > Number(condValue))) return false;
          break;
        case "gte":
          if (!(Number(value) >= Number(condValue))) return false;
          break;
        case "lt":
          if (!(Number(value) < Number(condValue))) return false;
          break;
        case "lte":
          if (!(Number(value) <= Number(condValue))) return false;
          break;
        case "ne":
          if (value === condValue) return false;
          break;
        default:
          throw new Error(`Unknown operator: ${op}`);
      }
    } else {
      // Simple equality check
      if (value !== condition) {
        return false;
      }
    }
  }
  return true;
}

// Find rows matching a condition
export function findRows(rows: TableRow[], where: WhereCondition): TableRow[] {
  return rows.filter((row) => matchesCondition(row, where));
}

// Get workbook info (quick overview)
export function getWorkbookInfo(filePath: string): WorkbookInfo {
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb);
  const { headers, rows } = sheetToTable(ws);

  return {
    sheets: wb.SheetNames,
    activeSheet: {
      name: wb.SheetNames[0],
      headers,
      rowCount: rows.length,
      columnCount: headers.length,
    },
    sampleRows: rows.slice(0, 3),
  };
}

// Get table data with options
export function getTableData(
  filePath: string,
  options: {
    sheetName?: string;
    columns?: string[];
    limit?: number;
    offset?: number;
  } = {}
): TableData {
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, options.sheetName);
  let { headers, rows } = sheetToTable(ws);

  const totalRows = rows.length;

  // Apply column filter
  if (options.columns && options.columns.length > 0) {
    rows = filterColumns(rows, options.columns);
    headers = options.columns.filter((c) => headers.includes(c));
  }

  // Apply pagination
  const { items, hasMore } = paginate(rows, options.limit, options.offset);

  return {
    headers,
    rows: items,
    rowCount: items.length,
    totalRows,
    hasMore,
  };
}

// Find rows with options
export function findTableRows(
  filePath: string,
  where: WhereCondition,
  options: {
    sheetName?: string;
    columns?: string[];
    limit?: number;
    offset?: number;
  } = {}
): { rows: TableRow[]; count: number; totalMatches: number; hasMore: boolean } {
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, options.sheetName);
  let { rows } = sheetToTable(ws);

  // Apply where condition
  const matched = findRows(rows, where);
  const totalMatches = matched.length;

  // Apply column filter
  let filteredRows = matched;
  if (options.columns && options.columns.length > 0) {
    filteredRows = filterColumns(matched, options.columns);
  }

  // Apply pagination
  const { items, hasMore } = paginate(filteredRows, options.limit, options.offset);

  return {
    rows: items,
    count: items.length,
    totalMatches,
    hasMore,
  };
}

// Helper to write workbook using buffer (bypasses format detection issues)
function writeWorkbook(wb: XLSX.WorkBook, filePath: string): void {
  const absPath = path.resolve(filePath);
  const data = XLSX.write(wb, { bookType: "numbers", type: "buffer" });
  fs.writeFileSync(absPath, data);
}

// Update rows matching condition
export function updateTableRows(
  filePath: string,
  where: WhereCondition,
  set: { [column: string]: unknown },
  options: { sheetName?: string } = {}
): { updatedCount: number } {
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, options.sheetName);
  const { headers, rows } = sheetToTable(ws);

  let updatedCount = 0;

  // Update matching rows
  for (const row of rows) {
    if (matchesCondition(row, where)) {
      for (const [col, val] of Object.entries(set)) {
        row[col] = val;
      }
      updatedCount++;
    }
  }

  // Convert back to sheet
  const newAoa: unknown[][] = [headers];
  for (const row of rows) {
    newAoa.push(headers.map((h) => row[h]));
  }

  const newWs = XLSX.utils.aoa_to_sheet(newAoa);
  wb.Sheets[options.sheetName || wb.SheetNames[0]] = newWs;

  writeWorkbook(wb, filePath);

  return { updatedCount };
}

// Add rows to table
export function addTableRows(
  filePath: string,
  newRows: TableRow | TableRow[],
  options: { sheetName?: string } = {}
): { addedCount: number; totalRows: number } {
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, options.sheetName);
  const { headers, rows } = sheetToTable(ws);

  const rowsToAdd = Array.isArray(newRows) ? newRows : [newRows];

  // Add any new columns from the new rows
  const allHeaders = new Set(headers);
  for (const row of rowsToAdd) {
    for (const key of Object.keys(row)) {
      allHeaders.add(key);
    }
  }
  const finalHeaders = Array.from(allHeaders);

  // Add the new rows
  rows.push(...rowsToAdd);

  // Convert back to sheet
  const newAoa: unknown[][] = [finalHeaders];
  for (const row of rows) {
    newAoa.push(finalHeaders.map((h) => row[h] ?? null));
  }

  const newWs = XLSX.utils.aoa_to_sheet(newAoa);
  wb.Sheets[options.sheetName || wb.SheetNames[0]] = newWs;

  writeWorkbook(wb, filePath);

  return { addedCount: rowsToAdd.length, totalRows: rows.length };
}

// Delete rows matching condition
export function deleteTableRows(
  filePath: string,
  where: WhereCondition,
  options: { sheetName?: string } = {}
): { deletedCount: number; remainingRows: number } {
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, options.sheetName);
  const { headers, rows } = sheetToTable(ws);

  const originalCount = rows.length;
  const remaining = rows.filter((row) => !matchesCondition(row, where));
  const deletedCount = originalCount - remaining.length;

  // Convert back to sheet
  const newAoa: unknown[][] = [headers];
  for (const row of remaining) {
    newAoa.push(headers.map((h) => row[h]));
  }

  const newWs = XLSX.utils.aoa_to_sheet(newAoa);
  wb.Sheets[options.sheetName || wb.SheetNames[0]] = newWs;

  writeWorkbook(wb, filePath);

  return { deletedCount, remainingRows: remaining.length };
}

// Batch update rows with different values per condition
export function batchUpdateTableRows(
  filePath: string,
  updates: Array<{ where: WhereCondition; set: { [column: string]: unknown } }>,
  options: { sheetName?: string } = {}
): { updatedCount: number; details: Array<{ where: WhereCondition; matched: number }> } {
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, options.sheetName);
  const { headers, rows } = sheetToTable(ws);

  let totalUpdated = 0;
  const details: Array<{ where: WhereCondition; matched: number }> = [];

  for (const { where, set } of updates) {
    let matched = 0;
    for (const row of rows) {
      if (matchesCondition(row, where)) {
        for (const [col, val] of Object.entries(set)) {
          row[col] = val;
        }
        matched++;
        totalUpdated++;
      }
    }
    details.push({ where, matched });
  }

  // Convert back to sheet
  const newAoa: unknown[][] = [headers];
  for (const row of rows) {
    newAoa.push(headers.map((h) => row[h]));
  }

  const newWs = XLSX.utils.aoa_to_sheet(newAoa);
  wb.Sheets[options.sheetName || wb.SheetNames[0]] = newWs;

  writeWorkbook(wb, filePath);

  return { updatedCount: totalUpdated, details };
}
