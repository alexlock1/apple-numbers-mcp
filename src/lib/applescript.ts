import { exec } from "child_process";
import { promisify } from "util";
import * as path from "path";
import { sheetToTable, readWorkbook, getSheet, matchesCondition, WhereCondition } from "./numbers.js";

const execAsync = promisify(exec);

export interface RGBColor {
  r: number; // 0-255
  g: number;
  b: number;
}

// Convert 0-255 RGB to Numbers' 0-65535 scale
function toNumbersColor(color: RGBColor): { r: number; g: number; b: number } {
  return {
    r: Math.round((color.r / 255) * 65535),
    g: Math.round((color.g / 255) * 65535),
    b: Math.round((color.b / 255) * 65535),
  };
}

// Execute AppleScript and return result
async function runAppleScript(script: string): Promise<string> {
  try {
    const { stdout } = await execAsync(`osascript -e '${script.replace(/'/g, "'\"'\"'")}'`);
    return stdout.trim();
  } catch (error) {
    const err = error as Error & { stderr?: string };
    throw new Error(`AppleScript error: ${err.stderr || err.message}`);
  }
}

// Set a single row's background color
export async function setRowColor(
  filePath: string,
  sheetName: string | undefined,
  rowIndex: number, // 1-based (includes header row)
  color: RGBColor
): Promise<void> {
  const absPath = path.resolve(filePath);
  const numbersColor = toNumbersColor(color);

  // AppleScript to set row background color
  const script = `
tell application "Numbers"
  activate
  set theDoc to open POSIX file "${absPath}"
  delay 0.5
  tell theDoc
    tell sheet 1
      tell table 1
        set background color of row ${rowIndex} to {${numbersColor.r}, ${numbersColor.g}, ${numbersColor.b}}
      end tell
    end tell
  end tell
  save theDoc
end tell
  `.trim();

  await runAppleScript(script);
}

// Color multiple rows based on column value
export async function colorRowsByValue(
  filePath: string,
  sheetName: string | undefined,
  column: string,
  colorMap: { [value: string]: RGBColor }
): Promise<{ coloredCount: number }> {
  const absPath = path.resolve(filePath);

  // First, read the file to find which rows need coloring
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, sheetName);
  const { headers, rows } = sheetToTable(ws);

  // Find column index
  const colIndex = headers.indexOf(column);
  if (colIndex === -1) {
    throw new Error(`Column '${column}' not found. Available: ${headers.join(", ")}`);
  }

  // Build list of row colors
  const rowColors: { rowIndex: number; color: { r: number; g: number; b: number } }[] = [];

  for (let i = 0; i < rows.length; i++) {
    const cellValue = String(rows[i][column] ?? "");
    if (colorMap[cellValue]) {
      rowColors.push({
        rowIndex: i + 2, // +1 for header row, +1 for 1-based indexing
        color: toNumbersColor(colorMap[cellValue]),
      });
    }
  }

  if (rowColors.length === 0) {
    return { coloredCount: 0 };
  }

  // Build AppleScript to color all rows at once
  const colorCommands = rowColors
    .map(({ rowIndex, color }) =>
      `set background color of row ${rowIndex} to {${color.r}, ${color.g}, ${color.b}}`
    )
    .join("\n        ");

  const script = `
tell application "Numbers"
  activate
  set theDoc to open POSIX file "${absPath}"
  delay 0.5
  tell theDoc
    tell sheet 1
      tell table 1
        ${colorCommands}
      end tell
    end tell
  end tell
  save theDoc
end tell
  `.trim();

  await runAppleScript(script);

  return { coloredCount: rowColors.length };
}

// Set a cell value via AppleScript
export async function setCellValue(
  filePath: string,
  rowIndex: number, // 1-based (row 1 is header)
  colIndex: number, // 1-based
  value: unknown
): Promise<void> {
  const absPath = path.resolve(filePath);

  // Format value for AppleScript
  let formattedValue: string;
  if (value === null || value === undefined) {
    formattedValue = '""';
  } else if (typeof value === "number") {
    formattedValue = String(value);
  } else if (typeof value === "boolean") {
    formattedValue = value ? "true" : "false";
  } else {
    // Escape quotes in string values
    formattedValue = `"${String(value).replace(/"/g, '\\"')}"`;
  }

  const script = `
tell application "Numbers"
  activate
  set theDoc to open POSIX file "${absPath}"
  delay 0.3
  tell theDoc
    tell sheet 1
      tell table 1
        set value of cell ${colIndex} of row ${rowIndex} to ${formattedValue}
      end tell
    end tell
  end tell
  save theDoc
end tell
  `.trim();

  await runAppleScript(script);
}

// Update multiple cells in batch (more efficient than individual calls)
export async function updateCellsBatch(
  filePath: string,
  updates: Array<{ rowIndex: number; colIndex: number; value: unknown }>
): Promise<void> {
  if (updates.length === 0) return;

  const absPath = path.resolve(filePath);

  // Build cell update commands
  const cellCommands = updates.map(({ rowIndex, colIndex, value }) => {
    let formattedValue: string;
    if (value === null || value === undefined) {
      formattedValue = '""';
    } else if (typeof value === "number") {
      formattedValue = String(value);
    } else if (typeof value === "boolean") {
      formattedValue = value ? "true" : "false";
    } else {
      formattedValue = `"${String(value).replace(/"/g, '\\"')}"`;
    }
    return `set value of cell ${colIndex} of row ${rowIndex} to ${formattedValue}`;
  }).join("\n        ");

  const script = `
tell application "Numbers"
  activate
  set theDoc to open POSIX file "${absPath}"
  delay 0.3
  tell theDoc
    tell sheet 1
      tell table 1
        ${cellCommands}
      end tell
    end tell
  end tell
  save theDoc
end tell
  `.trim();

  await runAppleScript(script);
}

// Update rows matching a condition (uses SheetJS to read, AppleScript to write)
export async function updateRowsViaAppleScript(
  filePath: string,
  sheetName: string | undefined,
  where: WhereCondition,
  set: { [column: string]: unknown }
): Promise<{ updatedCount: number }> {
  // Read the file to find matching rows
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, sheetName);
  const { headers, rows } = sheetToTable(ws);

  // Find column indices for the columns we're updating
  const colIndices: { [col: string]: number } = {};
  for (const col of Object.keys(set)) {
    const idx = headers.indexOf(col);
    if (idx === -1) {
      throw new Error(`Column '${col}' not found. Available: ${headers.join(", ")}`);
    }
    colIndices[col] = idx + 1; // 1-based for AppleScript
  }

  // Find matching rows and build updates
  const updates: Array<{ rowIndex: number; colIndex: number; value: unknown }> = [];

  for (let i = 0; i < rows.length; i++) {
    if (matchesCondition(rows[i], where)) {
      const rowIndex = i + 2; // +1 for header, +1 for 1-based
      for (const [col, value] of Object.entries(set)) {
        updates.push({
          rowIndex,
          colIndex: colIndices[col],
          value,
        });
      }
    }
  }

  if (updates.length === 0) {
    return { updatedCount: 0 };
  }

  // Update via AppleScript
  await updateCellsBatch(filePath, updates);

  return { updatedCount: updates.length / Object.keys(set).length };
}

// Batch update rows with different values per condition
export async function batchUpdateRowsViaAppleScript(
  filePath: string,
  sheetName: string | undefined,
  updatesList: Array<{ where: WhereCondition; set: { [column: string]: unknown } }>
): Promise<{ updatedCount: number; details: Array<{ where: WhereCondition; matched: number }> }> {
  // Read the file once
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, sheetName);
  const { headers, rows } = sheetToTable(ws);

  // Pre-compute column indices
  const allCols = new Set<string>();
  for (const { set } of updatesList) {
    for (const col of Object.keys(set)) {
      allCols.add(col);
    }
  }

  const colIndices: { [col: string]: number } = {};
  for (const col of allCols) {
    const idx = headers.indexOf(col);
    if (idx === -1) {
      throw new Error(`Column '${col}' not found. Available: ${headers.join(", ")}`);
    }
    colIndices[col] = idx + 1; // 1-based
  }

  // Build all updates
  const cellUpdates: Array<{ rowIndex: number; colIndex: number; value: unknown }> = [];
  const details: Array<{ where: WhereCondition; matched: number }> = [];
  let totalUpdated = 0;

  for (const { where, set } of updatesList) {
    let matched = 0;
    for (let i = 0; i < rows.length; i++) {
      if (matchesCondition(rows[i], where)) {
        const rowIndex = i + 2;
        for (const [col, value] of Object.entries(set)) {
          cellUpdates.push({
            rowIndex,
            colIndex: colIndices[col],
            value,
          });
        }
        matched++;
        totalUpdated++;
      }
    }
    details.push({ where, matched });
  }

  if (cellUpdates.length > 0) {
    await updateCellsBatch(filePath, cellUpdates);
  }

  return { updatedCount: totalUpdated, details };
}

// Close Numbers document (optional cleanup)
export async function closeDocument(filePath: string): Promise<void> {
  const absPath = path.resolve(filePath);

  const script = `
tell application "Numbers"
  set theDoc to document "${path.basename(absPath)}"
  close theDoc saving yes
end tell
  `.trim();

  try {
    await runAppleScript(script);
  } catch {
    // Ignore errors if document isn't open
  }
}

// Convert Numbers' 0-65535 RGB back to 0-255 scale
function fromNumbersColor(r: number, g: number, b: number): RGBColor {
  return {
    r: Math.round((r / 65535) * 255),
    g: Math.round((g / 65535) * 255),
    b: Math.round((b / 65535) * 255),
  };
}

// Check if a color is essentially white/no color (default background)
function isWhiteOrDefault(color: RGBColor): boolean {
  // Allow some tolerance for near-white colors
  return color.r >= 250 && color.g >= 250 && color.b >= 250;
}

// Create a color key for grouping (rounds to reduce minor variations)
function colorKey(color: RGBColor): string {
  const r = Math.round(color.r / 10) * 10;
  const g = Math.round(color.g / 10) * 10;
  const b = Math.round(color.b / 10) * 10;
  return `${r},${g},${b}`;
}

// Read background colors of all rows via AppleScript
export async function getRowColors(
  filePath: string,
  rowCount: number
): Promise<(RGBColor | null)[]> {
  const absPath = path.resolve(filePath);

  // AppleScript to get all row colors at once
  const script = `
tell application "Numbers"
  activate
  set theDoc to open POSIX file "${absPath}"
  delay 0.5
  set colorList to {}
  tell theDoc
    tell sheet 1
      tell table 1
        repeat with i from 2 to ${rowCount + 1}
          try
            set rowColor to background color of row i
            set end of colorList to rowColor
          on error
            set end of colorList to {65535, 65535, 65535}
          end try
        end repeat
      end tell
    end tell
  end tell
  return colorList
end tell
  `.trim();

  const output = await runAppleScript(script);

  // Parse AppleScript output: {{r, g, b}, {r, g, b}, ...}
  const colors: (RGBColor | null)[] = [];

  // Extract color tuples from the output
  const colorMatches = output.match(/\d+\s*,\s*\d+\s*,\s*\d+/g);

  if (colorMatches) {
    for (const match of colorMatches) {
      const [r, g, b] = match.split(",").map((s) => parseInt(s.trim(), 10));
      const color = fromNumbersColor(r, g, b);
      colors.push(isWhiteOrDefault(color) ? null : color);
    }
  }

  return colors;
}

export interface ColorMappingResult {
  colorMap: { [value: string]: RGBColor | null };
  confidence: number;
  sampleSize: number;
}

// Detect color mapping by correlating row colors with column values
export async function detectColorMapping(
  filePath: string,
  sheetName: string | undefined,
  column: string
): Promise<ColorMappingResult> {
  // Read the data to get column values
  const wb = readWorkbook(filePath);
  const ws = getSheet(wb, sheetName);
  const { headers, rows } = sheetToTable(ws);

  // Verify column exists
  if (!headers.includes(column)) {
    throw new Error(`Column '${column}' not found. Available: ${headers.join(", ")}`);
  }

  // Get row colors via AppleScript
  const colors = await getRowColors(filePath, rows.length);

  // Build value -> colors mapping
  const valueColors: Map<string, RGBColor[]> = new Map();
  const valueNullCount: Map<string, number> = new Map();
  let coloredRows = 0;

  for (let i = 0; i < rows.length; i++) {
    const value = String(rows[i][column] ?? "");
    const color = colors[i];

    if (color) {
      coloredRows++;
      if (!valueColors.has(value)) {
        valueColors.set(value, []);
      }
      valueColors.get(value)!.push(color);
    } else {
      valueNullCount.set(value, (valueNullCount.get(value) || 0) + 1);
    }
  }

  // Determine the most common color for each value
  const colorMap: { [value: string]: RGBColor | null } = {};
  let consistentMappings = 0;
  let totalMappings = 0;

  // Get all unique values
  const allValues = new Set([...valueColors.keys(), ...valueNullCount.keys()]);

  for (const value of allValues) {
    const colors = valueColors.get(value) || [];
    const nullCount = valueNullCount.get(value) || 0;

    if (colors.length === 0) {
      // All instances have no color
      colorMap[value] = null;
      consistentMappings++;
      totalMappings++;
    } else {
      // Find most common color for this value
      const colorCounts: Map<string, { count: number; color: RGBColor }> = new Map();

      for (const color of colors) {
        const key = colorKey(color);
        if (!colorCounts.has(key)) {
          colorCounts.set(key, { count: 0, color });
        }
        colorCounts.get(key)!.count++;
      }

      // Get the most common color
      let maxCount = 0;
      let dominantColor: RGBColor | null = null;

      for (const { count, color } of colorCounts.values()) {
        if (count > maxCount) {
          maxCount = count;
          dominantColor = color;
        }
      }

      colorMap[value] = dominantColor;

      // Check consistency (what % of this value's rows have the dominant color)
      const totalForValue = colors.length + nullCount;
      const consistency = maxCount / totalForValue;
      if (consistency >= 0.8) {
        consistentMappings++;
      }
      totalMappings++;
    }
  }

  const confidence = totalMappings > 0 ? consistentMappings / totalMappings : 0;

  return {
    colorMap,
    confidence: Math.round(confidence * 100) / 100,
    sampleSize: coloredRows,
  };
}
