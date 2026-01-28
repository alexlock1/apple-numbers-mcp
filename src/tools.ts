import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  getWorkbookInfo,
  getTableData,
  findTableRows,
  WhereCondition,
} from "./lib/numbers.js";
import {
  setRowColor,
  colorRowsByValue,
  detectColorMapping,
  updateRowsViaAppleScript,
  batchUpdateRowsViaAppleScript,
  addRowsViaAppleScript,
  deleteRowsViaAppleScript,
  RGBColor,
} from "./lib/applescript.js";

// Schema for where conditions
const whereValueSchema = z.union([
  z.string(),
  z.number(),
  z.boolean(),
  z.null(),
  z.object({
    op: z.enum(["contains", "gt", "lt", "gte", "lte", "ne"]),
    value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
  }),
]);

const whereSchema = z.record(z.string(), whereValueSchema);

// Schema for cell values
const cellValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

export function registerTools(server: McpServer): void {
  // 1. numbers_get_info - Quick overview
  server.registerTool(
    "numbers_get_info",
    {
      description:
        "Get a quick overview of a Numbers file: sheet names, headers, row count, and first 3 rows as preview",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
      },
    },
    async ({ filePath }) => {
      try {
        const info = getWorkbookInfo(filePath);
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(info, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // 2. numbers_get_table - Get table data with pagination
  server.registerTool(
    "numbers_get_table",
    {
      description:
        "Get table data from a Numbers file with optional pagination and column filtering",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
        sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
        columns: z
          .array(z.string())
          .optional()
          .describe("Only return these columns (omit for all columns)"),
        limit: z.number().optional().describe("Maximum number of rows to return"),
        offset: z.number().optional().describe("Number of rows to skip (for pagination)"),
      },
    },
    async ({ filePath, sheetName, columns, limit, offset }) => {
      try {
        const data = getTableData(filePath, { sheetName, columns, limit, offset });
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(data, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // 3. numbers_find_rows - Find rows matching conditions
  server.registerTool(
    "numbers_find_rows",
    {
      description:
        "Find rows in a Numbers table matching conditions. Supports equality (Status: 'Done') and operators (Priority: {op: 'gt', value: 5})",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
        sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
        where: whereSchema.describe(
          "Conditions to match. Examples: {Status: 'Todo'} or {Name: {op: 'contains', value: 'test'}}"
        ),
        columns: z
          .array(z.string())
          .optional()
          .describe("Only return these columns (omit for all columns)"),
        limit: z.number().optional().describe("Maximum number of rows to return"),
        offset: z.number().optional().describe("Number of rows to skip (for pagination)"),
      },
    },
    async ({ filePath, sheetName, where, columns, limit, offset }) => {
      try {
        const result = findTableRows(filePath, where as WhereCondition, {
          sheetName,
          columns,
          limit,
          offset,
        });
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // 4. numbers_update_rows - Update rows matching conditions
  server.registerTool(
    "numbers_update_rows",
    {
      description:
        "Update rows in a Numbers table that match conditions. Sets ALL matching rows to the SAME values. For different values per row, use numbers_batch_update instead. Changes are saved to the file.",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
        sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
        where: whereSchema.describe("Conditions to match rows for update"),
        set: z.record(z.string(), cellValueSchema).describe("Column values to set on matching rows"),
      },
    },
    async ({ filePath, sheetName, where, set }) => {
      try {
        const result = await updateRowsViaAppleScript(filePath, sheetName, where as WhereCondition, set);
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // 5. numbers_add_row - Add rows to table
  server.registerTool(
    "numbers_add_row",
    {
      description: "Add one or more rows to a Numbers table. Changes are saved to the file.",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
        sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
        rows: z
          .union([
            z.record(z.string(), cellValueSchema),
            z.array(z.record(z.string(), cellValueSchema)),
          ])
          .describe("Row(s) to add as object(s) with column names as keys"),
      },
    },
    async ({ filePath, sheetName, rows }) => {
      try {
        const result = await addRowsViaAppleScript(filePath, sheetName, rows);
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // 6. numbers_delete_rows - Delete rows matching conditions
  server.registerTool(
    "numbers_delete_rows",
    {
      description:
        "Delete rows from a Numbers table that match conditions. Changes are saved to the file.",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
        sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
        where: whereSchema.describe("Conditions to match rows for deletion"),
      },
    },
    async ({ filePath, sheetName, where }) => {
      try {
        const result = await deleteRowsViaAppleScript(filePath, sheetName, where as WhereCondition);
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // 6b. numbers_batch_update - Update multiple rows with different values
  server.registerTool(
    "numbers_batch_update",
    {
      description:
        "Update multiple rows with DIFFERENT values in a single write operation. Each update specifies its own where condition and set values. More efficient than multiple numbers_update_rows calls when setting different values per row.",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
        sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
        updates: z
          .array(
            z.object({
              where: whereSchema.describe("Conditions to match rows"),
              set: z.record(z.string(), cellValueSchema).describe("Column values to set"),
            })
          )
          .describe(
            "Array of {where, set} pairs. Each pair updates matching rows with its own values."
          ),
      },
    },
    async ({ filePath, sheetName, updates }) => {
      try {
        const result = await batchUpdateRowsViaAppleScript(
          filePath,
          sheetName,
          updates as Array<{ where: WhereCondition; set: { [column: string]: unknown } }>
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // 7. numbers_set_row_color - Set row background color (requires macOS)
  server.registerTool(
    "numbers_set_row_color",
    {
      description:
        "Set the background color of a specific row. Requires macOS (uses AppleScript). Numbers app will open briefly.",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
        sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
        rowIndex: z.number().describe("1-based row index (row 1 is the header)"),
        color: z
          .object({
            r: z.number().min(0).max(255).describe("Red component (0-255)"),
            g: z.number().min(0).max(255).describe("Green component (0-255)"),
            b: z.number().min(0).max(255).describe("Blue component (0-255)"),
          })
          .describe("RGB color values"),
      },
    },
    async ({ filePath, sheetName, rowIndex, color }) => {
      try {
        await setRowColor(filePath, sheetName, rowIndex, color as RGBColor);
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ success: true, rowIndex, color }, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // 8. numbers_color_rows_by_value - Color rows based on column value
  server.registerTool(
    "numbers_color_rows_by_value",
    {
      description:
        "Color multiple rows based on a column's value. For example, color all 'Done' rows green and 'In Progress' rows blue. Requires macOS.",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
        sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
        column: z.string().describe("Column name to check for values"),
        colorMap: z
          .record(
            z.string(),
            z.object({
              r: z.number().min(0).max(255),
              g: z.number().min(0).max(255),
              b: z.number().min(0).max(255),
            })
          )
          .describe(
            "Map of column values to RGB colors. Example: {'Done': {r: 144, g: 238, b: 144}, 'In Progress': {r: 173, g: 216, b: 230}}"
          ),
      },
    },
    async ({ filePath, sheetName, column, colorMap }) => {
      try {
        const result = await colorRowsByValue(
          filePath,
          sheetName,
          column,
          colorMap as { [value: string]: RGBColor }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  // 9. numbers_detect_color_mapping - Auto-detect color scheme from existing styling
  server.registerTool(
    "numbers_detect_color_mapping",
    {
      description:
        "Auto-detect color mapping by reading existing row colors and correlating with column values. Returns a colorMap that can be used with numbers_color_rows_by_value. Requires macOS.",
      inputSchema: {
        filePath: z.string().describe("Absolute path to the .numbers file"),
        sheetName: z.string().optional().describe("Sheet name (defaults to first sheet)"),
        column: z.string().describe("Column name to correlate colors with (e.g., 'Status')"),
      },
    },
    async ({ filePath, sheetName, column }) => {
      try {
        const result = await detectColorMapping(filePath, sheetName, column);
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error: ${(error as Error).message}` }],
          isError: true,
        };
      }
    }
  );
}
