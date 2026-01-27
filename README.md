# Numbers MCP Server

An MCP (Model Context Protocol) server for reading and writing Apple Numbers spreadsheets.

## Features

- **Read Operations**: Get table info, fetch rows with pagination, find rows with flexible conditions
- **Write Operations**: Update rows, add rows, delete rows, batch updates with different values per row
- **Styling** (macOS only): Set row colors, color rows by column value, auto-detect existing color schemes

## Tools

| Tool | Description |
|------|-------------|
| `numbers_get_info` | Quick overview of a Numbers file (sheets, headers, row count, preview) |
| `numbers_get_table` | Get table data with optional pagination and column filtering |
| `numbers_find_rows` | Find rows matching conditions (equality, contains, gt, lt, gte, lte, ne) |
| `numbers_update_rows` | Update all rows matching a condition to the same values |
| `numbers_add_row` | Add one or more rows to a table |
| `numbers_delete_rows` | Delete rows matching a condition |
| `numbers_batch_update` | Update multiple rows with different values in one operation |
| `numbers_set_row_color` | Set background color of a specific row |
| `numbers_color_rows_by_value` | Color rows based on a column's value (e.g., status) |
| `numbers_detect_color_mapping` | Auto-detect color scheme from existing styling |

## Installation

Clone and build:

```bash
git clone https://github.com/mycleaningfam/numbers-mcp-server.git
cd numbers-mcp-server
npm install
npm run build
```

## Configuration

Add to your MCP settings (e.g., `.mcp.json` or Claude Desktop config):

```json
{
  "mcpServers": {
    "numbers": {
      "command": "node",
      "args": ["/path/to/numbers-mcp-server/dist/index.js"]
    }
  }
}
```

## Usage Examples

### Get table overview
```
numbers_get_info({ filePath: "/path/to/file.numbers" })
```

### Find rows with conditions
```
numbers_find_rows({
  filePath: "/path/to/file.numbers",
  where: { Status: "In Progress" },
  columns: ["ID", "Task", "Status"],
  limit: 10
})
```

### Update rows
```
numbers_update_rows({
  filePath: "/path/to/file.numbers",
  where: { Status: "Done" },
  set: { Priority: "P1" }
})
```

### Batch update with different values
```
numbers_batch_update({
  filePath: "/path/to/file.numbers",
  updates: [
    { where: { ID: 1 }, set: { Priority: "P0" } },
    { where: { ID: 2 }, set: { Priority: "P1" } },
    { where: { ID: 3 }, set: { Priority: "P2" } }
  ]
})
```

### Color rows by status
```
numbers_color_rows_by_value({
  filePath: "/path/to/file.numbers",
  column: "Status",
  colorMap: {
    "Done": { r: 144, g: 238, b: 144 },
    "In Progress": { r: 173, g: 216, b: 230 }
  }
})
```

## Requirements

- Node.js 18+
- macOS (required for styling features - uses AppleScript to control Numbers app)
- Numbers app installed (for styling features)

## How It Works

- **Reading**: Uses SheetJS (xlsx) library to parse .numbers files
- **Writing**: Uses AppleScript to update cells via the Numbers app (SheetJS free version cannot write .numbers format)
- **Styling**: Uses AppleScript to set row background colors through Numbers app automation

## License

MIT
