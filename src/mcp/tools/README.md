# MCP Tools Design

Docs: https://mcp-z.github.io/mcp-sheets
## Spreadsheet/Sheet Identification Pattern

### Design Principle

**Lookup tools handle flexible resolution. Operation tools expect direct IDs.**

```
LOOKUP TOOLS (flexible input → concrete IDs output)
├── spreadsheet-find  → accepts SpreadsheetRefSchema (URL/name/ID)
│                     → returns { id, gid[] } for all sheets
└── sheet-find        → accepts id + SheetRefSchema (title/gid)
                      → returns { id, gid, title }

OPERATION TOOLS (require direct IDs from lookup tools)
└── All other tools   → accept SpreadsheetIdSchema + SheetGidSchema
                      → NO flexible resolution, use IDs directly
```

### Rationale

1. **Clear separation of concerns**: Lookup tools find things, operation tools do things
2. **Reduced API calls**: Operation tools make 1 call instead of 2-3
3. **Schema matches implementation**: When schema says "ID", the code expects an ID
4. **Predictable agent workflow**: Find first, operate second

### Agent Workflow

```
1. Agent receives task: "Add data to the Sales spreadsheet"

2. Agent calls spreadsheet-find:
   Input:  { spreadsheetRef: "Sales" }
   Output: { id: "abc123...", sheets: [{ gid: "0", title: "Q1" }, ...] }

3. Agent calls operation tool with concrete IDs:
   Input:  { id: "abc123...", gid: "0", ... }
   Output: { success, ... }
```

### Schema Definitions

| Schema | Description | Used By |
|--------|-------------|---------|
| `SpreadsheetRefSchema` | Flexible: URL, name, or ID | `spreadsheet-find` only |
| `SheetRefSchema` | Flexible: title or gid | `sheet-find` only |
| `SpreadsheetIdSchema` | Direct: spreadsheet ID from URL `d/{id}` | All operation tools |
| `SheetGidSchema` | Direct: sheet ID from URL `gid={gid}` | All operation tools |

### Tool Categories

#### Lookup Tools (2)
- `spreadsheet-find` - Find spreadsheet by URL/name/ID, returns all sheet metadata
- `sheet-find` - Find sheet by title/gid within a known spreadsheet

#### Spreadsheet Operations (4)
- `spreadsheet-create` - Create new spreadsheet
- `spreadsheet-copy` - Copy entire spreadsheet
- `spreadsheet-rename` - Rename spreadsheet

#### Sheet Operations (5)
- `sheet-create` - Create new sheet in spreadsheet
- `sheet-delete` - Delete sheets by gid
- `sheet-rename` - Rename sheet
- `sheet-copy` - Copy sheet within same spreadsheet
- `sheet-copy-to` - Copy sheet to different spreadsheet

#### Data Operations (12)
- `rows-get` - Read row data from range
- `rows-append` - Append rows with deduplication
- `rows-csv-append` - Append rows from CSV
- `columns-get` - Get first row (headers)
- `columns-update` - Upsert rows by key columns
- `values-search` - Search for values
- `values-batch-update` - Update multiple ranges
- `values-clear` - Clear cell values
- `values-csv-update` - Update from CSV
- `values-replace` - Find and replace values

#### Formatting Operations (3)
- `cells-format` - Apply cell formatting
- `validation-set` - Set data validation rules
- `chart-create` - Create charts

#### Dimension Operations (2)
- `dimensions-batch-update` - Insert/delete/append rows/columns
- `dimensions-move` - Move rows/columns

#### Utility (1)
- `csv-get-columns` - Get columns from external CSV (not a spreadsheet operation)
