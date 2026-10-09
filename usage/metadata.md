# Read spreadsheet metadata

Configure and authenticate the Sheets server as described in the [README](../README.md). Use `spreadsheet-find` to get the workbook `id` and sheet `gid`. Both metadata tools read existing spreadsheets.

| Tool | Input | Selection |
| --- | --- | --- |
| `spreadsheet-metadata-get` | `id` | Workbook structure across all sheets, without grid data |
| `cells-metadata-get` | `id`, `gid`, `ranges` | Selected metadata in finite ranges on one sheet |

## Workbook structure

Call `spreadsheet-metadata-get` with the workbook ID:

```json
{"id":"spreadsheet-id"}
```

It returns full workbook and sheet properties, including available locale, time zone, hidden state, dimensions, and frozen rows or columns. Selected structural collections include named ranges, charts, conditional formats, protections, tables, basic filters, filter views, banded ranges, merges, slicers, developer metadata, data sources, and data-source schedules.

A success result excerpt for a workbook with two sheets:

```json
{
  "result": {
    "type": "success",
    "id": "spreadsheet-id",
    "metadata": {
      "spreadsheetId": "spreadsheet-id",
      "properties": {"title": "Quarterly Report", "locale": "en_US"},
      "sheets": [
        {"properties": {"sheetId": 0, "title": "Sales"}},
        {"properties": {"sheetId": 123, "title": "Summary"}}
      ]
    }
  }
}
```

This selection excludes grid data, cell values, formulas, direct cell formatting, and row or column dimension metadata. It is a structural overview, not a full workbook export.

## Metadata in cells

Call `cells-metadata-get` with the workbook ID, sheet ID, and required sheet-local ranges:

```json
{"id":"spreadsheet-id","gid":"0","ranges":["B2:D4"]}
```

Each range must be a single cell such as `A1` or a rectangle with finite endpoints such as `B2:D4`. Omit the sheet title. Whole rows, whole columns, sheet-qualified ranges, reversed endpoints, and empty ranges are invalid. To inspect another sheet, make another call with its `gid`.

The package permits at most 50 ranges and 250,000 aggregate cells per call. Overlapping ranges count separately. These limits bound the requested cells, not the response size. Sheets of type `DATA_SOURCE` or `OBJECT` are unsupported.

The cell fields are `note`, `dataValidation`, `pivotTable`, and `dataSourceTable`. Values, formulas, direct formatting, and workbook structures are excluded. Google returns definitions only where they exist, including the anchor cell for a pivot table.

A success result excerpt with a note in C2:

```json
{
  "result": {
    "type": "success",
    "id": "spreadsheet-id",
    "gid": "0",
    "requestedRanges": ["'Sales'!B2:D4"],
    "metadata": {
      "spreadsheetId": "spreadsheet-id",
      "sheets": [{
        "properties": {"sheetId": 0, "title": "Sales"},
        "data": [{
          "startRow": 1,
          "startColumn": 1,
          "rowData": [{"values": [{}, {"note": "Review this figure"}]}]
        }]
      }]
    }
  }
}
```

`startRow` and `startColumn` are zero-based offsets. An omitted offset means zero. Add the `rowData` array index to `startRow` and the `values` array index to `startColumn` to locate each cell. Keep empty row and cell objects in place; trailing cells or metadata fields may be absent. Do not treat the result as a dense rectangle or infer cell values from missing metadata.

## Results and errors

Both tools return the result as JSON text in MCP `content` and as `structuredContent.result`. Successful results contain `type: "success"`, the input IDs, `metadata`, and `fieldMask`, the field selection sent to Google. The excerpts above omit `fieldMask` and unrelated returned properties. Cell results also contain sheet-qualified `requestedRanges` in request order; the tool resolves and quotes the sheet title for you.

An unauthenticated call can return `type: "auth_required"`; follow the server's authentication instructions before retrying. Invalid cell ranges, limits, missing sheets, and unsupported sheet types produce an invalid-parameters error. Provider failures remain errors.
