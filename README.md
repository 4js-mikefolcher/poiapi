# POI API for Genero BDL

A Genero BDL package for creating Microsoft Excel (.xlsx) files using the Apache POI Java library.

## Features

- **TSpreadsheet** - Simple, flat data export with automatic data type formatting
- **TSpreadsheetXtend** - Grouped data with subtotals, subtitles, multi-level nesting, and multi-sheet workbooks
- **tableExcelExport** - One-call export directly from a UI Table widget, with support for column reordering, sorting, and aggregate totals
- Automatic Excel formatting for all Genero data types (MONEY, DECIMAL, INTEGER, FLOAT, DATE, DATETIME, etc.)
- Dates and currency written in the locale of the machine the Genero application runs on, honouring `DBDATE`, `DBFORMAT` and `DBMONEY`, with a named locale, explicit format codes or viewer-resolved formatting as alternatives
- Native Excel formula generation for subtotals (SUM, AVG, COUNT, MIN, MAX)
- Page setup with landscape orientation, frozen panes, repeating headers, and page numbers

## Installation

### Genero Package Manager (fglpkg)

The recommended installation method is via the Genero Package Manager. You can download `fglpkg` from https://github.com/4js-mikefolcher/fglpkg/releases.

Add `poiapi` as a dependency in your project's `fglpkg.json`:

```json
{
  "name": "my-project",
  "version": "1.0.0",
  "dependencies": {
    "fgl": {
      "poiapi": "^1.9.0"
    }
  }
}
```

Run `fglpkg install` to download the package and its Java dependencies. The exact
versions are declared in [fglpkg.json](fglpkg.json) and pinned in `fglpkg-lock.json`;
[USERGUIDE.md](USERGUIDE.md#understanding-what-fglpkg-installs) lists them.

### Manual Installation

1. Download Apache POI and its transitive dependencies from
   [Maven Central](https://search.maven.org/) — the full list, with versions, is in
   [USERGUIDE.md](USERGUIDE.md#understanding-what-fglpkg-installs)
2. Add every one of those JAR files to your `CLASSPATH`
3. Copy the compiled `.42m` files from `com/fourjs/poiapi/` to a directory on your `FGLLDPATH`

## Quick Start

### Simple Export (TSpreadsheet)

```4gl
IMPORT util
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_api

DEFINE excelHandler fgl_spreadsheet_api.TSpreadsheet

CALL excelHandler.init()
CALL excelHandler.setHeaders(myHeaders)
CALL excelHandler.setRecordDefinition(base.TypeInfo.create(myRecord))
CALL excelHandler.setTitle("My Report")

IF excelHandler.createSpreadsheet(util.JSONArray.fromFGL(myDataArray)) THEN
    DISPLAY SFMT("Created: %1", excelHandler.getFilename())
END IF
```

### Grouped Export with Subtotals (TSpreadsheetXtend)

```4gl
IMPORT util
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_xapi
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_helper

DEFINE excelHandler fgl_spreadsheet_xapi.TSpreadsheetXtend

CALL excelHandler.init()
CALL excelHandler.setColumnInfo(myColumnInfo)
CALL excelHandler.setRecordDefinition(base.TypeInfo.create(myRecord))
CALL excelHandler.setTitle("Grouped Report")
CALL excelHandler.setGroupColumn(TRUE)
CALL excelHandler.addSubTitle("My Report Title")

# Add rows with group boundaries
CALL excelHandler.addGroupHeaderRow("dept", "Sales Department")
FOR idx = 1 TO salesData.getLength()
    CALL excelHandler.addDataRow(util.JSONObject.fromFGL(salesData[idx]))
END FOR
CALL excelHandler.addGroupFooterRow("dept")

IF excelHandler.createSpreadsheet() THEN
    DISPLAY SFMT("Created: %1", excelHandler.getFilename())
END IF
```

### UI Table Export (tableExcelExport)

```4gl
IMPORT util
IMPORT FGL com.fourjs.poiapi.fgl_table_export

DISPLAY ARRAY dataList TO s_table.*
    ON ACTION export_to_excel
        VAR filename = fgl_table_export.tableExcelExport(
            "s_table", util.JSONArray.fromFGL(dataList))
        CALL fgl_putfile(filename, "gbc")
END DISPLAY
```

### Controlling the Formats

Dates and money follow `DBDATE`, `DBFORMAT` and `DBMONEY`, falling back to the
locale of the machine the program runs on. Call these before building a
spreadsheet to override that; they apply to all three APIs.

```4gl
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_helper

# Export in a named locale rather than the machine's
CALL fgl_spreadsheet_helper.setLocale("de-DE")

# Or pin one column type to a format code of your own
CALL fgl_spreadsheet_helper.setDateFormat("yyyy-mm-dd")

# Or hand the whole decision back to whoever opens the file
CALL fgl_spreadsheet_helper.setFormatMode(
    fgl_spreadsheet_helper.cFormatModeViewer)
```

`getAvailableLocales()` returns every locale the runtime can format for, ready
to put in front of the user — see
[USERGUIDE.md](USERGUIDE.md#overriding-the-formats).

## Package Structure

```
com.fourjs.poiapi/
├── fgl_excel.4gl                  # Low-level Apache POI Java wrapper
├── fgl_structures.4gl             # Row stack and range utilities for grouping
├── fgl_spreadsheet_helper.4gl     # Shared types, constants, and formatting utilities
├── fgl_spreadsheet_interface.4gl  # ISpreadsheet interface definition
├── fgl_spreadsheet_api.4gl        # TSpreadsheet API (simple export)
├── fgl_spreadsheet_xapi.4gl       # TSpreadsheetXtend API (grouped/multi-sheet export)
└── fgl_table_export.4gl           # UI Table export function
```

## Test Program

The `src/` directory contains `fgl_excel_api_test.4gl`, a test program with six examples:

1. **excelAPIExample** - Basic TSpreadsheet export with all data types
2. **excelXAPIExample** - TSpreadsheetXtend with multi-level grouping (5/10/20 row intervals) and subtotals
3. **excelMultisheetExample** - Multi-sheet workbook with four sheets using different record types and configurations
4. **excelTable** - UI Table export without aggregates
5. **xtendExcelTable** - UI Table export with aggregate totals defined in the form
6. **excelFormatModesExample** - One workbook whose eight sheets each carry the same data under a different format mode, override or locale, so they can be compared side by side

Every table screen also carries an **Export to Excel (locale...)** action, which
opens a modal picker over `getAvailableLocales()` and exports in the locale
chosen — a worked example of a locale picker.

## Documentation

See [USERGUIDE.md](USERGUIDE.md) for detailed API documentation, method references, and complete code examples.

## Credits

The `fgl_excel.4gl` module originated from Reuben's [fgl_apache_poi](https://github.com/FourjsGenero/fgl_apache_poi) repository. This project builds on that foundation with a higher-level API focused on Excel export, updated for Apache POI 5.x, and extended with grouping, subtotals, multi-sheet support, and UI table integration.

## License

4Js License
