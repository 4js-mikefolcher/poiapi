# POI API for Genero BDL

A Genero BDL package for creating Microsoft Excel (.xlsx) files using the Apache POI Java library.

## Features

- **TSpreadsheet** - Simple, flat data export with automatic data type formatting
- **TSpreadsheetXtend** - Grouped data with subtotals, subtitles, multi-level nesting, and multi-sheet workbooks
- **tableExcelExport** - One-call export directly from a UI Table widget, with support for column reordering, sorting, and aggregate totals
- Automatic Excel formatting for all Genero data types (MONEY, DECIMAL, INTEGER, FLOAT, DATE, DATETIME, etc.)
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
      "poiapi": "6.0.0"
    }
  }
}
```

Run `fglpkg install` to download the package and its Java dependencies (Apache POI 5.2.3, Log4j 2.25.3).

### Manual Installation

1. Download the Apache POI libraries from https://poi.apache.org/
2. Set the `POI_HOME` environment variable to the download location
3. Add the POI JAR files to your `CLASSPATH`:
   ```
   $(CLASSPATH);$(POI_DIR)/poiapi-4js-5.2.3.jar;$(POI_DIR)/log4j-core-2.19.0.jar
   ```

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

The `src/` directory contains `fgl_excel_api_test.4gl`, a test program with five examples:

1. **excelAPIExample** - Basic TSpreadsheet export with all data types
2. **excelXAPIExample** - TSpreadsheetXtend with multi-level grouping (5/10/20 row intervals) and subtotals
3. **excelMultisheetExample** - Multi-sheet workbook with four sheets using different record types and configurations
4. **excelTable** - UI Table export without aggregates
5. **xtendExcelTable** - UI Table export with aggregate totals defined in the form

## Documentation

See [USERGUIDE.md](USERGUIDE.md) for detailed API documentation, method references, and complete code examples.

## Credits

The `fgl_excel.4gl` module originated from Reuben's [fgl_apache_poi](https://github.com/FourjsGenero/fgl_apache_poi) repository. This project builds on that foundation with a higher-level API focused on Excel export, updated for Apache POI 5.2.3, and extended with grouping, subtotals, multi-sheet support, and UI table integration.

## License

4Js License
