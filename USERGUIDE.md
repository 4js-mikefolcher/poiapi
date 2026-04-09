# POI API User Guide

## Overview

The POI API package (`com.fourjs.poiapi`) provides a Genero BDL library for creating Microsoft Excel (.xlsx) files using the Apache POI Java library. The package offers three levels of abstraction:

| API | Module | Use Case |
|-----|--------|----------|
| **TSpreadsheet** | `fgl_spreadsheet_api` | Simple, flat data export with automatic formatting |
| **TSpreadsheetXtend** | `fgl_spreadsheet_xapi` | Grouped data with subtotals, subtitles, and multi-sheet workbooks |
| **tableExcelExport** | `fgl_table_export` | One-call export directly from a UI Table widget |

---

## Installation

### Prerequisites

- Genero BDL 4.x or later
- Java Runtime Environment

### Using the Genero Package Manager (fglpkg)

The recommended way to install the POI API package is through the Genero Package Manager (`fglpkg`).

#### Step 1: Install fglpkg

Download the `fglpkg` binary for your platform from [GitHub Releases](https://github.com/4js-mikefolcher/fglpkg/releases) and place it in a directory on your `PATH`:

```bash
# macOS/Linux example
sudo cp fglpkg-darwin-arm64 /usr/local/bin/fglpkg
sudo chmod +x /usr/local/bin/fglpkg
```

#### Step 2: Understand local vs. global installation

`fglpkg` supports two installation scopes. The choice between them depends on how your environment manages Genero versions.

| Scope | Packages stored in | Environment command | Best for |
|-------|-------------------|---------------------|----------|
| **Local** (`--local`) | `.fglpkg/` inside the project directory | `eval "$(fglpkg env)"` | Development workstations where multiple Genero versions may coexist |
| **Global** (`--global`) | `~/.fglpkg/` (user home) | `eval "$(fglpkg env --global)"` | CI/CD pipelines and build servers with a single Genero version |

**Local installation** is recommended for individual development environments. Because compiled BDL modules (`.42m` files) are tied to a specific Genero major version, keeping packages local to each project avoids conflicts when different projects target different Genero versions. Each project gets its own isolated set of packages and JARs.

**Global installation** is better suited for CI/CD pipelines and build servers where a single Genero version is installed. Packages are shared across all projects, avoiding redundant downloads on every build. The global home directory can be overridden with the `FGLPKG_HOME` environment variable if needed.

When you run `fglpkg install` from within a project directory (one that contains a `fglpkg.json`), it defaults to local installation. Outside a project directory, it defaults to global. You can override the default with the `--local` or `--global` flag.

#### Step 3: Configure your shell environment

Add one of the following lines to your shell profile (`~/.bashrc`, `~/.zshrc`, or equivalent) so that installed packages are available to Genero at compile and runtime:

**For local (development) installation:**

```bash
# Run from within your project directory, or set up per-project scripts
eval "$(fglpkg env)"

# If using GST, run this command and copy the output to put into your project environment variables
fglpkg env --gst
```

Since local environment paths are project-specific, you typically run this in your project's build script or terminal session rather than in your global shell profile.

**For global (CI/CD) installation:**

```bash
# Add to shell profile for build servers
eval "$(fglpkg env --global)"
```

Both commands export the `FGLLDPATH` and `CLASSPATH` environment variables with paths to the installed BDL packages and Java JAR files. Reload your shell or run `source ~/.bashrc` (or `source ~/.zshrc`) after making changes to your profile.

#### Step 4: Add poiapi as a dependency

If your project does not already have a `fglpkg.json` file, create one by running:

```bash
cd /path/to/your/project
fglpkg init
```

Then add `poiapi` as a dependency. You can do this by running:

```bash
# Local install (default when inside a project directory)
fglpkg install poiapi

# Or explicitly global for CI/CD
fglpkg install --global poiapi
```

Or by manually editing your project's `fglpkg.json` to include the dependency:

```json
{
  "name": "my-project",
  "version": "1.0.0",
  "dependencies": {
    "fgl": {
      "poiapi": "^1.2.0"
    }
  }
}
```

Then run:

```bash
fglpkg install
```

The package manager resolves and downloads both the compiled BDL modules and all required Java dependencies (Apache POI 5.3.0 and its transitive dependencies) automatically. The `fglpkg.lock` file is created alongside `fglpkg.json` to pin exact versions for reproducible installs across machines.

#### Step 5: Verify the installation

Confirm the package is installed:

```bash
fglpkg list
```

You should see `poiapi` listed with its version number.

#### Understanding what fglpkg installs

The `poiapi` package declares its Java dependencies in its own `fglpkg.json`. When you install `poiapi`, `fglpkg` automatically fetches these JAR files from Maven Central:

| Library | Version | Purpose |
|---------|---------|---------|
| poi | 5.3.0 | Base POI library |
| poi-ooxml | 5.3.0 | OOXML support for .xlsx format |
| poi-ooxml-lite | 5.3.0 | Lightweight OOXML schemas |
| xmlbeans | 5.2.2 | XML binding layer |
| commons-compress | 1.27.1 | ZIP/package handling |
| commons-collections4 | 4.5.0 | Utility collections |
| commons-math3 | 3.6.1 | Math utilities |
| log4j-api | 2.17.1 | Logging facade |
| commons-io | 2.18.0 | I/O utilities |
| curvesapi | 1.08 | Curve/spline math |
| commons-codec | 1.17.2 | Encoding utilities |

You do not need to download or manage any of these manually.

### Manual Installation

If you are not using `fglpkg`, you can install the dependencies manually:

1. Download the Apache POI 5.3.0 libraries and all transitive dependencies listed above from [Maven Central](https://search.maven.org/) or the [Apache POI website](https://poi.apache.org/)
2. Add all JAR files to your `CLASSPATH` environment variable
3. Copy the compiled `.42m` files from `com/fourjs/poiapi/` to a directory on your `FGLLDPATH`

---

## Supported Data Types

The package automatically applies Excel formatting based on Genero data types:

| Genero Type | Excel Format |
|-------------|--------------|
| `MONEY` | Currency (e.g., `$1,234.56`) |
| `DECIMAL(p,s)` | Numeric with appropriate precision (e.g., `#,##0.0000`) |
| `INTEGER`, `SMALLINT` | Integer |
| `FLOAT`, `SMALLFLOAT` | Decimal |
| `DATE` | Date (e.g., `mm/dd/yyyy`) |
| `DATETIME YEAR TO SECOND` | Date and time (e.g., `mm/dd/yyyy hh:mm:ss AM/PM`) |
| `DATETIME HOUR TO SECOND` | Time only (e.g., `hh:mm:ss AM/PM`) |
| `STRING`, `VARCHAR`, `CHAR` | Plain text |

---

## API 1: TSpreadsheet (Simple Export)

Use `TSpreadsheet` when you need a straightforward export of a record array to a single Excel sheet with column headers and formatted data.

### Import

```4gl
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_api
```

### Steps

1. Define your data record and populate an array
2. Initialize the spreadsheet handler
3. Set column headers, record definition, and title
4. Call `createSpreadsheet()` with a JSON array of your data
5. Retrieve the output filename

### Example

```4gl
IMPORT util
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_api

# Step 1 - Define your data record
TYPE TEmployee RECORD
    emp_name   VARCHAR(50),
    department VARCHAR(30),
    salary     MONEY(10,2),
    hire_date  DATE
END RECORD

DEFINE employees DYNAMIC ARRAY OF TEmployee
DEFINE empRec TEmployee

MAIN

    # ... populate employees array ...

    # Step 2 - Initialize
    DEFINE excelHandler fgl_spreadsheet_api.TSpreadsheet
    CALL excelHandler.init()

    # Step 3 - Configure headers and record definition
    DEFINE headers DYNAMIC ARRAY OF STRING = [
        "Employee Name",
        "Department",
        "Salary",
        "Hire Date"
    ]
    CALL excelHandler.setHeaders(headers)
    CALL excelHandler.setRecordDefinition(base.TypeInfo.create(empRec))
    CALL excelHandler.setTitle("Employee Report")

    # Step 4 - Generate the spreadsheet
    IF excelHandler.createSpreadsheet(util.JSONArray.fromFGL(employees)) THEN
        # Step 5 - Get the output file
        DISPLAY SFMT("File created: %1", excelHandler.getFilename())
    END IF

END MAIN
```

### TSpreadsheet Method Reference

| Method | Description |
|--------|-------------|
| `init()` | Reset all properties to their initial state |
| `setHeaders(headers DYNAMIC ARRAY OF STRING)` | Set column header labels |
| `setRecordDefinition(parentNode om.DomNode)` | Set field names and types from a `base.TypeInfo.create()` node |
| `setTitle(title STRING)` | Set the sheet title |
| `createSpreadsheet(jsonArray util.JSONArray) RETURNS BOOLEAN` | Generate the Excel file; returns `TRUE` on success |
| `getFilename() RETURNS STRING` | Get the output file path (auto-generated temp file if not set) |

### Key Points

- The record definition is obtained from `base.TypeInfo.create(recordVariable)`. This introspects the variable's type metadata to determine field names and data types.
- Data is passed as a `util.JSONArray` using `util.JSONArray.fromFGL(yourArray)`.
- If you do not set a filename, one is automatically generated in the system temp directory.
- Column headers appear as bold, centered text in the first row.
- Columns are auto-sized to fit content.

---

## API 2: TSpreadsheetXtend (Grouped Data with Subtotals)

Use `TSpreadsheetXtend` when you need hierarchical grouping, subtotal formulas, subtitle rows, a "Report Group" column, or multi-sheet workbooks.

### Import

```4gl
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_xapi
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_helper
```

### Core Concepts

**Row Types** - Data is added row-by-row rather than as a bulk array. Each row is one of three types:

- **Data Row** - A normal row of field values
- **Group Header Row** - A merged row spanning all columns that marks the start of a group
- **Group Footer Row** - A row containing subtotal formulas for the preceding group

**Column Info** - Instead of simple headers, you provide a `TColumnInfo` array that defines both the header title and what aggregate calculation (if any) should appear in group footer rows.

**Groups** - Groups work as a stack. You push group headers onto the stack and pop them by adding footer rows. Groups can be nested to create multi-level subtotals.

### Aggregate Constants

Defined in `fgl_spreadsheet_helper`:

| Constant | Excel Formula | Description |
|----------|---------------|-------------|
| `cExcelSum` | `SUM(...)` | Sum of values |
| `cExcelAvg` | `AVG(...)` | Average of values |
| `cExcelCount` | `COUNTA(...)` | Count of non-empty cells |
| `cExcelMin` | `MIN(...)` | Minimum value |
| `cExcelMax` | `MAX(...)` | Maximum value |
| `cExcelNone` | *(none)* | No aggregate for this column |

### Example: Single Sheet with Groups

```4gl
IMPORT util
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_xapi
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_helper

TYPE TEarnings RECORD
    empl_num   CHAR(8),
    empl_name  VARCHAR(50),
    dept_code  CHAR(4),
    dept_desc  VARCHAR(50),
    gross_earn DECIMAL(12,2),
    net_earn   DECIMAL(12,2)
END RECORD

DEFINE earnings DYNAMIC ARRAY OF TEarnings
DEFINE earningRec TEarnings

MAIN
    DEFINE excelHandler fgl_spreadsheet_xapi.TSpreadsheetXtend
    DEFINE idx INTEGER
    DEFINE prevDept CHAR(4)

    # ... populate earnings array, sorted by dept_code ...

    # Step 1 - Initialize
    CALL excelHandler.init()

    # Step 2 - Set record definition and column info
    CALL excelHandler.setRecordDefinition(base.TypeInfo.create(earningRec))

    DEFINE colInfo DYNAMIC ARRAY OF fgl_spreadsheet_helper.TColumnInfo = [
        (colTitle: "Employee ID",   colCalc: fgl_spreadsheet_helper.cExcelCount),
        (colTitle: "Employee Name", colCalc: fgl_spreadsheet_helper.cExcelNone),
        (colTitle: "Dept Code",     colCalc: fgl_spreadsheet_helper.cExcelNone),
        (colTitle: "Dept Desc",     colCalc: fgl_spreadsheet_helper.cExcelNone),
        (colTitle: "Gross Earnings",colCalc: fgl_spreadsheet_helper.cExcelSum),
        (colTitle: "Net Earnings",  colCalc: fgl_spreadsheet_helper.cExcelSum)
    ]
    CALL excelHandler.setColumnInfo(colInfo)

    # Step 3 - Configure the spreadsheet
    CALL excelHandler.setTitle("Earnings Report")
    CALL excelHandler.setGroupColumn(TRUE)
    CALL excelHandler.addSubTitle("Quarterly Earnings Report")
    CALL excelHandler.addSubTitle("Q1 2026")

    # Step 4 - Add rows with group boundaries
    FOR idx = 1 TO earnings.getLength()

        # Close previous group when department changes
        IF prevDept IS NOT NULL AND prevDept != earnings[idx].dept_code THEN
            CALL excelHandler.addGroupFooterRow(prevDept)
        END IF

        # Open new group when department changes
        IF prevDept IS NULL OR prevDept != earnings[idx].dept_code THEN
            CALL excelHandler.addGroupHeaderRow(
                earnings[idx].dept_code,
                earnings[idx].dept_desc
            )
        END IF

        # Add data row
        CALL excelHandler.addDataRow(util.JSONObject.fromFGL(earnings[idx]))
        LET prevDept = earnings[idx].dept_code
    END FOR

    # Close the last group
    CALL excelHandler.addGroupFooterRow(prevDept)

    # Step 5 - Generate the spreadsheet
    IF excelHandler.createSpreadsheet() THEN
        DISPLAY SFMT("File created: %1", excelHandler.getFilename())
    END IF

END MAIN
```

### Understanding Group Nesting

Groups work as a stack. The order in which you push headers and pop footers determines the nesting. Here is an example with three levels of grouping at intervals of 20, 10, and 5 rows:

```4gl
FOR idx = 1 TO dataList.getLength()

    # Push group headers (outermost to innermost)
    IF idx MOD 20 == 1 THEN
        CALL excelHandler.addGroupHeaderRow("20", SFMT("Group %1 - %2", idx, idx+19))
    END IF
    IF idx MOD 10 == 1 THEN
        CALL excelHandler.addGroupHeaderRow("10", SFMT("Group %1 - %2", idx, idx+9))
    END IF
    IF idx MOD 5 == 1 THEN
        CALL excelHandler.addGroupHeaderRow("5", SFMT("Group %1 - %2", idx, idx+4))
    END IF

    # Add data row
    CALL excelHandler.addDataRow(util.JSONObject.fromFGL(dataList[idx]))

    # Pop group footers (innermost to outermost)
    IF idx MOD 5 == 0 THEN
        CALL excelHandler.addGroupFooterRow("5")
    END IF
    IF idx MOD 10 == 0 THEN
        CALL excelHandler.addGroupFooterRow("10")
    END IF
    IF idx MOD 20 == 0 THEN
        CALL excelHandler.addGroupFooterRow("20")
    END IF

END FOR
```

This produces Excel output with:
- A subtotal row after every 5 data rows
- A subtotal row after every 10 data rows (aggregating the two 5-row subtotals)
- A subtotal row after every 20 data rows (aggregating the two 10-row subtotals)
- A grand total row at the bottom of the sheet

**Important**: Group headers must be pushed from outermost to innermost, and group footers must be popped from innermost to outermost. The `group_id` string passed to `addGroupHeaderRow` and `addGroupFooterRow` is for your tracking purposes; the stack position determines the actual nesting.

### Example: Multi-Sheet Workbook

To create a workbook with multiple sheets, enable multi-sheet mode and call `initNewSheet()` before configuring each sheet. Call `createFile()` once at the end.

```4gl
DEFINE excelHandler fgl_spreadsheet_xapi.TSpreadsheetXtend

# Step 1 - Initialize and enable multi-sheet mode
CALL excelHandler.init()
CALL excelHandler.setMultiSheetMode(TRUE)

# Step 2 - Configure and create the first sheet
CALL excelHandler.initNewSheet()
CALL excelHandler.setRecordDefinition(base.TypeInfo.create(dataRec))
CALL excelHandler.setColumnInfo(sheet1ColumnInfo)
CALL excelHandler.setTitle("Sheet One")
CALL excelHandler.addSubTitle("First Sheet")
# ... add data rows, group headers/footers ...
CALL excelHandler.createSpreadsheet()

# Step 3 - Configure and create the second sheet
CALL excelHandler.initNewSheet()
CALL excelHandler.setRecordDefinition(base.TypeInfo.create(otherRec))
CALL excelHandler.setColumnInfo(sheet2ColumnInfo)
CALL excelHandler.setTitle("Sheet Two")
CALL excelHandler.addSubTitle("Second Sheet")
# ... add data rows, group headers/footers ...
CALL excelHandler.createSpreadsheet()

# Step 4 - Write the workbook to disk
CALL excelHandler.createFile()
DISPLAY SFMT("File created: %1", excelHandler.getFilename())
```

**Key points for multi-sheet mode:**
- Call `init()` once, then `setMultiSheetMode(TRUE)` before any sheet creation
- Use `initNewSheet()` (not `init()`) for each subsequent sheet -- this preserves the workbook reference while resetting all other properties
- Each sheet can have its own record definition, column info, title, subtitles, and grouping settings
- `createSpreadsheet()` renders each sheet but does not write to disk in multi-sheet mode
- Call `createFile()` once after all sheets have been created to write the workbook

### TSpreadsheetXtend Method Reference

| Method | Description |
|--------|-------------|
| `init()` | Reset all properties; creates a fresh state |
| `initNewSheet()` | Reset for a new sheet while preserving the workbook (for multi-sheet) |
| `setRecordDefinition(parentNode om.DomNode)` | Set field names and types from `base.TypeInfo.create()` |
| `setHeaders(headers DYNAMIC ARRAY OF STRING)` | Set simple column headers (use `setColumnInfo` instead for aggregates) |
| `setColumnInfo(colInfo DYNAMIC ARRAY OF TColumnInfo)` | Set column headers with aggregate formulas |
| `setTitle(title STRING)` | Set the sheet tab name |
| `setGroupColumn(groupCol BOOLEAN)` | Show/hide the "Report Group" column (default: `FALSE`) |
| `setDisplayGrandTotals(display BOOLEAN)` | Show/hide grand total row at bottom (default: `TRUE`) |
| `setMultiSheetMode(mode BOOLEAN)` | Enable multi-sheet workbook mode (default: `FALSE`) |
| `addSubTitle(title STRING)` | Add a merged subtitle row above the column headers |
| `addDataRow(rowData util.JSONObject)` | Add a data row |
| `addGroupHeaderRow(group_id STRING, group_title STRING)` | Push a group header onto the stack |
| `addGroupFooterRow(group_id STRING)` | Pop the innermost group and generate subtotal formulas |
| `createSpreadsheet() RETURNS BOOLEAN` | Render the sheet (and write to file unless in multi-sheet mode) |
| `createFile()` | Write the workbook to disk (multi-sheet mode only) |
| `getFilename() RETURNS STRING` | Get the output file path |

---

## API 3: tableExcelExport (UI Table Export)

Use `tableExcelExport` when you want a one-call export of a DISPLAY ARRAY or INPUT ARRAY table to Excel. It reads column metadata (titles, types, visibility, sort order, aggregate types) directly from the AUI tree, so no manual configuration is needed.

### Import

```4gl
IMPORT FGL com.fourjs.poiapi.fgl_table_export
```

### Usage

Call `tableExcelExport()` from within a DISPLAY ARRAY or INPUT ARRAY block, passing the screen record name and a JSON array of your data:

```4gl
IMPORT util
IMPORT FGL com.fourjs.poiapi.fgl_table_export

DEFINE dataList DYNAMIC ARRAY OF TMyRecord

OPEN WINDOW w WITH FORM "myform"

DISPLAY ARRAY dataList TO s_table.*

    ON ACTION export_to_excel ATTRIBUTES(TEXT="Export to Excel")
        VAR filename = fgl_table_export.tableExcelExport(
            "s_table",
            util.JSONArray.fromFGL(dataList)
        )
        IF filename.getLength() > 0 THEN
            CALL fgl_putfile(filename, "gbc")
        END IF

    ON ACTION CANCEL
        EXIT DISPLAY

END DISPLAY

CLOSE WINDOW w
```

### What It Does Automatically

- Reads column titles, data types, positions, and visibility from the form's AUI tree
- Skips `PhantomColumn` entries and hidden columns
- Respects front-end column reordering (uses `tabIndex` positions)
- Detects front-end sorting and re-sorts data in the export to match
- Reads aggregate type attributes (`SUM`, `AVG`, `COUNT`, `MIN`, `MAX`) from the form definition
- Creates a subtitle row using the window title
- Auto-sizes columns and applies data type formatting

### Adding Aggregates to the Form

To include aggregate totals in the export, define them on the form's TABLE widget. In your `.per` file, add `AGGREGATE` entries inside the `ATTRIBUTES` section of the TABLE:

```per
TABLE
{
 [f001  |f002          |f003         |f004          ]
 [f001  |f002          |f003         |f004          ]
}
ATTRIBUTES
    f001 = formonly.empName, TITLE = "Name";
    f002 = formonly.department, TITLE = "Department";
    f003 = formonly.grossEarnings, TITLE = "Gross";
    f004 = formonly.netEarnings, TITLE = "Net";
END
END

AGGREGATE
    tot001: SUM OF f003, TITLE = "Total:";
    tot002: SUM OF f004;
END
```

When `tableExcelExport` reads the AUI tree, it picks up the `aggregateType` attribute from each column and maps it to the appropriate Excel formula constant.

### Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `tableName` | `STRING` | The screen record name of the table (e.g., `"s_table"`) |
| `jsonData` | `util.JSONArray` | The data array serialized with `util.JSONArray.fromFGL()` |

**Returns**: `STRING` - The path to the generated Excel file, or an empty string on failure.

---

## Spreadsheet Output Features

All spreadsheets created by the library include these formatting features:

### Styling
- **Header row**: Bold text, centered alignment, bottom border
- **Data rows**: Formatted per data type (currency, decimal, date, etc.)
- **Group headers**: Bold, merged across all columns
- **Subtitle rows**: Bold, centered, merged across all columns
- **Group footers**: Contain Excel formulas (not static values)

### Page Setup (TSpreadsheetXtend only)
- Landscape orientation
- Grid lines printed
- Header row repeated on each printed page
- Page number in footer
- Frozen panes for subtitle and header rows

### Formula Generation

Group footer rows contain native Excel formulas rather than computed values. For example, a `SUM` aggregate on column C for rows 5-10 generates:

```
SUM(C5:C10)
```

When groups contain non-contiguous rows (due to nested sub-groups), the library generates multi-range formulas:

```
SUM(SUM(C5:C10),SUM(C15:C20))
```

For `COUNT` aggregates, the formula uses addition across ranges:

```
COUNTA(C5:C10)+COUNTA(C15:C20)
```

---

## Helper Types Reference

The `fgl_spreadsheet_helper` module defines shared types used across the API:

```4gl
PUBLIC TYPE TFields RECORD
    fieldName STRING,     # Field name in the record
    fieldType STRING      # Genero data type (e.g., "INTEGER", "MONEY(10,2)")
END RECORD

PUBLIC TYPE TColumnInfo RECORD
    colTitle STRING,      # Excel column header text
    colCalc STRING        # Aggregate formula constant (cExcelSum, cExcelNone, etc.)
END RECORD

PUBLIC TYPE TDataRow RECORD
    rowType STRING,       # "DATA", "GROUPHEADER", or "GROUPFOOTER"
    rowData util.JSONObject  # Row data as JSON
END RECORD
```

---

## Package Structure

```
com.fourjs.poiapi/
├── fgl_excel.4gl              # Low-level Apache POI Java wrapper
├── fgl_structures.4gl         # Row stack and range utilities
├── fgl_spreadsheet_helper.4gl # Data type formatting and constants
├── fgl_spreadsheet_interface.4gl  # ISpreadsheet interface definition
├── fgl_spreadsheet_api.4gl    # TSpreadsheet (simple export)
├── fgl_spreadsheet_xapi.4gl   # TSpreadsheetXtend (grouped export)
└── fgl_table_export.4gl       # UI Table export function
```

### Module Responsibilities

- **fgl_excel** - Direct Java interop layer. Creates workbooks, sheets, rows, cells, styles, and fonts via `IMPORT JAVA`. You generally do not call this module directly.
- **fgl_structures** - Manages the group row stack used by `TSpreadsheetXtend` to track which data rows belong to which group, enabling correct formula ranges.
- **fgl_spreadsheet_helper** - Defines shared types (`TFields`, `TColumnInfo`, `TDataRow`), aggregate constants, and utility functions for date/time conversion and cell style creation.
- **fgl_spreadsheet_api** - The `TSpreadsheet` type for simple, flat exports.
- **fgl_spreadsheet_xapi** - The `TSpreadsheetXtend` type for grouped exports with subtotals and multi-sheet support.
- **fgl_table_export** - The `tableExcelExport()` function that bridges the UI table AUI tree to the `TSpreadsheetXtend` API.

---

## Quick Reference: Choosing an API

| Scenario | API |
|----------|-----|
| Export a flat list of records to Excel | `TSpreadsheet` |
| Export grouped data with subtotals | `TSpreadsheetXtend` |
| Export with multiple sheets in one workbook | `TSpreadsheetXtend` with `setMultiSheetMode(TRUE)` |
| Export a UI table widget as-is | `tableExcelExport()` |
| Export a UI table with aggregate totals | `tableExcelExport()` with aggregates defined in the `.per` form |
