PACKAGE com.fourjs.poiapi
IMPORT util

IMPORT FGL com.fourjs.poiapi.fgl_excel
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_helper
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_api
IMPORT FGL com.fourjs.poiapi.fgl_structures

PUBLIC TYPE TSpreadsheetXtend RECORD
   spreadsheet TSpreadsheet,
   groupCol BOOLEAN,
   cellOffset INTEGER,
   displayGrandTotals BOOLEAN,
   colInfo DYNAMIC ARRAY OF TColumnInfo,
   dataRows DYNAMIC ARRAY OF TDataRow,
   subTitles DYNAMIC ARRAY OF STRING,
   multiSheetMode BOOLEAN
END RECORD

PRIVATE DEFINE cellStyleDict DICTIONARY OF fgl_excel.cellStyleType
PRIVATE DEFINE headerStyleDict DICTIONARY OF fgl_excel.cellStyleType
PRIVATE DEFINE footerStyleDict DICTIONARY OF fgl_excel.cellStyleType
PRIVATE DEFINE formulaStyleDict DICTIONARY OF fgl_excel.cellStyleType

PRIVATE DEFINE calcRowStack TRowStack

PUBLIC FUNCTION (self TSpreadsheetXtend) init() RETURNS ()

    CALL self.spreadsheet.init()
    CALL self.colInfo.clear()
    CALL self.dataRows.clear()
    LET self.groupCol = FALSE
    LET self.cellOffset = 1
    LET self.displayGrandTotals = TRUE
    CALL self.subTitles.clear()
    LET self.multiSheetMode = FALSE

END FUNCTION #init

PUBLIC FUNCTION (self TSpreadsheetXtend) initNewSheet() RETURNS ()
    DEFINE currentWorkbook fgl_excel.workbookType

    LET currentWorkbook = self.spreadsheet.workbook
    
    CALL self.spreadsheet.init()
    CALL self.colInfo.clear()
    CALL self.dataRows.clear()
    LET self.groupCol = FALSE
    LET self.cellOffset = 1
    LET self.displayGrandTotals = TRUE
    CALL self.subTitles.clear()

    LET self.spreadsheet.workbook = currentWorkbook

END FUNCTION #ini

PUBLIC FUNCTION (self TSpreadsheetXtend) getFilename() RETURNS (STRING)

   RETURN self.spreadsheet.getFilename()

END FUNCTION

PUBLIC FUNCTION (self TSpreadsheetXtend) setRecordDefinition(parentNode om.DomNode)
	RETURNS ()

   CALL self.spreadsheet.setRecordDefinition(parentNode)

END FUNCTION #setRecordDefinition

PUBLIC FUNCTION (self TSpreadsheetXtend) getRecordDefinition() RETURNS (DYNAMIC ARRAY OF TFields)

	RETURN self.spreadsheet.getRecordDefinition()

END FUNCTION #getRecordDefinition

PUBLIC FUNCTION (self TSpreadsheetXtend) setHeaders(headers DYNAMIC ARRAY OF STRING)
	RETURNS ()

   CALL self.spreadsheet.setHeaders(headers)

END FUNCTION #setHeaders

PUBLIC FUNCTION (self TSpreadsheetXtend) getHeaders()
	RETURNS (DYNAMIC ARRAY OF STRING)

	RETURN self.spreadsheet.getHeaders()

END FUNCTION #setHeaders

PUBLIC FUNCTION (self TSpreadsheetXtend) setTitle(title STRING) RETURNS ()

   CALL self.spreadsheet.setTitle(title)

END FUNCTION #setTitle

PUBLIC FUNCTION (self TSpreadsheetXtend) getTitle() RETURNS (STRING)

	RETURN self.spreadsheet.getTitle()

END FUNCTION #setTitle

PUBLIC FUNCTION (self TSpreadsheetXtend) setGroupColumn(groupCol BOOLEAN) RETURNS ()

   LET self.groupCol = groupCol
   LET self.cellOffset = IIF (groupCol, 0, 1)

END FUNCTION #setGroupColumn

PUBLIC FUNCTION (self TSpreadsheetXtend) getGroupColumn() RETURNS (BOOLEAN)

   RETURN self.groupCol

END FUNCTION #getGroupColumn

PUBLIC FUNCTION (self TSpreadsheetXtend) setDisplayGrandTotals(displayGrandTotals BOOLEAN) RETURNS ()

   LET self.displayGrandTotals = displayGrandTotals

END FUNCTION #setDisplayGrandTotals

PUBLIC FUNCTION (self TSpreadsheetXtend) getDisplayGrandTotals() RETURNS (BOOLEAN)

   RETURN self.displayGrandTotals

END FUNCTION #getDisplayGrandTotals

PUBLIC FUNCTION (self TSpreadsheetXtend) setMultiSheetMode(multiSheetMode BOOLEAN) RETURNS ()

   LET self.multiSheetMode = multiSheetMode

END FUNCTION #setMultiSheetMode

PUBLIC FUNCTION (self TSpreadsheetXtend) getMultiSheetMode() RETURNS (BOOLEAN)

   RETURN self.multiSheetMode

END FUNCTION #getMultiSheetMode

PUBLIC FUNCTION (self TSpreadsheetXtend) addSubTitle(title STRING) RETURNS ()
    DEFINE idx INTEGER

    CALL self.subTitles.appendElement()
    LET idx = self.subTitles.getLength()
    LET self.subTitles[idx] = title

END FUNCTION #addSubTitle

PUBLIC FUNCTION (self TSpreadsheetXtend) addDataRow(rowData util.JSONObject) RETURNS ()
   DEFINE idx INTEGER

   LET idx = self.dataRows.getLength() + 1
   LET self.dataRows[idx].rowType = cDataRowType
   LET self.dataRows[idx].rowData = rowData

END FUNCTION #addDataRow

PUBLIC FUNCTION (self TSpreadsheetXtend) addGroupHeaderRow(group_id STRING, group_title STRING) RETURNS ()
   DEFINE idx INTEGER
   DEFINE r_header THeaderRow

   LET r_header.group_id = group_id
   LET r_header.group_title = group_title

   LET idx = self.dataRows.getLength() + 1
   LET self.dataRows[idx].rowType = cGroupHeaderRowType
   LET self.dataRows[idx].rowData = util.JSONObject.fromFGL(r_header)

END FUNCTION #addGroupHeaderRow

PUBLIC FUNCTION (self TSpreadsheetXtend) addGroupFooterRow(group_id STRING) RETURNS ()
   DEFINE idx INTEGER
   DEFINE r_header THeaderRow

   LET r_header.group_id = group_id

   LET idx = self.dataRows.getLength() + 1
   LET self.dataRows[idx].rowType = cGroupFooterRowType
   LET self.dataRows[idx].rowData = util.JSONObject.fromFGL(r_header)

END FUNCTION #addGroupHeaderRow

PUBLIC FUNCTION (self TSpreadsheetXtend) getCellIndex(fieldName STRING) RETURNS INTEGER
   DEFINE idx INTEGER

   FOR idx = 1 TO self.spreadsheet.fields.getLength()
      IF self.spreadsheet.fields[idx].fieldName == fieldName THEN
         IF self.groupCol THEN
            RETURN idx
         ELSE
         RETURN (idx - 1)
         END IF
      END IF
   END FOR

   RETURN -1

END FUNCTION #getCellIndex

PUBLIC FUNCTION (self TSpreadsheetXtend) setColumnInfo(colInfo DYNAMIC ARRAY OF TColumnInfo) RETURNS ()

   CALL self.colInfo.clear()
   CALL colInfo.copyTo(self.colInfo)

END FUNCTION #setColumnInfo

PUBLIC FUNCTION (self TSpreadsheetXtend) getColumnInfo() RETURNS (DYNAMIC ARRAY OF TColumnInfo)
   DEFINE colInfo DYNAMIC ARRAY OF TColumnInfo

   CALL self.colInfo.copyTo(colInfo)
   RETURN colInfo

END FUNCTION #getColumnInfo

PUBLIC FUNCTION (self TSpreadsheetXtend) createSpreadsheet() RETURNS BOOLEAN
    DEFINE excelRow fgl_excel.rowType
    DEFINE excelCell fgl_excel.cellType
    DEFINE headerStyle fgl_excel.cellStyleType
    DEFINE headerFont fgl_excel.fontType
    DEFINE idx INTEGER
    DEFINE rowIdx INTEGER = 0
    DEFINE hasGroups BOOLEAN = FALSE
    DEFINE lastRow INTEGER = 0
    DEFINE firstTime BOOLEAN = TRUE
    CONSTANT cReportGroup = "Report Group"

    LET self.cellOffset = IIF (self.groupCol, 0, 1)
	TRY
		#Initialize Workbook and Spreadsheet
        IF self.spreadsheet.workbook IS NULL THEN
            LET self.spreadsheet.workbook = fgl_excel.workbook_create()
        ELSE
            LET firstTime = FALSE
        END IF
        LET self.spreadsheet.sheet = workbook_createsheet_with_name(self.spreadsheet.workbook, self.getTitle())

		#Create header style
        CALL fgl_excel.style_create(self.spreadsheet.workbook) RETURNING headerStyle
		CALL fgl_excel.font_create(self.spreadsheet.workbook) RETURNING headerFont
		CALL fgl_excel.font_set(headerFont, "weight", "bold")

		CALL fgl_excel.style_set(headerStyle, "alignment","center")
        CALL fgl_excel.style_font_set(headerStyle, headerFont)

        CALL fgl_excel.set_border_style(headerStyle, FALSE, FALSE, TRUE, FALSE)

        IF self.subTitles.getLength() > 0 THEN
            CALL self.createSubtitleRows(rowIdx)
            LET rowIdx = self.subTitles.getLength()
        END IF

        #Add colinfo first
        IF self.colInfo.getLength() > 0 THEN

            #Use the colinfo array to build the column headers
            LET excelRow = fgl_excel.sheet_createrow(self.spreadsheet.sheet, rowIdx)
            LET rowIdx = rowIdx + 1
            IF self.groupCol THEN
                LET excelCell = fgl_excel.row_createcell(excelRow, 0)
                CALL fgl_excel.cell_value_set(excelCell, cReportGroup)
                CALL fgl_excel.cell_style_set(excelCell, headerStyle)
            END IF

            FOR idx = 1 TO self.colInfo.getLength()
                LET excelCell = fgl_excel.row_createcell(excelRow, idx - self.cellOffset)
                CALL fgl_excel.cell_value_set(excelCell, self.colInfo[idx].colTitle)
                CALL fgl_excel.cell_style_set(excelCell, headerStyle)
            END FOR

        ELSE

            #Add column headers
            LET excelRow = fgl_excel.sheet_createrow(self.spreadsheet.sheet, rowIdx)
            LET rowIdx = rowIdx + 1
            FOR idx = 1 TO self.spreadsheet.headers.getLength()
                LET excelCell = fgl_excel.row_createcell(excelRow, idx - self.cellOffset)
                CALL fgl_excel.cell_value_set(excelCell, self.spreadsheet.headers[idx])
                CALL fgl_excel.cell_style_set(excelCell, headerStyle)
            END FOR

        END IF

        #Initialize module variables
        IF firstTime THEN
            CALL cellStyleDict.clear()
            CALL headerStyleDict.clear()
            CALL footerStyleDict.clear()
            CALL formulaStyleDict.clear()
        END IF
        CALL calcRowStack.init()

        #Now loop through the data
        FOR idx = 1 TO self.dataRows.getLength()

            #Create a new Excel row
            LET excelRow = fgl_excel.sheet_createrow(self.spreadsheet.sheet, rowIdx)

            CASE self.dataRows[idx].rowType

                WHEN cDataRowType #Typical Data Row
                   CALL self.createDataRow(idx, rowIdx, excelRow)

                WHEN cGroupHeaderRowType
                   CALL self.createGroupHeaderRow(idx, rowIdx, excelRow)
                   LET hasGroups = TRUE

                WHEN cGroupFooterRowType
                   CALL self.createGroupFooterRow(excelRow)

            END CASE
            LET rowIdx = rowIdx + 1

        END FOR
        LET lastRow = IIF(self.displayGrandTotals, 0, 1)
        WHILE calcRowStack.currentLevel() > lastRow
            #Create a new Excel row
            LET excelRow = fgl_excel.sheet_createrow(self.spreadsheet.sheet, rowIdx)
            CALL self.createGroupFooterRow(excelRow)
            LET rowIdx = rowIdx + 1
        END WHILE

        #Autosize the columns
        CALL self.autoSizeColumns()

        #Freeze the subtitle and header rows
        CALL fgl_excel.freeze_rows(self.spreadsheet.sheet, 0, self.subTitles.getLength() + 1)

        #PageSetup
        CALL fgl_excel.set_header_footer(self.spreadsheet.sheet, 0.3, 0.3)
        CALL fgl_excel.set_margins(self.spreadsheet.sheet, 0.75, 0.75, 0.7, 0.7)
       { CALL fgl_excel.set_sheet_header(
            self.spreadsheet.sheet,
            NULL,
            self.getTitle(),
            util.Datetime.format(
                CURRENT YEAR TO SECOND,
                "%d/%m/%Y %I:%M:%S %p"
            )
        )}
		#4JS Year month date for left side and format changed on 07112022...
		CALL fgl_excel.set_sheet_header(
            self.spreadsheet.sheet,
            util.Datetime.format(
                CURRENT YEAR TO SECOND,
                "%B %d, %Y"
            ),
            self.getTitle(),
		    NULL 
        )
        CALL fgl_excel.set_print_settings(self.spreadsheet.sheet)

        IF NOT self.multiSheetMode THEN
            #Write to File
            CALL self.createFile()
        END IF

	CATCH

		RETURN FALSE

	END TRY

	RETURN TRUE

END FUNCTION #createSpreadsheet

PUBLIC FUNCTION (self TSpreadsheetXtend) createDataRow(dataIdx INTEGER, rowIdx INTEGER, excelRow fgl_excel.rowType) RETURNS ()
   DEFINE idx INTEGER
   DEFINE cellField TFields
   DEFINE excelCell fgl_excel.cellType

   FOR idx = 1 TO self.spreadsheet.fields.getLength()
      LET cellField = self.spreadsheet.fields[idx]
      IF self.dataRows[dataIdx].rowData.has(cellField.fieldName) THEN
         LET excelCell = fgl_excel.row_createcell(excelRow, idx - self.cellOffset)
         CALL setDataCell(
            self.spreadsheet.workbook,
            excelCell,
            cellField.fieldType,
            self.dataRows[dataIdx].rowData.get(cellField.fieldName)
         )
      END IF
   END FOR

   CALL calcRowStack.addRow(rowIdx)

END FUNCTION #createDataRow

PUBLIC FUNCTION (self TSpreadsheetXtend) createGroupHeaderRow(dataIdx INTEGER, rowIdx INTEGER, excelRow fgl_excel.rowType) RETURNS ()
   DEFINE lastCol INTEGER
   DEFINE idx INTEGER
   DEFINE excelCell fgl_excel.cellType
   DEFINE r_header THeaderRow
   DEFINE headerStyle fgl_excel.cellStyleType
   DEFINE headerFont fgl_excel.fontType
	CONSTANT cHeaderIdx = "XAPI_HEADER"

   CALL self.dataRows[dataIdx].rowData.toFGL(r_header)
   CALL calcRowStack.pushGroup(r_header.group_title)

   IF headerStyleDict.contains(cHeaderIdx) THEN
       LET headerStyle = headerStyleDict[cHeaderIdx]
   ELSE
       #Create group header style
       CALL fgl_excel.font_create(self.spreadsheet.workbook) RETURNING headerFont
       CALL fgl_excel.font_set(headerFont, "weight", "bold")

       CALL fgl_excel.style_create(self.spreadsheet.workbook) RETURNING headerStyle
       CALL fgl_excel.style_set(headerStyle, "alignment","justify")
       CALL fgl_excel.style_font_set(headerStyle, headerFont)
       CALL fgl_excel.set_border_style(headerStyle, FALSE, FALSE, TRUE, FALSE)
       LET  headerStyleDict[cHeaderIdx] = headerStyle
   END IF

   IF self.groupCol THEN
      LET excelCell = fgl_excel.row_createcell(excelRow, 0)
      CALL fgl_excel.cell_value_set(excelCell, r_header.group_title)
      CALL fgl_excel.cell_style_set(excelCell, headerStyle)
   END IF

   FOR idx = 1 TO self.colInfo.getLength()
      LET excelCell = fgl_excel.row_createcell(excelRow, idx - self.cellOffset)
      CALL fgl_excel.cell_style_set(excelCell, headerStyle)
      IF idx == 1 AND NOT self.groupCol THEN
         CALL fgl_excel.cell_value_set(excelCell, r_header.group_title)
      ELSE
         CALL fgl_excel.cell_value_set(excelCell, NULL)
      END IF
   END FOR
   LET lastCol = self.colInfo.getLength() - self.cellOffset
   CALL fgl_excel.merge_cells(self.spreadsheet.sheet, rowIdx, 0, lastCol)

END FUNCTION #createGroupHeaderRow

PUBLIC FUNCTION (self TSpreadsheetXtend) createGroupFooterRow(excelRow fgl_excel.rowType) RETURNS ()
   DEFINE idx INTEGER
   DEFINE excelCell fgl_excel.cellType
   DEFINE formulaText STRING
   DEFINE footerStyle fgl_excel.cellStyleType
   DEFINE groupIdx INTEGER
   DEFINE groupRows TIntArray
   DEFINE calcRows TStringArray
   DEFINE joinString STRING
   DEFINE offset INTEGER = 0
   DEFINE footerTitle STRING

   CALL calcRowStack.popGroup()
      RETURNING footerTitle, groupRows

   IF self.groupCol THEN
      LET excelCell = fgl_excel.row_createcell(excelRow, 0)
      CALL fgl_excel.cell_value_set(excelCell, footerTitle)
      CALL fgl_excel.cell_style_set(excelCell, footerStyle)
      LET offset = 1
   END IF

   FOR idx = 1 TO self.colInfo.getLength()
      LET formulaText = NULL
      IF self.colInfo[idx].colCalc == cExcelNone THEN
         LET excelCell = fgl_excel.row_createcell(excelRow, (idx - self.cellOffset))
         CALL fgl_excel.cell_value_set(excelCell, "")
         CALL fgl_excel.cell_style_set(excelCell, footerStyle)
      ELSE
         LET calcRows = fgl_structures.getFormattedStrings(column2Letter(idx + offset), groupRows)
         LET excelCell = fgl_excel.row_createcell(excelRow, (idx - self.cellOffset))
         IF calcRows.getLength() == 1 THEN
            LET formulaText = SFMT(
               "%1%2",
               self.colInfo[idx].colCalc,
               calcRows[1]
            )
         ELSE
            LET joinString = IIF(self.colInfo[idx].colCalc == cExcelCount, "+", ",")
            FOR groupIdx = 1 TO calcRows.getLength()
               IF groupIdx == 1 THEN
                  IF self.colInfo[idx].colCalc != cExcelCount THEN
                     LET formulaText = SFMT("%1(",
                        self.colInfo[idx].colCalc
                     )
                  END IF
               ELSE
                  LET formulaText = SFMT("%1%2", formulaText, joinString)
               END IF
               LET formulaText = SFMT(
                  "%1%2%3",
                  formulaText,
                  self.colInfo[idx].colCalc,
                  calcRows[groupIdx]
               )
               IF groupIdx == calcRows.getLength() AND self.colInfo[idx].colCalc != cExcelCount THEN
                  LET formulaText = SFMT("%1)", formulaText)
               END IF
            END FOR
         END IF
         CALL setFormulaCell(
            self.spreadsheet.workbook,
            excelCell,
            self.spreadsheet.fields[idx].fieldType,
            formulaText,
            footerStyle
         )
      END IF
   END FOR

END FUNCTION #createGroupFooterRow

PUBLIC FUNCTION (self TSpreadsheetXtend) autoSizeColumns() RETURNS()
   DEFINE idx INTEGER
   DEFINE len INTEGER

   LET len = IIF(self.groupCol, 1, 0) + self.colInfo.getLength()

   FOR idx = 1 TO len
      CALL fgl_excel.auto_size_column(self.spreadsheet.sheet, idx - 1)
   END FOR

END FUNCTION #autoSizeColumns

PUBLIC FUNCTION (self TSpreadsheetXtend) createSubtitleRows(startIdx INTEGER) RETURNS ()
   DEFINE lastCol INTEGER
   DEFINE colIdx INTEGER
   DEFINE rowIdx INTEGER
   DEFINE excelRow fgl_excel.rowType
   DEFINE excelCell fgl_excel.cellType
   DEFINE titleStyle fgl_excel.cellStyleType
   DEFINE titleFont fgl_excel.fontType
   DEFINE title STRING
   
    CALL fgl_excel.font_create(self.spreadsheet.workbook) RETURNING titleFont
    CALL fgl_excel.font_set(titleFont, "weight", "bold")

    CALL fgl_excel.style_create(self.spreadsheet.workbook) RETURNING titleStyle
    CALL fgl_excel.style_set(titleStyle, "alignment","center")
    CALL fgl_excel.style_font_set(titleStyle, titleFont)

    IF self.colInfo.getLength() > 0 THEN
        LET lastCol = self.colInfo.getLength() - self.cellOffset
    ELSE
        LET lastCol = self.getHeaders().getLength() - self.cellOffset
    END IF

    FOR rowIdx = 1 TO self.subTitles.getLength()
        LET excelRow = fgl_excel.sheet_createrow(self.spreadsheet.sheet, rowIdx - 1 + startIdx)
        LET title = self.subTitles[rowIdx]
        IF self.groupCol THEN
            LET excelCell = fgl_excel.row_createcell(excelRow, 0)
            CALL fgl_excel.cell_value_set(excelCell, title)
            CALL fgl_excel.cell_style_set(excelCell, titleStyle)
        END IF
        
        FOR colIdx = 1 TO IIF(self.colInfo.getLength() > 0, self.colInfo.getLength(), self.spreadsheet.headers.getLength())
            LET excelCell = fgl_excel.row_createcell(excelRow, colIdx - self.cellOffset)
            CALL fgl_excel.cell_style_set(excelCell, titleStyle)
            IF colIdx == 1 AND NOT self.groupCol THEN
                CALL fgl_excel.cell_value_set(excelCell, title)
            ELSE
                CALL fgl_excel.cell_value_set(excelCell, NULL)
            END IF
        END FOR
        CALL fgl_excel.merge_cells(self.spreadsheet.sheet, rowIdx - 1 + startIdx, 0, lastCol)
    END FOR

END FUNCTION #createSubtitleRows

PUBLIC FUNCTION (self TSpreadsheetXtend) createFile() RETURNS ()

    #Write to File
    CALL fgl_excel.workbook_writeToFile(self.spreadsheet.workbook, self.getFilename())

END FUNCTION #createFile

PRIVATE FUNCTION setDataCell(workbook fgl_excel.workbookType, excelCell fgl_excel.cellType, fieldType STRING, cellValue STRING) RETURNS ()
   DEFINE cellStyle fgl_excel.cellStyleType
   DEFINE dtYearToSecond DATETIME YEAR TO SECOND
   DEFINE dtHourToSecond DATETIME HOUR TO SECOND
    
   IF cellStyleDict.contains(fieldType) THEN
      LET cellStyle = cellStyleDict[fieldType]
   ELSE
      LET cellStyle = getCellStyleForDataType(workbook, fieldType)
      LET cellStyleDict[fieldType] = cellStyle
   END IF

   #set the cell style and value
   CASE
      WHEN fieldType MATCHES "DEC*"
         #set the field data and the style of the cell
         CALL fgl_excel.cell_number_set(excelCell, NVL(cellValue, 0))
         CALL fgl_excel.cell_style_set(excelCell, cellStyle)

      WHEN fieldType MATCHES "*INT*"
         #set the field data and the style of the cell
         CALL fgl_excel.cell_number_set(excelCell, NVL(cellValue, 0))
         CALL fgl_excel.cell_style_set(excelCell, cellStyle)

      WHEN fieldType MATCHES "*MONEY*"
         #set the field data and the style of the cell
         CALL fgl_excel.cell_number_set(excelCell, NVL(cellValue, 0))
         CALL fgl_excel.cell_style_set(excelCell, cellStyle)

      WHEN fieldType MATCHES "*FLOAT*"
         #set the field data and the style of the cell
         CALL fgl_excel.cell_number_set(excelCell, NVL(cellValue, 0))
         CALL fgl_excel.cell_style_set(excelCell, cellStyle)

      WHEN fieldType == "DATE"
         #set the field data and the style of the cell
			IF cellValue IS NOT NULL THEN
				VAR dateValue = dateConverter(cellValue)
				LET cellValue = dateValue USING "yyyy-mm-dd"
			END IF
         CALL fgl_excel.cell_date_set(excelCell, cellValue)
         CALL fgl_excel.cell_style_set(excelCell, cellStyle)

      WHEN fieldType MATCHES "DATETIME YEAR*"
         #set the field data and the style of the cell
         LET dtYearToSecond = datetimeConverter(cellValue)
         CALL fgl_excel.cell_datetime_set(excelCell, dtYearToSecond)
         CALL fgl_excel.cell_style_set(excelCell, cellStyle)

      WHEN fieldType MATCHES "DATETIME HOUR*"
         #set the field data and the style of the cell
         LET dtHourToSecond = timeConverter(cellValue)
         CALL fgl_excel.cell_time_set(excelCell, dtHourToSecond)
         CALL fgl_excel.cell_style_set(excelCell, cellStyle)

      OTHERWISE
         #No formatting for string, varchar, or char data types
         CALL fgl_excel.cell_value_set(excelCell, cellValue)

   END CASE

END FUNCTION

PRIVATE FUNCTION setFormulaCell(workbook fgl_excel.workbookType, 
                                excelCell fgl_excel.cellType,
                                fieldType STRING, 
                                cellValue STRING,
                                cellStyle fgl_excel.cellStyleType)
										  RETURNS ()

   DEFINE clonedStyle fgl_excel.cellStyleType
   DEFINE styleKey STRING

   IF cellValue MATCHES "COUNT[(]" THEN
      LET fieldType = "INTEGER"
   END IF

   LET styleKey = fieldType

   IF formulaStyleDict.contains(styleKey) THEN
      LET clonedStyle = formulaStyleDict[styleKey]
   ELSE
      LET clonedStyle = fgl_excel.create_style_from_style(workbook, cellStyle)
      CALL setCellStyleForDataType(workbook, clonedStyle, fieldType)
      LET formulaStyleDict[styleKey] = clonedStyle
   END IF

   #set the cell style and value
   CASE
      WHEN fieldType MATCHES "DEC*"
         #set the field data and the style of the cell
         CALL fgl_excel.cell_formula_set(excelCell, cellValue)
         CALL fgl_excel.cell_style_set(excelCell, clonedStyle)

      WHEN fieldType MATCHES "*INT*"
         #set the field data and the style of the cell
         CALL fgl_excel.cell_formula_set(excelCell, cellValue)
         CALL fgl_excel.cell_style_set(excelCell, clonedStyle)

      WHEN fieldType MATCHES "*MONEY*"
         #set the field data and the style of the cell
         CALL fgl_excel.cell_formula_set(excelCell, cellValue)
         CALL fgl_excel.cell_style_set(excelCell, clonedStyle)

      WHEN fieldType MATCHES "*FLOAT*"
         #set the field data and the style of the cell
         CALL fgl_excel.cell_formula_set(excelCell, cellValue)
         CALL fgl_excel.cell_style_set(excelCell, clonedStyle)

      OTHERWISE
         #set the field data and the style of the cell
         CALL fgl_excel.cell_formula_set(excelCell, cellValue)
         CALL fgl_excel.cell_style_set(excelCell, clonedStyle)

   END CASE

END FUNCTION #setFormulaCell

PRIVATE FUNCTION column2Letter(cellIdx INTEGER) RETURNS STRING
   DEFINE cellId STRING
   CONSTANT asciiOffset = 64
   CONSTANT alphaLen = 26

   IF cellIdx > alphaLen THEN
      LET cellId = ASCII((cellIdx / alphaLen) + asciiOffset),
                   ASCII((cellIdx MOD alphaLen) + asciiOffset)
   ELSE
      LET cellId = ASCII((cellIdx+asciiOffset))
   END IF

   RETURN cellId

END FUNCTION #column2Letter


