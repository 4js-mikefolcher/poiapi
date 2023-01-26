PACKAGE com.fourjs.poiapi

IMPORT util
IMPORT os

IMPORT FGL com.fourjs.poiapi.fgl_excel
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_helper

PUBLIC TYPE TSpreadsheet RECORD
	filename STRING,
	workbook fgl_excel.workbookType,
	sheet fgl_excel.sheetType,
	fields DYNAMIC ARRAY OF TFields,
	title STRING,
	headers DYNAMIC ARRAY OF STRING
END RECORD

PRIVATE DEFINE cellStyleDict DICTIONARY OF fgl_excel.cellStyleType

PUBLIC FUNCTION (self TSpreadsheet) init() RETURNS ()

	LET self.filename = NULL
	LET self.workbook = NULL
	LET self.sheet = NULL
	CALL self.fields.clear()
	LET self.title = NULL
	CALL self.headers.clear()

END FUNCTION #init

PUBLIC FUNCTION (self TSpreadsheet) getFilename() RETURNS (STRING)

	IF self.filename IS NULL OR self.filename.getLength() == 0 THEN
		LET self.filename = os.Path.makeTempName(), ".xlsx"
	END IF
	RETURN self.filename

END FUNCTION

PUBLIC FUNCTION (self TSpreadsheet) setRecordDefinition(parentNode om.DomNode)
	RETURNS ()
	DEFINE nodeList om.NodeList
	DEFINE idx INTEGER
	DEFINE fieldNode om.DomNode

	CALL self.fields.clear()
	LET nodeList = parentNode.selectByTagName("Field")
	FOR idx = 1 TO nodeList.getLength()
		LET fieldNode = nodeList.item(idx)
		LET self.fields[idx].fieldName = fieldNode.getAttribute("name")
		LET self.fields[idx].fieldType = fieldNode.getAttribute("type")
	END FOR

END FUNCTION #setRecordDefinition

PUBLIC FUNCTION (self TSpreadsheet) getRecordDefinition() RETURNS (DYNAMIC ARRAY OF TFields)

	RETURN self.fields

END FUNCTION #getRecordDefinition

PUBLIC FUNCTION (self TSpreadsheet) setHeaders(headers DYNAMIC ARRAY OF STRING)
	RETURNS ()

	CALL self.headers.clear()
	CALL headers.copyTo(self.headers)

END FUNCTION #setHeaders

PUBLIC FUNCTION (self TSpreadsheet) getHeaders()
	RETURNS (DYNAMIC ARRAY OF STRING)

	RETURN self.headers

END FUNCTION #setHeaders

PUBLIC FUNCTION (self TSpreadsheet) setTitle(title STRING) RETURNS ()

	LET self.title = title

END FUNCTION #setTitle

PUBLIC FUNCTION (self TSpreadsheet) getTitle() RETURNS (STRING)

	RETURN self.title

END FUNCTION #setTitle

PUBLIC FUNCTION (self TSpreadsheet) createSpreadsheet(jsonArray util.JSONArray) RETURNS BOOLEAN
	DEFINE excelRow fgl_excel.rowType
	DEFINE excelCell fgl_excel.cellType
	DEFINE headerStyle fgl_excel.cellStyleType
	DEFINE headerFont fgl_excel.fontType
	DEFINE idx, idx2 INTEGER
	DEFINE jsonObj util.JSONObject
	DEFINE cellField TFields
	DEFINE cellStyle fgl_excel.cellStyleType
    DEFINE dtYearToSecond DATETIME YEAR TO SECOND
    DEFINE dtHourToSecond DATETIME HOUR TO SECOND

	TRY
		#Initialize Workbook and Spreadsheet
		LET self.workbook = fgl_excel.workbook_create()
		LET self.sheet = fgl_excel.workbook_createsheet(self.workbook)

		#Create header style
		CALL fgl_excel.font_create(self.workbook) RETURNING headerFont
		CALL fgl_excel.font_set(headerFont, "weight", "bold")

		CALL fgl_excel.style_create(self.workbook) RETURNING headerStyle
		CALL fgl_excel.style_set(headerStyle, "alignment","center")
		CALL fgl_excel.style_font_set(headerStyle, headerFont)
   
		#Add column headers
		LET excelRow = fgl_excel.sheet_createrow(self.sheet, 0)
		FOR idx = 1 TO self.headers.getLength()
			LET excelCell = fgl_excel.row_createcell(excelRow, idx-1)
			CALL fgl_excel.cell_value_set(excelCell, self.headers[idx])
			CALL fgl_excel.cell_style_set(excelCell, headerStyle)
		END FOR

		CALL cellStyleDict.clear()
		FOR idx = 1 TO jsonArray.getLength()
			LET jsonObj = jsonArray.get(idx)
			LET excelRow = fgl_excel.sheet_createrow(self.sheet, idx)
			FOR idx2 = 1 TO self.fields.getLength()
				LET cellField = self.fields[idx2]
				IF jsonObj.has(cellField.fieldName) THEN

					#create the excel row
					LET excelCell = fgl_excel.row_createcell(excelRow, idx2-1)

					#get the cached style
					IF cellStyleDict.contains(cellField.fieldType) THEN
						LET cellStyle = cellStyleDict[cellField.fieldType]
					ELSE
						LET cellStyle = NULL
					END IF

					#set the cell style and value
					CASE
						WHEN cellField.fieldType MATCHES "DEC*"
							#Handle Decimal Field Formatting and value
							IF cellStyle IS NULL THEN
								#get the decimal string format and build a cell style from it
								LET cellStyle = fgl_excel.cell_style_builtin_create(self.workbook, fgl_excel.cDecimalFormat)
								LET cellStyleDict[cellField.fieldType] = cellStyle
							END IF
							#set the field data and the style of the cell
							CALL fgl_excel.cell_number_set(excelCell, jsonObj.get(cellField.fieldName))
							CALL fgl_excel.cell_style_set(excelCell, cellStyle)

						WHEN cellField.fieldType MATCHES "*INT*"
							#Handle Integer Field Formatting and value
							IF cellStyle IS NULL THEN
								#set integer format
								LET cellStyle = fgl_excel.cell_style_builtin_create(self.workbook, fgl_excel.cIntegerFormat)
								LET cellStyleDict[cellField.fieldType] = cellStyle
							END IF
							CALL fgl_excel.cell_number_set(excelCell, jsonObj.get(cellField.fieldName))
							CALL fgl_excel.cell_style_set(excelCell, cellStyle)

						WHEN cellField.fieldType MATCHES "*MONEY*"
							#Handle Money Field Formatting and value
							IF cellStyle IS NULL THEN
								#set money format and value
								LET cellStyle = fgl_excel.cell_style_builtin_create(self.workbook, fgl_excel.cMoneyFormat)
								LET cellStyleDict[cellField.fieldType] = cellStyle
							END IF
							CALL fgl_excel.cell_number_set(excelCell, jsonObj.get(cellField.fieldName))
							CALL fgl_excel.cell_style_set(excelCell, cellStyle)

						WHEN cellField.fieldType MATCHES "*FLOAT*"
							#Handle floating point Field Formatting and value
							IF cellStyle IS NULL THEN
								#get the float string format and build a cell style from it
								LET cellStyle = fgl_excel.cell_style_builtin_create(self.workbook, fgl_excel.cDecimalFormat)
								LET cellStyleDict[cellField.fieldType] = cellStyle
							END IF
							#set the field data and the style of the cell
							CALL fgl_excel.cell_number_set(excelCell, jsonObj.get(cellField.fieldName))
							CALL fgl_excel.cell_style_set(excelCell, cellStyle)

						WHEN cellField.fieldType == "DATE"
							#Handle Date Field Formatting and value
							IF cellStyle IS NULL THEN
								#get the date string format and build a cell style from it
								LET cellStyle = fgl_excel.cell_style_builtin_create(self.workbook, fgl_excel.cDateFormat)
								LET cellStyleDict[cellField.fieldType] = cellStyle
							END IF
							#set the field data and the style of the cell
							CALL fgl_excel.cell_date_set(excelCell, jsonObj.get(cellField.fieldName))
							CALL fgl_excel.cell_style_set(excelCell, cellStyle)

						WHEN cellField.fieldType MATCHES "DATETIME YEAR*"
							#Handle Datetime Field Formatting and value
							IF cellStyle IS NULL THEN
								#get the datetime string format and build a cell style from it
								LET cellStyle = fgl_excel.cell_style_builtin_create(self.workbook, fgl_excel.cDatetimeFormat)
								LET cellStyleDict[cellField.fieldType] = cellStyle
							END IF
							#set the field data and the style of the cell
                            LET dtYearToSecond = datetimeConverter(jsonObj.get(cellField.fieldName))
							CALL fgl_excel.cell_datetime_set(excelCell, dtYearToSecond)
							CALL fgl_excel.cell_style_set(excelCell, cellStyle)

						WHEN cellField.fieldType MATCHES "DATETIME HOUR*"
							#Handle Datetime Field Formatting and value
							IF cellStyle IS NULL THEN
								#get the datetime string format and build a cell style from it
								LET cellStyle = fgl_excel.cell_style_builtin_create(self.workbook, fgl_excel.cTimeFormat)
								LET cellStyleDict[cellField.fieldType] = cellStyle
							END IF
							#set the field data and the style of the cell
                            LET dtHourToSecond = timeConverter(jsonObj.get(cellField.fieldName))
							CALL fgl_excel.cell_time_set(excelCell, dtHourToSecond)
							CALL fgl_excel.cell_style_set(excelCell, cellStyle)

						OTHERWISE
							#No formatting for string, varchar, or char data types
							CALL fgl_excel.cell_value_set(excelCell, jsonObj.get(cellField.fieldName))

					END CASE

				END IF
			END FOR
		END FOR

      #Autosize the columns
      CALL self.autoSizeColumns()

		#Write to File
		CALL fgl_excel.workbook_writeToFile(self.workbook, self.getFilename());

	CATCH

		RETURN FALSE

	END TRY

	RETURN TRUE

END FUNCTION #createSpreadsheet

PUBLIC FUNCTION (self TSpreadsheet) autoSizeColumns() RETURNS()
   DEFINE idx INTEGER

   FOR idx = 1 TO self.fields.getLength()
      CALL fgl_excel.auto_size_column(self.sheet, idx - 1)
   END FOR

END FUNCTION #autoSizeColumns