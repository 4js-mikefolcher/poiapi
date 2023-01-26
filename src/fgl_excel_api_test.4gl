IMPORT util
IMPORT os
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_api
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_xapi
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_helper

TYPE TData RECORD
	stringField STRING,
    numericString STRING,
	moneyField MONEY(12,2),
	decimalField DECIMAL(8,4),
	integerField INTEGER,
	smallintField SMALLINT,
	dateField DATE,
	datetimeField DATETIME YEAR TO SECOND,
    datetimeField2 DATETIME YEAR TO MINUTE,
    datetimeField3 DATETIME HOUR TO SECOND,
    datetimeField4 DATETIME HOUR TO MINUTE,
	charField CHAR(35),
	varcharField VARCHAR(100),
	floatField FLOAT,
	smallfloatField SMALLFLOAT,
	booleanField BOOLEAN
END RECORD
DEFINE dataList DYNAMIC ARRAY OF TData
DEFINE dataRec TData

PRIVATE TYPE TEarnings RECORD
    empl_num  CHAR(8),
    empl_name VARCHAR(50),
    dept_code CHAR(4),
    dept_desc VARCHAR(50),
    gross_earn DECIMAL(12,2),
    net_earn   DECIMAL(12,2)
END RECORD

MAIN
	DEFINE idx INTEGER
	DEFINE excelHandler fgl_spreadsheet_api.TSpreadsheet

	CALL STARTLOG("fgl_excel_api_test.log")
    IF os.Path.separator() == "/" THEN
        RUN "env > env.txt"
    ELSE
        RUN "set > env.txt"
    END IF

	FOR idx = 1 TO 100
		LET dataRec.stringField = SFMT("Row #%1", idx)
        LET dataRec.numericString = idx USING "&&&&&&&&&"
		LET dataRec.charField = SFMT("%1 x %1 y %1 z %1", idx)
		LET dataRec.varcharField = SFMT("'%1' '%1' '%1' '%1'", ASCII(idx + 32))
		LET dataRec.booleanField = idx MOD 2
		LET dataRec.dateField = TODAY - idx UNITS DAY
		LET dataRec.datetimeField = CURRENT YEAR TO SECOND
        LET dataRec.datetimeField2 = CURRENT YEAR TO MINUTE
        LET dataRec.datetimeField3 = CURRENT HOUR TO SECOND
        LET dataRec.datetimeField4 = CURRENT HOUR TO MINUTE
		IF dataRec.booleanField THEN
			#odd
			LET dataRec.smallintField = -1 * idx
			LET dataRec.integerField = -2 * idx
			LET dataRec.floatField = -7.17 * idx
			LET dataRec.smallfloatField = -1.1 * idx
			LET dataRec.moneyField = -11.69 * idx
			LET dataRec.decimalField = -13.3954 * idx
		ELSE
			#even
			LET dataRec.smallintField = idx
			LET dataRec.integerField = 2 * idx
			LET dataRec.floatField = 7.153 * idx
			LET dataRec.smallfloatField = 1.1 * idx
			LET dataRec.moneyField = 11.9 * idx
			LET dataRec.decimalField = 13.3954 * idx
		END IF
        IF idx MOD 11 == 0 THEN
            LET dataRec.dateField = NULL
            LET dataRec.datetimeField = NULL
            LET dataRec.datetimeField2 = NULL
            LET dataRec.datetimeField3 = NULL
            LET dataRec.datetimeField4 = NULL
        END IF
		LET dataList[idx] = dataRec
	END FOR

   CALL excelAPIExample()
   IF int_flag THEN
       RETURN
   END IF
   
   CALL excelXAPIExample()
   IF int_flag THEN
       RETURN
   END IF

   CALL excelMultisheetExample()

END MAIN

PRIVATE FUNCTION excelAPIExample() RETURNS ()
	DEFINE excelHandler fgl_spreadsheet_api.TSpreadsheet

	#Steps to create an excel spreadsheet from a array of records
	CALL excelHandler.init()
	CALL excelHandler.setHeaders(excelHeader())
	CALL excelHandler.setRecordDefinition(base.TypeInfo.create(dataRec))
	CALL excelHandler.setTitle("Test Spreadsheet API")
	IF excelHandler.createSpreadsheet(util.JSONArray.fromFGL(dataList)) THEN
		DISPLAY SFMT("Excel file path: %1", excelHandler.getFilename())
		IF arg_val(1) == "web" THEN
			MENU "API 1 Example"
                COMMAND "Download"
                    CALL fgl_putfile(excelHandler.getFilename(), "gbc")
                COMMAND "Next"
                    EXIT MENU
                COMMAND "Exit"
                    LET int_flag = TRUE
                    EXIT MENU
			END MENU
		END IF
	END IF

END FUNCTION

PRIVATE FUNCTION excelHeader() RETURNS DYNAMIC ARRAY OF STRING
	DEFINE headerList DYNAMIC ARRAY OF STRING = [
		"String",
        "Numeric String",
		"Money",
		"Decimal",
		"Integer",
		"Small Integer",
		"Date",
		"Datetime Year to Second",
        "Datetime Year to Minute",
        "Datetime Hour to Second",
        "Datetime Hour to Minute",
		"Char",
		"Varchar",
		"Float",
		"Small Float",
		"Boolean"
	]

	RETURN headerList

END FUNCTION

PRIVATE FUNCTION excelXAPIExample() RETURNS ()
	DEFINE excelHandler fgl_spreadsheet_xapi.TSpreadsheetXtend
    DEFINE idx INTEGER

	#Steps to create an excel spreadsheet from a array of records
	CALL excelHandler.init()
	CALL excelHandler.setColumnInfo(columnInfoArray())
	CALL excelHandler.setRecordDefinition(base.TypeInfo.create(dataRec))
	CALL excelHandler.setTitle("Test Spreadsheet XAPI")
    CALL excelHandler.setGroupColumn(TRUE)
    CALL excelHandler.addSubTitle("Test Spreadsheet XAPI")
    CALL excelHandler.addSubTitle("This is a test of the sub title stuff")
    #CALL excelHandler.setDisplayGrandTotals(FALSE)
    FOR idx = 1 TO dataList.getLength()

      #Add the headers
      IF idx MOD 20 == 1 THEN
         CALL excelHandler.addGroupHeaderRow("20", SFMT("Group %1 - %2", idx, (idx-1+20)))
      END IF
      IF idx MOD 10 == 1 THEN
         CALL excelHandler.addGroupHeaderRow("10", SFMT("Group %1 - %2", idx, (idx-1+10)))
      END IF
      IF idx MOD 5 == 1 THEN
         CALL excelHandler.addGroupHeaderRow("5", SFMT("Group %1 - %2", idx, (idx-1+5)))
      END IF

      #Add the actual data
      CALL excelHandler.addDataRow(util.JSONObject.fromFGL(dataList[idx]))

      #Add the footers (subtotals)
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

	IF excelHandler.createSpreadsheet() THEN
		DISPLAY SFMT("Excel file path: %1", excelHandler.getFilename())
		IF arg_val(1) == "web" THEN
			MENU "API 2 Example"
                COMMAND "Download"
                   CALL fgl_putfile(excelHandler.getFilename(), "gbc")
                COMMAND "Next"
                    EXIT MENU
                COMMAND "Exit"
                    LET int_flag = TRUE
                    EXIT MENU
			END MENU
		END IF
	END IF

END FUNCTION

PRIVATE FUNCTION excelMultisheetExample() RETURNS ()
	DEFINE excelHandler fgl_spreadsheet_xapi.TSpreadsheetXtend
    DEFINE sheetCount INTEGER

	#Steps to create an excel spreadsheet from a array of records
	CALL excelHandler.init()
    CALL excelHandler.setMultisheetMode(TRUE)

    FOR sheetCount = 1 TO 4
        CASE sheetCount
            WHEN 1
                CALL createSheetOne(excelHandler, sheetCount)
            WHEN 2
                CALL createSheetTwo(excelHandler, sheetCount)
            WHEN 3
                CALL createSheetOne(excelHandler, sheetCount)
            WHEN 4
                CALL createSheetTwo(excelHandler, sheetCount)
        END CASE

        IF NOT excelHandler.createSpreadsheet() THEN
            RETURN
        END IF
    END FOR

    CALL excelHandler.createFile()

    DISPLAY SFMT("Excel file path: %1", excelHandler.getFilename())
    IF arg_val(1) == "web" THEN
        MENU "Multisheet Example"
            COMMAND "Download"
               CALL fgl_putfile(excelHandler.getFilename(), "gbc")
            COMMAND "Exit"
                LET int_flag = TRUE
                EXIT MENU
        END MENU
    END IF

END FUNCTION

PRIVATE FUNCTION createSheetOne(excelHandler fgl_spreadsheet_xapi.TSpreadsheetXtend INOUT, sheetIdx INTEGER) RETURNS ()
    DEFINE idx INTEGER

    CALL excelHandler.initNewSheet()
	CALL excelHandler.setColumnInfo(columnInfoArray())
	CALL excelHandler.setRecordDefinition(base.TypeInfo.create(dataRec))
	CALL excelHandler.setTitle(SFMT("Multisheet - Sheet %1", sheetIdx))
    CALL excelHandler.addSubTitle("Multiple Sheet Example XAPI")
    CALL excelHandler.addSubTitle(SFMT("This is sheet #%1", sheetIdx))

    IF sheetIdx > 2 THEN
        CALL excelHandler.setDisplayGrandTotals(FALSE)
        CALL excelHandler.setGroupColumn(FALSE)
    ELSE
        CALL excelHandler.setDisplayGrandTotals(TRUE)
        CALL excelHandler.setGroupColumn(TRUE)
    END IF

    FOR idx = 1 TO dataList.getLength()

        #Add the headers
        IF idx MOD 20 == 1 THEN
            CALL excelHandler.addGroupHeaderRow("20", SFMT("Group %1 - %2", idx, (idx-1+20)))
        END IF
        IF idx MOD 10 == 1 THEN
            CALL excelHandler.addGroupHeaderRow("10", SFMT("Group %1 - %2", idx, (idx-1+10)))
        END IF
        IF idx MOD 5 == 1 THEN
            CALL excelHandler.addGroupHeaderRow("5", SFMT("Group %1 - %2", idx, (idx-1+5)))
        END IF

        #Add the actual data
        CALL excelHandler.addDataRow(util.JSONObject.fromFGL(dataList[idx]))

        #Add the footers (subtotals)
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

END FUNCTION

PRIVATE FUNCTION createSheetTwo(excelHandler fgl_spreadsheet_xapi.TSpreadsheetXtend INOUT, sheetIdx INTEGER) RETURNS ()
    DEFINE idx INTEGER
    DEFINE earning TEarnings
    DEFINE earnings DYNAMIC ARRAY OF TEarnings
    DEFINE dept_code CHAR(4)

    CALL excelHandler.initNewSheet()
	CALL excelHandler.setColumnInfo(earningsInfoArray())
	CALL excelHandler.setRecordDefinition(base.TypeInfo.create(earning))
	CALL excelHandler.setTitle(SFMT("Multisheet - Sheet %1", sheetIdx))
    CALL excelHandler.addSubTitle("Multiple Sheet Example XAPI")
    CALL excelHandler.addSubTitle(SFMT("This is sheet #%1", sheetIdx))

    IF sheetIdx > 2 THEN
        CALL excelHandler.setDisplayGrandTotals(FALSE)
        CALL excelHandler.setGroupColumn(FALSE)
    ELSE
        CALL excelHandler.setDisplayGrandTotals(TRUE)
        CALL excelHandler.setGroupColumn(TRUE)
    END IF

    LET earnings = fillEarningsInfo()
    FOR idx = 1 TO earnings.getLength()

        #Add the footers (subtotals)
        IF dept_code IS NOT NULL AND dept_code != earnings[idx].dept_code THEN
            CALL excelHandler.addGroupFooterRow(dept_code)
        END IF

        IF dept_code IS NULL OR dept_code != earnings[idx].dept_code THEN
            CALL excelHandler.addGroupHeaderRow(earnings[idx].dept_code, earnings[idx].dept_desc)
        END IF

        #Add the actual data
        CALL excelHandler.addDataRow(util.JSONObject.fromFGL(earnings[idx]))

        LET dept_code = earnings[idx].dept_code
        
    END FOR

END FUNCTION

PRIVATE FUNCTION columnInfoArray() RETURNS (DYNAMIC ARRAY OF TColumnInfo)
   DEFINE colInfoArray DYNAMIC ARRAY OF TColumnInfo = [
		(colTitle: "String", colCalc: fgl_spreadsheet_helper.cExcelNone),
        (colTitle: "Numeric String", colCalc: fgl_spreadsheet_helper.cExcelCount),
		(colTitle: "Money", colCalc: fgl_spreadsheet_helper.cExcelSum),
		(colTitle: "Decimal", colCalc: fgl_spreadsheet_helper.cExcelSum),
		(colTitle: "Integer", colCalc: fgl_spreadsheet_helper.cExcelSum),
		(colTitle: "Small Integer", colCalc: fgl_spreadsheet_helper.cExcelSum),
		(colTitle: "Date", colCalc: fgl_spreadsheet_helper.cExcelNone),
		(colTitle: "Datetime Year to Second", colCalc: fgl_spreadsheet_helper.cExcelNone),
        (colTitle: "Datetime Year to Minute", colCalc: fgl_spreadsheet_helper.cExcelNone),
        (colTitle: "Datetime Hour to Second", colCalc: fgl_spreadsheet_helper.cExcelNone),
        (colTitle: "Datetime Hour to Minute", colCalc: fgl_spreadsheet_helper.cExcelNone),
		(colTitle: "Char", colCalc: fgl_spreadsheet_helper.cExcelNone),
		(colTitle: "Varchar", colCalc: fgl_spreadsheet_helper.cExcelNone),
		(colTitle: "Float", colCalc: fgl_spreadsheet_helper.cExcelSum),
		(colTitle: "Small Float", colCalc: fgl_spreadsheet_helper.cExcelSum),
		(colTitle: "Boolean", colCalc: fgl_spreadsheet_helper.cExcelNone)
	]

   RETURN colInfoArray

END FUNCTION #columnInfoArray

PRIVATE FUNCTION earningsInfoArray() RETURNS (DYNAMIC ARRAY OF TColumnInfo)
   DEFINE colInfoArray DYNAMIC ARRAY OF TColumnInfo = [
		(colTitle: "Employee ID", colCalc: fgl_spreadsheet_helper.cExcelCount),
        (colTitle: "Employee Name", colCalc: fgl_spreadsheet_helper.cExcelNone),
		(colTitle: "Department Code", colCalc: fgl_spreadsheet_helper.cExcelNone),
		(colTitle: "Department Description", colCalc: fgl_spreadsheet_helper.cExcelNone),
		(colTitle: "Gross Earnings", colCalc: fgl_spreadsheet_helper.cExcelSum),
		(colTitle: "Net Earnings", colCalc: fgl_spreadsheet_helper.cExcelSum)
	]

   RETURN colInfoArray

END FUNCTION #earningsInfoArray

PRIVATE FUNCTION fillEarningsInfo() RETURNS (DYNAMIC ARRAY OF TEarnings)
    DEFINE earningsInfo DYNAMIC ARRAY OF TEarnings
    DEFINE idx INTEGER

    CALL earningsInfo.clear()
    FOR idx = 1 TO 33
        LET earningsInfo[idx].empl_num = SFMT("010086%1", idx USING "&&")
        LET earningsInfo[idx].empl_name = SFMT("%1 %2", getFirstName(idx), getLastName(idx))
        CASE
            WHEN idx < 6
                LET earningsInfo[idx].dept_code = "IT"
                LET earningsInfo[idx].dept_desc = "Internal IT" 
            WHEN idx < 19
                LET earningsInfo[idx].dept_code = "HR"
                LET earningsInfo[idx].dept_desc = "Human Resources"
            WHEN idx < 27
                LET earningsInfo[idx].dept_code = "MT"
                LET earningsInfo[idx].dept_desc = "Maintenance"
            OTHERWISE
                LET earningsInfo[idx].dept_code = "RD"
                LET earningsInfo[idx].dept_desc = "Research and Development"
        END CASE
        LET earningsInfo[idx].gross_earn = 1317 * idx - (17 / idx) * 3
        LET earningsInfo[idx].net_earn = earningsInfo[idx].gross_earn * 0.63
    END FOR

    RETURN earningsInfo

END FUNCTION #fillEarningsInfo

PRIVATE FUNCTION getFirstName(idx INTEGER) RETURNS (STRING)
    DEFINE nameIdx INTEGER
    DEFINE firstNameList DYNAMIC ARRAY OF STRING = [
        "Mark",
        "Jennifer",
        "Raaj",
        "Sally",
        "Fred",
        "Henry",
        "Becky",
        "Roger",
        "Laura"
    ]
    LET nameIdx = (idx MOD firstNameList.getLength()) + 1
    RETURN firstNameList[nameIdx]

END FUNCTION

PRIVATE FUNCTION getLastName(idx INTEGER) RETURNS (STRING)
    DEFINE nameIdx INTEGER
    DEFINE lastNameList DYNAMIC ARRAY OF STRING = [
        "Smith",
        "Jones",
        "Patil",
        "Richardson",
        "Adams",
        "Miller",
        "Orville",
        "Fletcher",
        "Zimmerman",
        "Young",
        "Potter"
    ]
    LET nameIdx = (idx MOD lastNameList.getLength()) + 1
    RETURN lastNameList[nameIdx]

END FUNCTION