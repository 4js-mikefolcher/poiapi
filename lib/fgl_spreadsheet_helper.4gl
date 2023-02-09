PACKAGE com.fourjs.poiapi
IMPORT util
IMPORT FGL com.fourjs.poiapi.fgl_excel

PUBLIC TYPE TFields RECORD
	fieldName   STRING,
	fieldType   STRING
END RECORD

PUBLIC CONSTANT cExcelSum = "SUM"
PUBLIC CONSTANT cExcelSubTotal = "SUBTOTAL"
PUBLIC CONSTANT cExcelAvg = "AVG"
PUBLIC CONSTANT cExcelNone = "NONE"
PUBLIC CONSTANT cExcelCount = "COUNTA"
PUBLIC CONSTANT cExcelMin = "MIN"
PUBLIC CONSTANT cExcelMax = "MAX"

PUBLIC TYPE TColumnInfo RECORD
   colTitle STRING,
   colCalc STRING
END RECORD

PUBLIC CONSTANT cDataRowType = "DATA"
PUBLIC CONSTANT cGroupHeaderRowType = "GROUPHEADER"
PUBLIC CONSTANT cGroupFooterRowType = "GROUPFOOTER"

PUBLIC TYPE TDataRow RECORD
   rowType STRING,
   rowData util.JSONObject
END RECORD

PUBLIC TYPE THeaderRow RECORD
   group_id STRING,
   group_title STRING
END RECORD

PUBLIC FUNCTION getCellStyleForDataType(workbook fgl_excel.workbookType, fglDataType STRING) RETURNS fgl_excel.cellStyleType
   DEFINE cellStyle fgl_excel.cellStyleType

   LET cellStyle = fgl_excel.style_create(workbook)
   CALL setCellStyleForDataType(workbook, cellStyle, fglDataType)
   RETURN cellStyle

END FUNCTION #getCellStyleForDataType

PUBLIC FUNCTION setCellStyleForDataType(workbook fgl_excel.workbookType,
                                        cellStyle fgl_excel.cellStyleType,
                                        fglDataType STRING) RETURNS ()
    CONSTANT cDatetimeFormat = "mm/dd/yyyy hh:mm:ss AM/PM"
    CONSTANT cDatetimeFormat2 = "hh:mm:ss AM/PM"

    #builds the cell style
    CASE
      WHEN fglDataType MATCHES "DEC*"
         #set decimal style
         CALL fgl_excel.set_style_format(workbook, cellStyle, createDecFormat(fglDataType))
      WHEN fglDataType MATCHES "*INT*"
         #set integer format
         CALL fgl_excel.set_builtin_style_format(cellStyle, fgl_excel.cIntegerFormat)
      WHEN fglDataType MATCHES "*MONEY*"
         #set money format and value
         CALL fgl_excel.set_builtin_style_format(cellStyle, fgl_excel.cMoneyFormat)
      WHEN fglDataType MATCHES "*FLOAT*"
         #Handle floating point Field Formatting
         CALL fgl_excel.set_builtin_style_format(cellStyle, fgl_excel.cDecimalFormat)
      WHEN fglDataType == "DATE"
         #get the date string format and build a cell style from it
         CALL fgl_excel.set_builtin_style_format(cellStyle, fgl_excel.cDateFormat)
      WHEN fglDataType MATCHES "DATETIME YEAR*"
         #get the datetime string format and build a cell style from it
         CALL fgl_excel.set_style_format(workbook, cellStyle, cDatetimeFormat)
      WHEN fglDataType MATCHES "DATETIME HOUR*"
         #get the datetime string format and build a cell style from it
         CALL fgl_excel.set_style_format(workbook, cellStyle, cDatetimeFormat2)

    END CASE

END FUNCTION #setCellStyleForDataType

PRIVATE FUNCTION createDecFormat(fglDataType STRING) RETURNS STRING
	DEFINE startPos     INTEGER
	DEFINE endPos       INTEGER
	DEFINE decPrecision INTEGER
	DEFINE decScale     INTEGER
	DEFINE idx          INTEGER
	DEFINE numStr       STRING
	DEFINE currChr      CHAR(1)

	LET startPos = fglDataType.getIndexOf("(", 1)
	LET endPos = fglDataType.getIndexOf(")", 1)

	FOR idx = startPos + 1 TO fglDataType.getLength()
		LET currChr = fglDataType.getCharAt(idx)
		IF currChr MATCHES "[0-9]" THEN
			LET numStr = numStr.append(currChr)
		ELSE
			EXIT FOR
		END IF
	END FOR

	LET decPrecision = numStr
	LET numStr = ""

	FOR idx = endPos - 1 TO 1 STEP -1
		LET currChr = fglDataType.getCharAt(idx)
		IF currChr MATCHES "[0-9]" THEN
			LET numStr = currChr, numStr.trim()
		ELSE
			EXIT FOR
		END IF
	END FOR

	LET decScale = numStr
	LET numStr = ""

	IF decPrecision > 3 THEN
		LET numStr = "#,##0"
	ELSE
		LET numStr = "##0"
	END IF

	IF decScale > 0 THEN
		LET numStr = numStr.append(".")
		FOR idx = 1 TO decScale
			LET numStr = numStr.append("0")
		END FOR
	END IF

	#Build a format string of the form "#,##0.00;[Red](#,##0.00)"
	LET numStr = SFMT("%1;[Red](%1)", numStr)
	RETURN numStr

END FUNCTION #createDecFormat

PUBLIC FUNCTION datetimeConverter(str STRING) RETURNS DATETIME YEAR TO SECOND
   DEFINE conValue DATETIME YEAR TO SECOND
   DEFINE formatList DYNAMIC ARRAY OF STRING = [
      "%Y-%m-%d %T",
      "%Y-%m-%d %R",
      "%Y-%m-%d"
   ]
   DEFINE idx INTEGER

   INITIALIZE conValue TO NULL
   IF str IS NULL THEN
      RETURN conValue
   END IF

   FOR idx = 1 TO formatList.getLength()
      LET conValue = util.Datetime.parse(str, formatList[idx])
      IF conValue IS NOT NULL THEN
         RETURN conValue
      END IF
   END FOR
   
   RETURN conValue

END FUNCTION #datetimeConverter

PUBLIC FUNCTION timeConverter(str STRING) RETURNS DATETIME HOUR TO SECOND
   DEFINE conValue DATETIME YEAR TO SECOND
   DEFINE formatList DYNAMIC ARRAY OF STRING = [
      "%T",
      "%R"
   ]
   DEFINE idx INTEGER

   INITIALIZE conValue TO NULL
   IF str IS NULL THEN
      RETURN conValue
   END IF

   FOR idx = 1 TO formatList.getLength()
      LET conValue = util.Datetime.parse(str, formatList[idx])
      IF conValue IS NOT NULL THEN
         RETURN conValue
      END IF
   END FOR

   RETURN conValue

END FUNCTION #timeConverter

PUBLIC FUNCTION dateConverter(dateValue STRING) RETURNS (DATE)
   DEFINE dateType DATE
   DEFINE idx INTEGER
   DEFINE singleChar CHAR(1)
   CONSTANT cSlash = "/"
   CONSTANT cDash = "-"

	IF dateValue IS NULL THEN
		RETURN NULL
	END IF

   VAR charFound = FALSE
   #Determine the format of the string
   FOR idx = 1 TO dateValue.getLength()
      LET singleChar = dateValue.getCharAt(idx)
      IF singleChar == cSlash OR singleChar == cDash THEN
         LET charFound = TRUE
         EXIT FOR
      END IF
   END FOR

   IF charFound THEN
      CASE
         WHEN idx == 5 AND singleChar == cDash
            #Assume the yyyy-mm-dd format 
            LET dateType = util.Date.parse(dateValue, "yyyy-mm-dd")
         WHEN idx == 3 and singleChar == cDash
            IF isUSADateFormat() THEN
               #Assume the mm-dd-yyyy format 
               LET dateType = util.Date.parse(dateValue, "mm-dd-yyyy")
            ELSE
               #Assume the dd-mm-yyyy format
               LET dateType = util.Date.parse(dateValue, "dd-mm-yyyy")
            END IF
         WHEN idx == 3 and singleChar == cSlash
            IF isUSADateFormat() THEN
               #Assume the mm/dd/yyyy format 
               LET dateType = util.Date.parse(dateValue, "mm/dd/yyyy")
            ELSE
               #Assume the dd/mm/yyyy format 
               LET dateType = util.Date.parse(dateValue, "dd/mm/yyyy")
            END IF
      END CASE
   ELSE
      RETURN DATE(dateValue)
   END IF

   RETURN dateType

END FUNCTION #dateConverter

PRIVATE DEFINE dateFormatUSA SMALLINT = -1
PRIVATE FUNCTION isUSADateFormat() RETURNS (BOOLEAN)

   IF dateFormatUSA > -1 THEN
      RETURN (dateFormatUSA == 1)
   END IF

   VAR dbDateValue = FGL_GETENV("DBDATE")
   IF dbDateValue IS NULL OR dbDateValue.getLength() == 0 THEN
      #Assume the USA date format if DBDATE is not set
      LET dateFormatUSA = 1
   ELSE
      IF dbDateValue MATCHES "MD*" THEN
         LET dateFormatUSA = 1
      ELSE
         LET dateFormatUSA = 0
      END IF
   END IF

   RETURN (dateFormatUSA == 1)

END FUNCTION #isUSADateFormat



