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


#------------------------------------------------------------------------------
# Cell format selection
#
# Date, time and monetary cell formats are built from the format settings of
# the machine the Genero application runs on - DBDATE for dates, DBFORMAT or
# DBMONEY for the currency symbol - so the spreadsheet reads the same way the
# application does, on every machine that opens it.
#
# Excel resolves a built-in format id (14 "m/d/yy", 8 "$#,##0.00", ...) against
# the regional settings of the machine *viewing* the file, so a built-in cannot
# honour the application's settings. Every format below is therefore written
# into the workbook as an explicit format code.
#------------------------------------------------------------------------------

#Formats follow DBDATE / DBFORMAT / DBMONEY. This is the default.
PUBLIC CONSTANT cFormatModeLocale = "LOCALE"
#Formats are left to the regional settings of the machine viewing the file.
PUBLIC CONSTANT cFormatModeViewer = "VIEWER"
#Formats are ISO 8601 dates, 24 hour times and unadorned numbers.
PUBLIC CONSTANT cFormatModeISO = "ISO"

PRIVATE DEFINE formatMode STRING = cFormatModeLocale

#Bumped whenever the format settings change, so that the callers holding a
#cache of cell styles can tell that their cached styles are stale.
PRIVATE DEFINE formatGeneration INTEGER = 1

#Explicit format codes, set by the caller. These win over the mode.
PRIVATE DEFINE ovDateFormat STRING
PRIVATE DEFINE ovDatetimeFormat STRING
PRIVATE DEFINE ovTimeFormat STRING
PRIVATE DEFINE ovMoneyFormat STRING
PRIVATE DEFINE ovCurrencySymbol STRING

#Cached derivations of the environment, cleared when the settings change
PRIVATE DEFINE envDateCode STRING
PRIVATE DEFINE envCurrencySymbol STRING
PRIVATE DEFINE envCurrencyIsFront BOOLEAN
PRIVATE DEFINE envCurrencyRead BOOLEAN

#Selects how date, time and monetary cells are formatted. Pass one of
#cFormatModeLocale (the default), cFormatModeViewer or cFormatModeISO.
#Call before building a spreadsheet.
PUBLIC FUNCTION setFormatMode(mode STRING) RETURNS ()

   CASE mode
      WHEN cFormatModeLocale
         LET formatMode = mode
      WHEN cFormatModeViewer
         LET formatMode = mode
      WHEN cFormatModeISO
         LET formatMode = mode
      OTHERWISE
         #An unknown mode would silently format everything as text, so refuse it
         RETURN
   END CASE

   CALL invalidateFormats()

END FUNCTION #setFormatMode

PUBLIC FUNCTION getFormatMode() RETURNS STRING
   RETURN formatMode
END FUNCTION #getFormatMode

#Overrides the format code used for DATE cells, for example "yyyy-mm-dd".
#Pass NULL to go back to the format implied by the mode.
PUBLIC FUNCTION setDateFormat(formatCode STRING) RETURNS ()
   LET ovDateFormat = formatCode
   CALL invalidateFormats()
END FUNCTION #setDateFormat

#Overrides the format code used for DATETIME YEAR TO ... cells.
PUBLIC FUNCTION setDatetimeFormat(formatCode STRING) RETURNS ()
   LET ovDatetimeFormat = formatCode
   CALL invalidateFormats()
END FUNCTION #setDatetimeFormat

#Overrides the format code used for DATETIME HOUR TO ... cells.
PUBLIC FUNCTION setTimeFormat(formatCode STRING) RETURNS ()
   LET ovTimeFormat = formatCode
   CALL invalidateFormats()
END FUNCTION #setTimeFormat

#Overrides the format code used for MONEY cells. The code is used as given, so
#it carries its own currency symbol and its own number of decimal places.
PUBLIC FUNCTION setMoneyFormat(formatCode STRING) RETURNS ()
   LET ovMoneyFormat = formatCode
   CALL invalidateFormats()
END FUNCTION #setMoneyFormat

#Overrides only the currency symbol of MONEY cells, keeping the number of
#decimal places that the column type implies. The symbol leads the value.
PUBLIC FUNCTION setCurrencySymbol(symbol STRING) RETURNS ()
   LET ovCurrencySymbol = symbol
   CALL invalidateFormats()
END FUNCTION #setCurrencySymbol

#Drops every explicit format code and goes back to the current mode.
PUBLIC FUNCTION clearFormatOverrides() RETURNS ()
   LET ovDateFormat = NULL
   LET ovDatetimeFormat = NULL
   LET ovTimeFormat = NULL
   LET ovMoneyFormat = NULL
   LET ovCurrencySymbol = NULL
   CALL invalidateFormats()
END FUNCTION #clearFormatOverrides

#Callers caching cell styles compare this against the generation their cache
#was built under, and drop the cache when the two differ.
PUBLIC FUNCTION getFormatGeneration() RETURNS INTEGER
   RETURN formatGeneration
END FUNCTION #getFormatGeneration

PRIVATE FUNCTION invalidateFormats() RETURNS ()
   LET formatGeneration = formatGeneration + 1
   LET envDateCode = NULL
   LET envCurrencyRead = FALSE
   LET dateFormatUSA = -1
END FUNCTION #invalidateFormats

#------------------------------------------------------------------------------
# Reading the environment
#------------------------------------------------------------------------------

#Splits on a single character delimiter, keeping the empty fields - DBFORMAT
#relies on position, so "*::.:" must yield four fields and not two.
PRIVATE FUNCTION splitFields(source STRING, delimiter STRING) RETURNS DYNAMIC ARRAY OF STRING
   DEFINE fields DYNAMIC ARRAY OF STRING
   DEFINE idx INTEGER
   DEFINE start INTEGER = 1
   DEFINE count INTEGER = 0

   IF source IS NULL THEN
      RETURN fields
   END IF

   FOR idx = 1 TO source.getLength()
      IF source.getCharAt(idx) == delimiter THEN
         LET count = count + 1
         LET fields[count] = source.subString(start, idx - 1)
         LET start = idx + 1
      END IF
   END FOR

   LET count = count + 1
   LET fields[count] = source.subString(start, source.getLength())

   RETURN fields

END FUNCTION #splitFields

#The Excel date format code implied by DBDATE. DBDATE is
#  { DM | MD } { Y2 | Y3 | Y4 } { / | - | . | 0 } [C1]
#  { Y2 | Y3 | Y4 } { DM | MD } { / | - | . | 0 } [C1]
#and defaults to "MDY4/" on desktop and server platforms. The separator is
#written as a quoted literal so that it survives the viewer's regional
#settings along with the field order.
PRIVATE FUNCTION dateCodeFromDBDATE() RETURNS STRING
   DEFINE work STRING
   DEFINE separator STRING
   DEFINE lastChar STRING
   DEFINE yearCode STRING
   DEFINE code STRING
   DEFINE idx INTEGER
   DEFINE currChr STRING

   IF envDateCode IS NOT NULL THEN
      RETURN envDateCode
   END IF

   LET work = FGL_GETENV("DBDATE")
   IF work IS NULL OR work.getLength() == 0 THEN
      #Desktop and server platforms default to the United States format
      LET work = "MDY4/"
   END IF
   LET work = work.toUpperCase()

   #The C1 modifier selects the Ming Guo calendar, which re-bases the year.
   #Excel has no such calendar, so fall back to an unambiguous ISO date
   #rather than write a year that would be read as Gregorian.
   IF work.getLength() >= 2
      AND work.subString(work.getLength() - 1, work.getLength()) == "C1" THEN
      LET envDateCode = "yyyy\"-\"mm\"-\"dd"
      RETURN envDateCode
   END IF

   #The separator always goes last. "0" means no separator at all.
   LET separator = "/"
   LET lastChar = work.subString(work.getLength(), work.getLength())
   CASE lastChar
      WHEN "/"
         LET separator = "/"
      WHEN "-"
         LET separator = "-"
      WHEN "."
         LET separator = "."
      WHEN "0"
         LET separator = ""
      OTHERWISE
         #No separator given, or an invalid one: the slash is the default
         LET lastChar = NULL
   END CASE
   IF lastChar IS NOT NULL THEN
      LET work = work.subString(1, work.getLength() - 1)
   END IF

   #Y2 is a two digit year; Y3 has no Excel equivalent, so widen it to four
   LET yearCode = "yyyy"
   LET idx = work.getIndexOf("Y", 1)
   IF idx > 0 AND idx < work.getLength() THEN
      IF work.getCharAt(idx + 1) == "2" THEN
         LET yearCode = "yy"
      END IF
   END IF

   #Walk the units in the order DBDATE gives them
   LET idx = 1
   WHILE idx <= work.getLength()
      LET currChr = work.getCharAt(idx)
      CASE currChr
         WHEN "D"
            LET code = appendDatePart(code, "dd", separator)
         WHEN "M"
            LET code = appendDatePart(code, "mm", separator)
         WHEN "Y"
            LET code = appendDatePart(code, yearCode, separator)
            #Step over the digit count that follows the Y
            LET idx = idx + 1
      END CASE
      LET idx = idx + 1
   END WHILE

   IF code IS NULL THEN
      LET code = "yyyy\"-\"mm\"-\"dd"
   END IF

   LET envDateCode = code
   RETURN envDateCode

END FUNCTION #dateCodeFromDBDATE

PRIVATE FUNCTION appendDatePart(code STRING, part STRING, separator STRING) RETURNS STRING

   IF code IS NULL THEN
      RETURN part
   END IF
   IF separator IS NULL OR separator.getLength() == 0 THEN
      RETURN code.append(part)
   END IF

   RETURN code.append("\"").append(separator).append("\"").append(part)

END FUNCTION #appendDatePart

#The currency symbol and its position, from DBFORMAT if it is set, otherwise
#from DBMONEY. DBFORMAT is "[front]:[thousands]:decimal:[back]" and DBMONEY is
#a leading or trailing symbol either side of a "." or "," decimal separator.
#With neither set, desktop and server platforms use a leading "$".
PRIVATE FUNCTION readCurrencyFromEnv() RETURNS ()
   DEFINE value STRING
   DEFINE fields DYNAMIC ARRAY OF STRING
   DEFINE idx INTEGER
   DEFINE currChr STRING

   IF envCurrencyRead THEN
      RETURN
   END IF
   LET envCurrencyRead = TRUE
   LET envCurrencySymbol = NULL
   LET envCurrencyIsFront = TRUE

   LET value = FGL_GETENV("DBFORMAT")
   IF value IS NOT NULL AND value.getLength() > 0 THEN
      LET fields = splitFields(value, ":")
      #An asterisk in a field means "no symbol here"
      IF fields.getLength() >= 1
         AND fields[1].getLength() > 0 AND fields[1] != "*" THEN
         LET envCurrencySymbol = fields[1]
         LET envCurrencyIsFront = TRUE
         RETURN
      END IF
      IF fields.getLength() >= 4
         AND fields[4].getLength() > 0 AND fields[4] != "*" THEN
         LET envCurrencySymbol = fields[4]
         LET envCurrencyIsFront = FALSE
      END IF
      #DBFORMAT is set, so DBMONEY is not consulted at all
      RETURN
   END IF

   LET value = FGL_GETENV("DBMONEY")
   IF value IS NOT NULL AND value.getLength() > 0 THEN
      #The decimal separator is mandatory; the symbol is whatever sits on
      #one side of it, and that side decides where the symbol is displayed
      FOR idx = 1 TO value.getLength()
         LET currChr = value.getCharAt(idx)
         IF currChr == "." OR currChr == "," THEN
            IF idx > 1 THEN
               LET envCurrencySymbol = value.subString(1, idx - 1)
               LET envCurrencyIsFront = TRUE
            ELSE
               IF idx < value.getLength() THEN
                  LET envCurrencySymbol = value.subString(idx + 1, value.getLength())
                  LET envCurrencyIsFront = FALSE
               END IF
            END IF
            RETURN
         END IF
      END FOR
      #No decimal separator: the whole value is not a usable DBMONEY setting
      RETURN
   END IF

   #Neither is set: desktop and server platforms default to a leading dollar
   LET envCurrencySymbol = "$"
   LET envCurrencyIsFront = TRUE

END FUNCTION #readCurrencyFromEnv

#------------------------------------------------------------------------------
# Building format codes
#------------------------------------------------------------------------------

#The precision and scale carried by a type name. base.TypeInfo reports
#"DECIMAL", "DECIMAL(8)", "DECIMAL(8,4)" and - because a scale of 2 is the
#default for MONEY - reports MONEY(12,2) as "MONEY(12)". A scale of -1 means
#the type did not state one.
PRIVATE FUNCTION parsePrecisionScale(fglDataType STRING) RETURNS (INTEGER, INTEGER)
   DEFINE openPos INTEGER
   DEFINE closePos INTEGER
   DEFINE commaPos INTEGER
   DEFINE inner STRING
   DEFINE precision INTEGER = -1
   DEFINE scale INTEGER = -1

   IF fglDataType IS NULL THEN
      RETURN precision, scale
   END IF

   LET openPos = fglDataType.getIndexOf("(", 1)
   LET closePos = fglDataType.getIndexOf(")", 1)
   IF openPos == 0 OR closePos <= openPos + 1 THEN
      RETURN precision, scale
   END IF

   LET inner = fglDataType.subString(openPos + 1, closePos - 1)
   LET commaPos = inner.getIndexOf(",", 1)
   IF commaPos > 0 THEN
      LET precision = inner.subString(1, commaPos - 1)
      LET scale = inner.subString(commaPos + 1, inner.getLength())
   ELSE
      LET precision = inner
   END IF

   RETURN precision, scale

END FUNCTION #parsePrecisionScale

#"#,##0.00" style body: grouped above three digits, with the requested number
#of decimal places. A scale of -1 gives a floating decimal part, which is what
#a bare DECIMAL or a DECIMAL(p) carries.
PRIVATE FUNCTION numericBody(precision INTEGER, scale INTEGER) RETURNS STRING
   DEFINE body STRING
   DEFINE idx INTEGER

   IF precision > 0 AND precision <= 3 THEN
      LET body = "##0"
   ELSE
      LET body = "#,##0"
   END IF

   CASE
      WHEN scale > 0
         LET body = body.append(".")
         FOR idx = 1 TO scale
            LET body = body.append("0")
         END FOR
      WHEN scale < 0
         #Scale not stated: show a decimal part only when the value has one
         LET body = body.append(".######")
   END CASE

   RETURN body

END FUNCTION #numericBody

#Wraps a number body so that negatives show in red inside parentheses, the
#convention the package has always used for its numeric columns.
PRIVATE FUNCTION redNegative(body STRING) RETURNS STRING
   RETURN SFMT("%1;[Red](%1)", body)
END FUNCTION #redNegative

PRIVATE FUNCTION createDecFormat(fglDataType STRING) RETURNS STRING
   DEFINE precision INTEGER
   DEFINE scale INTEGER

   CALL parsePrecisionScale(fglDataType) RETURNING precision, scale

   RETURN redNegative(numericBody(precision, scale))

END FUNCTION #createDecFormat

#Quotes a currency symbol for use inside a format code. A double quote in the
#symbol would end the literal, so it is escaped.
PRIVATE FUNCTION quoteSymbol(symbol STRING) RETURNS STRING
   DEFINE escaped STRING

   LET escaped = symbol.replaceAll("\"", "\"\"")

   RETURN SFMT("\"%1\"", escaped)

END FUNCTION #quoteSymbol

#The format code for a MONEY column: the currency symbol of the machine the
#application runs on, and the decimal places the column type carries. MONEY
#defaults to a scale of 2, unlike DECIMAL which is floating when unstated.
PUBLIC FUNCTION getMoneyFormat(fglDataType STRING) RETURNS STRING
   DEFINE precision INTEGER
   DEFINE scale INTEGER
   DEFINE body STRING
   DEFINE symbol STRING

   IF ovMoneyFormat IS NOT NULL THEN
      RETURN ovMoneyFormat
   END IF
   IF formatMode == cFormatModeViewer AND ovCurrencySymbol IS NULL THEN
      RETURN NULL
   END IF

   CALL parsePrecisionScale(fglDataType) RETURNING precision, scale
   IF scale < 0 THEN
      LET scale = 2
   END IF
   LET body = numericBody(precision, scale)

   IF formatMode == cFormatModeISO THEN
      RETURN redNegative(body)
   END IF

   IF ovCurrencySymbol IS NOT NULL THEN
      RETURN redNegative(quoteSymbol(ovCurrencySymbol).append(body))
   END IF

   CALL readCurrencyFromEnv()
   IF envCurrencySymbol IS NULL OR envCurrencySymbol.getLength() == 0 THEN
      RETURN redNegative(body)
   END IF

   IF envCurrencyIsFront THEN
      RETURN redNegative(quoteSymbol(envCurrencySymbol).append(body))
   END IF

   RETURN redNegative(body.append(quoteSymbol(envCurrencySymbol)))

END FUNCTION #getMoneyFormat

#The format code for a DATE column. NULL in cFormatModeViewer, where no code
#is written and the machine viewing the file picks the format.
PUBLIC FUNCTION getDateFormat() RETURNS STRING

   IF ovDateFormat IS NOT NULL THEN
      RETURN ovDateFormat
   END IF
   IF formatMode == cFormatModeViewer THEN
      RETURN NULL
   END IF
   IF formatMode == cFormatModeISO THEN
      RETURN "yyyy\"-\"mm\"-\"dd"
   END IF

   RETURN dateCodeFromDBDATE()

END FUNCTION #getDateFormat

#The time part a DATETIME type calls for. Genero has no environment setting
#for the time of day, so a 24 hour clock is used - it cannot be misread the
#way a 12 hour clock without a locale can.
PRIVATE FUNCTION timePartForType(fglDataType STRING) RETURNS STRING

   CASE
      WHEN fglDataType MATCHES "*TO MINUTE*"
         RETURN "hh:mm"
      WHEN fglDataType MATCHES "*TO HOUR*"
         RETURN "hh"
      WHEN fglDataType MATCHES "*FRACTION*"
         RETURN "hh:mm:ss.000"
      OTHERWISE
         RETURN "hh:mm:ss"
   END CASE

END FUNCTION #timePartForType

#The format code for a DATETIME YEAR TO ... column: the same date format the
#DATE columns use, so that two date columns side by side read alike.
PUBLIC FUNCTION getDatetimeFormat(fglDataType STRING) RETURNS STRING
   DEFINE code STRING

   IF ovDatetimeFormat IS NOT NULL THEN
      RETURN ovDatetimeFormat
   END IF

   LET code = getDateFormat()
   IF code IS NULL THEN
      RETURN NULL
   END IF

   RETURN code.append(" ").append(timePartForType(fglDataType))

END FUNCTION #getDatetimeFormat

#The format code for a DATETIME HOUR TO ... column.
PUBLIC FUNCTION getTimeFormat(fglDataType STRING) RETURNS STRING

   IF ovTimeFormat IS NOT NULL THEN
      RETURN ovTimeFormat
   END IF
   IF formatMode == cFormatModeViewer THEN
      RETURN NULL
   END IF

   RETURN timePartForType(fglDataType)

END FUNCTION #getTimeFormat

#------------------------------------------------------------------------------
# Applying a format to a cell style
#------------------------------------------------------------------------------

PUBLIC FUNCTION getCellStyleForDataType(workbook fgl_excel.workbookType, fglDataType STRING) RETURNS fgl_excel.cellStyleType
   DEFINE cellStyle fgl_excel.cellStyleType
   #Module-wide directive - applies to all functions below; propagates errors to the caller
   WHENEVER ERROR RAISE

   LET cellStyle = fgl_excel.style_create(workbook)
   CALL setCellStyleForDataType(workbook, cellStyle, fglDataType)
   RETURN cellStyle

END FUNCTION #getCellStyleForDataType

#Writes an explicit format code into the workbook, or falls back to the
#built-in id when there is no code - which is what cFormatModeViewer asks for.
PRIVATE FUNCTION applyFormat(workbook fgl_excel.workbookType,
                             cellStyle fgl_excel.cellStyleType,
                             formatCode STRING,
                             builtinFormat SMALLINT) RETURNS ()

   IF formatCode IS NULL THEN
      CALL fgl_excel.set_builtin_style_format(cellStyle, builtinFormat)
   ELSE
      CALL fgl_excel.set_style_format(workbook, cellStyle, formatCode)
   END IF

END FUNCTION #applyFormat

PUBLIC FUNCTION setCellStyleForDataType(workbook fgl_excel.workbookType,
                                        cellStyle fgl_excel.cellStyleType,
                                        fglDataType STRING) RETURNS ()

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
         CALL applyFormat(workbook, cellStyle,
                          getMoneyFormat(fglDataType), fgl_excel.cMoneyFormat)
      WHEN fglDataType MATCHES "*FLOAT*"
         #Handle floating point Field Formatting
         CALL fgl_excel.set_builtin_style_format(cellStyle, fgl_excel.cDecimalFormat)
      WHEN fglDataType == "DATE"
         #get the date string format and build a cell style from it
         CALL applyFormat(workbook, cellStyle,
                          getDateFormat(), fgl_excel.cDateFormat)
      WHEN fglDataType MATCHES "DATETIME YEAR*"
         #get the datetime string format and build a cell style from it
         CALL applyFormat(workbook, cellStyle,
                          getDatetimeFormat(fglDataType), fgl_excel.cDatetimeFormat)
      WHEN fglDataType MATCHES "DATETIME HOUR*"
         #get the datetime string format and build a cell style from it
         CALL applyFormat(workbook, cellStyle,
                          getTimeFormat(fglDataType), fgl_excel.cTimeFormat)

    END CASE

END FUNCTION #setCellStyleForDataType

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



