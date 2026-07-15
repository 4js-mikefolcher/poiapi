PACKAGE com.fourjs.poiapi

IMPORT util
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_helper
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_xapi

PRIVATE TYPE TColumnMetaInfo RECORD
	colTitle 	STRING,
	colType  	STRING,
	colName  	STRING,
	colFieldName STRING,
	colIdx   	INTEGER,
	colPosition INTEGER,
	colHidden   BOOLEAN,
	colAggType  STRING,
	fieldIdx    INTEGER
END RECORD

PRIVATE TYPE TTableSort RECORD
	colIdx INTEGER,
	sortOrder STRING,
	colPosition INTEGER
END RECORD

#One resolved key of a Genero 6 multi-column sort specification
PRIVATE TYPE TSortKey RECORD
	colName STRING,
	colType STRING,
	descending BOOLEAN
END RECORD

PRIVATE TYPE TDataSort RECORD
	stringField STRING,
	numberField DECIMAL(20),
	dateField DATE,
	datetimeField DATETIME YEAR TO FRACTION (5),
	jsonRow util.JSONObject
END RECORD

PUBLIC FUNCTION tableExcelExport(tableName STRING, jsonData util.JSONArray) RETURNS (STRING)
	DEFINE columnHeaders DYNAMIC ARRAY OF TColumnInfo
	DEFINE excelApi TSpreadsheetXtend
	DEFINE columnInfo TColumnMetaInfo
	DEFINE colInfoList DYNAMIC ARRAY OF TColumnMetaInfo
	DEFINE tableSort TTableSort
	DEFINE sortList DYNAMIC ARRAY OF TDataSort
	DEFINE sortKeys DYNAMIC ARRAY OF TSortKey
	DEFINE useMultiSort BOOLEAN
	DEFINE sortKeyIdx INTEGER
	DEFINE sortSpecStr STRING

	VAR winElement = ui.Window.getCurrent()
	VAR root = winElement.getForm().getNode()
	VAR tableList = root.selectByTagName("Table")

	VAR tableNode om.DomNode = NULL
	VAR idx = 0
	VAR tableFound = FALSE
	FOR idx = 1 TO tableList.getLength()
		LET tableNode = tableList.item(idx)
		IF tableNode.getAttribute("tabName") == tableName THEN
			LET tableFound = TRUE
			EXIT FOR
		END IF
	END FOR

	IF tableFound THEN

		#Get Sort Information
		LET tableSort.colIdx = NVL(tableNode.getAttribute("sortColumn"), -1) + 1
		LET tableSort.sortOrder = NVL(tableNode.getAttribute("sortType"), "none")

		VAR fieldIdx = 0
		VAR minIdx = 0

		#Get the header information
		FOR idx = 1 TO tableNode.getChildCount()

			CALL debugOutput(SFMT("Header Column # %1", idx))

			#Get a reference to the table node
			VAR columnNode = tableNode.getChildByIndex(idx)
			IF columnNode.getTagName() == "PhantomColumn" THEN
				#Skip the column if it's a Phantom Column
				CONTINUE FOR
			END IF

			LET fieldIdx += 1
			INITIALIZE columnInfo.* TO NULL

			#Put the attributes we need into the columnInfo record
			CALL columnInfo.setFromNode(columnNode)
			LET columnInfo.colIdx = idx
			LET columnInfo.fieldIdx = fieldIdx

			#Add to the column header array, using the position from the AUI tree
			LET columnHeaders[columnInfo.colPosition].colTitle = columnInfo.colTitle
			LET columnHeaders[columnInfo.colPosition].colCalc = getAggregateType(columnInfo.colAggType)

			#Add to the column info array
			LET colInfoList[columnInfo.colPosition] = columnInfo

			IF minIdx == 0 OR columnInfo.colPosition < minIdx THEN
				LET minIdx = columnInfo.colPosition
			END IF

			CALL debugOutput(SFMT("Column Title: %1", columnInfo.colTitle))
			CALL debugOutput(SFMT("Column Aggregate Type: %1", columnInfo.colAggType))
			CALL debugOutput(SFMT("Column Position: %1", columnInfo.colPosition))

			#If the table is sorted, set the column position
			IF tableSort.colIdx > 0 AND tableSort.colIdx == idx THEN
				LET tableSort.colPosition = columnInfo.colPosition
			END IF

		END FOR

		LET idx = 1
		WHILE idx < minIdx
			#Fix indexing issue with tabIndex
			CALL colInfoList.deleteElement(idx)
			CALL columnHeaders.deleteElement(idx)
			LET minIdx -= 1
			IF tableSort.colPosition > 0 THEN
				LET tableSort.colPosition -= 1
			END IF
		END WHILE

		#Prune hidden columns from the colInfoList
		WHILE (idx := colInfoList.search("colHidden", TRUE)) > 0
				CALL colInfoList.deleteElement(idx)
				CALL columnHeaders.deleteElement(idx)
		END WHILE

		#Genero 6 records multi-column sorts in the "sortSpec" attribute.
		#When present it fully describes the ordering, so prefer it over the
		#legacy single-column "sortColumn"/"sortType" attributes.
		LET sortSpecStr = tableNode.getAttribute("sortSpec")
		IF sortSpecStr.getLength() > 0 THEN
			CALL parseSortSpec(sortSpecStr, colInfoList) RETURNING sortKeys
		END IF
		LET useMultiSort = (sortKeys.getLength() > 0)

		#Initialize the excel document
		CALL excelApi.init()
		CALL excelApi.setColumnInfo(columnHeaders)
		CALL excelApi.setTitle("Table Export")
		CALL excelApi.addSubTitle(winElement.getText())
		VAR recDef = om.DomDocument.create("Record").getDocumentElement()

		#Get the table data
		VAR sortColumn = ""
		LET idx = 1
		VAR valueIdx = 1

		#Loop through each row of data
		FOR valueIdx = 1 TO jsonData.getLength()
			VAR jsonRow = util.JSONObject.create()
			VAR dataRow util.JSONObject = jsonData.get(valueIdx)

			CALL debugOutput(SFMT("Data Row:\n%1", util.JSON.format(dataRow.toString())))

			#Loop through each column in the colInfoList array
			FOR idx = 1 TO colInfoList.getLength()
				VAR dataName STRING = dataRow.name(colInfoList[idx].colIdx)
				VAR dataValue STRING = dataRow.get(dataName)

				CALL debugOutput(SFMT("dataName: %1", dataName))
				CALL debugOutput(SFMT("dataValue: %1", dataValue))
				CALL jsonRow.put(colInfoList[idx].colName, dataValue)

				#For the first row, build the field type XML structure
				IF valueIdx == 1 THEN
					#Add column metadata on the first row only
					VAR child = recDef.createChild("Field")
					CALL child.setAttribute("name", colInfoList[idx].colName)
					CALL child.setAttribute("type", colInfoList[idx].colType)
				END IF

				#If the data is sorted on the frontend (legacy single-column
				#path only), capture the sort column value in the sortList
				IF NOT useMultiSort AND tableSort.colIdx > 0 AND tableSort.colIdx == colInfoList[idx].fieldIdx THEN
					LET sortColumn = sortList[valueIdx].setValue(dataValue, colInfoList[idx].colType)
				END IF

			END FOR

			#Set the record definition when we are on the first row
			IF valueIdx == 1 THEN
				CALL excelApi.setRecordDefinition(recDef)
				CALL debugOutput(recDef.toString())
			END IF

			CALL debugOutput(SFMT("Excel JSON Row:\n%1", util.JSON.format(jsonRow.toString())))

			IF useMultiSort OR tableSort.colIdx > 0 THEN
				#If sort is specified in the front-end, save the jsonRow in the
				#sortList for a second (sorted) pass
				LET sortList[valueIdx].jsonRow = jsonRow
			ELSE
				#If no sort is specified on the front-end, add the row to the Excel API
				CALL excelApi.addDataRow(jsonRow)
			END IF
		END FOR

		IF useMultiSort THEN
			#Reproduce the UI's multi-column ordering. Genero array sort is
			#stable, so we chain single-column sorts from the least- to the
			#most-significant key: rows that tie on a more significant key keep
			#the order established by the less significant keys already applied.
			FOR sortKeyIdx = sortKeys.getLength() TO 1 STEP -1
				VAR sortFieldName = ""
				FOR idx = 1 TO sortList.getLength()
					LET sortFieldName =
						sortList[idx].setValue(
							NVL(sortList[idx].jsonRow.get(sortKeys[sortKeyIdx].colName), ""),
							sortKeys[sortKeyIdx].colType)
				END FOR
				IF sortFieldName.getLength() > 0 THEN
					CALL sortList.sort(sortFieldName, sortKeys[sortKeyIdx].descending)
				END IF
			END FOR
			FOR idx = 1 TO sortList.getLength()
				CALL excelApi.addDataRow(sortList[idx].jsonRow)
			END FOR
		ELSE
			IF tableSort.colIdx > 0 THEN
				#If sorted on the front-end, sort the sortList and make a second pass to add to the excel sheet
				VAR reverseSort = IIF(tableSort.sortOrder.toLowerCase() == "desc", TRUE, FALSE)
				CALL sortList.sort(sortColumn, reverseSort)
				FOR idx = 1 TO sortList.getLength()
					CALL excelApi.addDataRow(sortList[idx].jsonRow)
				END FOR
			END IF
		END IF

	END IF

	#Get the Excel file path
	VAR excelFilename = ""
	IF excelApi.createSpreadsheet() THEN
		LET excelFilename = excelApi.getFilename()
	END IF

	#Return the Excel file path
	RETURN excelFilename

END FUNCTION #tableExcelExport

#Parse a Genero 6 multi-column sort specification into ordered sort keys.
#Format (space separated): "<colname>:{+/-}/{P/U}"
#  e.g. "integerfield:-/U booleanfield:+/U"  (a space after ':' is tolerated)
#  '+' = ascending, '-' = descending. The {P/U} type (SORT_GROUP_BY vs
#  SORT_USER) does not affect row ordering, so it is ignored here.
#Each spec column is matched case-insensitively against the exported columns
#(by colName or the AUI field name); columns that are not exported (e.g.
#hidden/pruned) are skipped. The list order is the sort priority, primary
#key first.
PRIVATE FUNCTION parseSortSpec(sortSpec STRING, colInfoList DYNAMIC ARRAY OF TColumnMetaInfo) RETURNS DYNAMIC ARRAY OF TSortKey
	DEFINE sortKeys DYNAMIC ARRAY OF TSortKey
	DEFINE rawTokens DYNAMIC ARRAY OF STRING
	DEFINE tokenizer base.StringTokenizer
	DEFINE i, k, colonPos INTEGER
	DEFINE token, colName, spec STRING
	DEFINE descending BOOLEAN

	#Split on whitespace (StringTokenizer collapses runs of spaces)
	LET tokenizer = base.StringTokenizer.create(sortSpec, " ")
	WHILE tokenizer.hasMoreTokens()
		LET rawTokens[rawTokens.getLength() + 1] = tokenizer.nextToken()
	END WHILE

	LET i = 1
	WHILE i <= rawTokens.getLength()
		LET token = rawTokens[i]
		LET colonPos = token.getIndexOf(":", 1)
		IF colonPos == 0 THEN
			#Not the start of an entry - skip defensively
			LET i = i + 1
			CONTINUE WHILE
		END IF

		LET colName = token.subString(1, colonPos - 1)
		#The "{+/-}/{P/U}" part is either attached ("col:+/U") or, when a
		#space follows the colon ("col: +/U"), it is the next token.
		IF colonPos >= token.getLength() THEN
			LET spec = ""
		ELSE
			LET spec = token.subString(colonPos + 1, token.getLength())
		END IF
		#An empty string literal is NULL in BDL, so test the length: when the
		#colon ended the token ("col:"), the direction is the next token.
		IF spec.getLength() == 0 AND i < rawTokens.getLength() THEN
			LET i = i + 1
			LET spec = rawTokens[i]
		END IF

		IF spec.getLength() > 0 THEN
			LET descending = (spec.getCharAt(1) == "-")

			#Resolve the column against the exported columns
			FOR k = 1 TO colInfoList.getLength()
				IF colInfoList[k].colName.toLowerCase() == colName.toLowerCase()
				   OR colInfoList[k].colFieldName.toLowerCase() == colName.toLowerCase() THEN
					LET sortKeys[sortKeys.getLength() + 1].colName = colInfoList[k].colName
					LET sortKeys[sortKeys.getLength()].colType = colInfoList[k].colType
					LET sortKeys[sortKeys.getLength()].descending = descending
					EXIT FOR
				END IF
			END FOR
		END IF

		LET i = i + 1
	END WHILE

	RETURN sortKeys
END FUNCTION #parseSortSpec

PRIVATE FUNCTION getAggregateType(aggregateType STRING) RETURNS STRING

	VAR excelFormula = cExcelNone
	CASE aggregateType.toUpperCase()
		WHEN "SUM"
			LET excelFormula = cExcelSum
		WHEN "AVG"
			LET excelFormula = cExcelAvg
		WHEN "MIN"
			LET excelFormula = cExcelMin
		WHEN "MAX"
			LET excelFormula = cExcelMax
		WHEN "COUNT"
			LET excelFormula = cExcelCount
		OTHERWISE
			LET excelFormula = cExcelNone
	END CASE

	RETURN excelFormula

END FUNCTION #getAggregateType

PRIVATE FUNCTION (self TColumnMetaInfo) setFromNode(node om.DomNode) RETURNS ()

	#Get the attributes we need
	LET self.colTitle = node.getAttribute("text")
	LET self.colType = node.getAttribute("varType")
	LET self.colHidden = IIF(node.getAttribute("hidden") > 0, TRUE, FALSE)
	LET self.colName = node.getAttribute("colName")
	LET self.colFieldName = node.getAttribute("name")
	LET self.colPosition = node.getAttribute("tabIndex")
	LET self.colAggType = NVL(node.getAttribute("aggregateType"), "none")

END FUNCTION #setFromNode

PRIVATE FUNCTION (self TDataSort) setValue(dataValue STRING, dataType STRING) RETURNS STRING

	VAR sortColumn = ""
	CASE
		WHEN dataType MATCHES "DEC*"
			LET self.numberField = dataValue
			LET sortColumn = "numberField"
		WHEN dataType MATCHES "*INT*"
			LET self.numberField = dataValue
			LET sortColumn = "numberField"
		WHEN dataType MATCHES "*FLOAT*"
			LET self.numberField = dataValue
			LET sortColumn = "numberField"
		WHEN dataType MATCHES "MONEY*"
			LET self.numberField = dataValue
			LET sortColumn = "numberField"
		OTHERWISE
			LET self.stringField = dataValue
			LET sortColumn = "stringField"
	END CASE

	RETURN sortColumn

END FUNCTION #setValue

PRIVATE CONSTANT cDebugMode = FALSE
PRIVATE FUNCTION debugOutput(outputMessage STRING) RETURNS ()

	IF cDebugMode THEN
		DISPLAY outputMessage
	END IF

END FUNCTION #debugOutput
