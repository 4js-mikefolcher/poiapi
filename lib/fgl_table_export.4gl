PACKAGE com.fourjs.poiapi

IMPORT util
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_helper
IMPORT FGL com.fourjs.poiapi.fgl_spreadsheet_xapi

PRIVATE TYPE TColumnMetaInfo RECORD
	colTitle 	STRING,
	colType  	STRING,
	colName  	STRING,
	colIdx   	INTEGER,
	colPosition INTEGER,
	colHidden   INTEGER,
	colAggType  STRING,
	fieldIdx    INTEGER
END RECORD

PRIVATE TYPE TTableSort RECORD
	colIdx INTEGER,
	sortOrder STRING,
	colPosition INTEGER
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

		#Get the header information
		FOR idx = 1 TO tableNode.getChildCount()

			#Get a reference to the table node
			VAR columnNode = tableNode.getChildByIndex(idx)
			IF columnNode.getTagName() == "PhantomColumn" THEN
				CONTINUE FOR
			END IF

			LET fieldIdx += 1
			INITIALIZE columnInfo.* TO NULL

			#Put the attributes we need into the columnInfo record
			CALL columnInfo.setFromNode(columnNode)
			LET columnInfo.colIdx = idx
			LET columnInfo.fieldIdx = fieldIdx

			#Add to the column header array
			LET columnHeaders[columnInfo.colPosition].colTitle = columnInfo.colTitle
			LET columnHeaders[columnInfo.colPosition].colCalc = getAggregateType(columnInfo.colAggType)

			#Add to the column info array
			LET colInfoList[columnInfo.colPosition] = columnInfo

			#If the table is sorted, set the column position
			IF tableSort.colIdx > 0 AND tableSort.colIdx == idx THEN
				LET tableSort.colPosition = columnInfo.colPosition
			END IF

		END FOR

		#Prune hidden columns from the colInfoList (Where hidden is 1)
		WHILE (idx := colInfoList.search("colHidden", 1)) > 0
				CALL colInfoList.deleteElement(idx)
				CALL columnHeaders.deleteElement(idx)
		END WHILE

		#Prune hidden columns from the colInfoList (Where hidden is 2)
		WHILE (idx := colInfoList.search("colHidden", 2)) > 0
				CALL colInfoList.deleteElement(idx)
				CALL columnHeaders.deleteElement(idx)
		END WHILE

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

			#Loop through each column in the colInfoList array
			FOR idx = 1 TO colInfoList.getLength()
				VAR dataName STRING = dataRow.name(colInfoList[idx].colIdx)
				VAR dataValue STRING = dataRow.get(dataName)
				CALL jsonRow.put(colInfoList[idx].colName, dataValue)

				#For the first row, build the field type XML structure
				IF valueIdx == 1 THEN
					#Add column metadata on the first row only
					VAR child = recDef.createChild("Field")
					CALL child.setAttribute("name", colInfoList[idx].colName)
					CALL child.setAttribute("type", colInfoList[idx].colType)
				END IF

				#If the data is sorted on the frontend, sort the sort column value in the sortList
				IF tableSort.colIdx > 0 AND tableSort.colIdx == colInfoList[idx].fieldIdx THEN
					LET sortColumn = sortList[valueIdx].setValue(dataValue, colInfoList[idx].colType)
				END IF

			END FOR

			#Set the record definition when we are on the first row
			IF valueIdx == 1 THEN
				CALL excelApi.setRecordDefinition(recDef)
			END IF

			IF tableSort.colIdx > 0 THEN
				#If sort is specified in the front-end, save the jsonRow in the sortList
				LET sortList[valueIdx].jsonRow = jsonRow
			ELSE
				#If no sort is specified on the front-end, add the row to the Excel API
				CALL excelApi.addDataRow(jsonRow)
			END IF
		END FOR

		IF tableSort.colIdx > 0 THEN
			#If sorted on the front-end, sort the sortList and make a second pass to add to the excel sheet
			VAR reverseSort = IIF(tableSort.sortOrder.toLowerCase() == "desc", TRUE, FALSE)
			CALL sortList.sort(sortColumn, reverseSort)
			FOR idx = 1 TO sortList.getLength()
				CALL excelApi.addDataRow(sortList[idx].jsonRow)
			END FOR
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
	LET self.colHidden = node.getAttribute("hidden")
	LET self.colName = node.getAttribute("colName")
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