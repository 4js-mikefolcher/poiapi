PACKAGE com.fourjs.poiapi

PUBLIC TYPE TStringArray DYNAMIC ARRAY OF STRING
PUBLIC TYPE TIntArray DYNAMIC ARRAY OF INTEGER

PUBLIC TYPE TRowStack RECORD
   rowList DYNAMIC ARRAY OF TIntArray,
   titleList TStringArray
END RECORD

PUBLIC FUNCTION (self TRowStack) init() RETURNS ()

   CALL self.rowList.clear()
   CALL self.titleList.clear()
   CALL self.rowList.appendElement() #add for grand totals
   LET self.titleList[1] = "Grand Totals"

END FUNCTION #init

PUBLIC FUNCTION (self TRowStack) pushGroup(title STRING) RETURNS ()

   CALL self.rowList.appendElement()
   CALL self.titleList.appendElement()
   LET self.titleList[self.titleList.getLength()] = title

END FUNCTION #pushGroup

PUBLIC FUNCTION (self TRowStack) addRow(rowIdx INTEGER) RETURNS ()
   DEFINE idx INTEGER
   DEFINE innerIdx INTEGER

   FOR idx = 1 TO self.currentLevel()
      LET innerIdx = self.rowList[idx].getLength() + 1
      LET self.rowList[idx][innerIdx] = rowIdx + 1
   END FOR

END FUNCTION #addRow

PUBLIC FUNCTION (self TRowStack) popGroup() RETURNS (STRING, TIntArray)
   DEFINE idx INTEGER
   DEFINE title STRING
   DEFINE currentList TIntArray

   LET idx = self.currentLevel()
   CALL self.rowList[idx].copyTo(currentList)
   CALL self.rowList.deleteElement(idx)
   LET title = self.titleList[idx]
   CALL self.titleList.deleteElement(idx)

   RETURN title, currentList

END FUNCTION

PUBLIC FUNCTION (self TRowStack) currentLevel() RETURNS (INTEGER)

   RETURN self.rowList.getLength()

END FUNCTION #currentLevel

PUBLIC FUNCTION getFormattedStrings(columnId STRING, rowList TIntArray) RETURNS TStringArray
   DEFINE idx INTEGER
   DEFINE startIdx INTEGER = 0
   DEFINE rangeIdx INTEGER = 0
   DEFINE prevIdx INTEGER = 0
   DEFINE rangeList TStringArray
   CONSTANT cRangeFormat = "(%1%2:%1%3)"

   FOR idx = 1 TO rowList.getLength()
      IF startIdx == 0 THEN
         LET startIdx = rowList[idx]
      END IF
      IF prevIdx < (rowList[idx] - 1) AND prevIdx > 0 THEN
         CALL rangeList.appendElement()
         LET rangeIdx = rangeList.getLength()
         LET rangeList[rangeIdx] = SFMT(cRangeFormat, columnId, startIdx, prevIdx)
         LET startIdx = rowList[idx]
      END IF
      LET prevIdx = rowList[idx]
   END FOR

   IF prevIdx > 0 THEN
      CALL rangeList.appendElement()
      LET rangeIdx = rangeList.getLength()
      LET rangeList[rangeIdx] = SFMT(cRangeFormat, columnId, startIdx, prevIdx)
   END IF

   RETURN rangeList

END FUNCTION #getFormattedStrings