PACKAGE com.fourjs.poiapi

IMPORT JAVA java.io.FileOutputStream
IMPORT JAVA java.io.FileInputStream

IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFWorkbook
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFSheet
IMPORT JAVA org.apache.poi.ss.usermodel.Row
IMPORT JAVA org.apache.poi.ss.usermodel.Cell
IMPORT JAVA org.apache.poi.ss.usermodel.CellStyle
IMPORT JAVA org.apache.poi.ss.usermodel.HorizontalAlignment
IMPORT JAVA org.apache.poi.ss.usermodel.Font
IMPORT JAVA org.apache.poi.ss.usermodel.PrintSetup
IMPORT JAVA org.apache.poi.ss.usermodel.DataFormat
IMPORT JAVA org.apache.poi.ss.util.CellRangeAddress
IMPORT JAVA org.apache.poi.ss.usermodel.IndexedColors
IMPORT JAVA org.apache.poi.ss.usermodel.FillPatternType
IMPORT JAVA org.apache.poi.ss.usermodel.BorderStyle
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFPrintSetup
IMPORT JAVA org.apache.poi.ss.usermodel.Header
IMPORT JAVA org.apache.poi.ss.usermodel.Footer
IMPORT JAVA org.apache.poi.ss.util.CellRangeAddress
IMPORT JAVA org.apache.poi.ss.usermodel.PageOrder
IMPORT JAVA java.util.Date
IMPORT JAVA java.util.Calendar
IMPORT JAVA java.text.SimpleDateFormat
IMPORT JAVA java.time.LocalDateTime
IMPORT JAVA java.time.ZoneId




PUBLIC TYPE workbookType XSSFWorkbook
PUBLIC TYPE sheetType XSSFSheet
PUBLIC TYPE rowType Row
PUBLIC TYPE cellType Cell
PUBLIC TYPE cellStyleType CellStyle
PUBLIC TYPE fontType Font
PUBLIC TYPE cellFormat DataFormat
PUBLIC TYPE cellRangeAddressType CellRangeAddress
PUBLIC TYPE excelColors IndexedColors
PUBLIC TYPE excelPatterns FillPatternType
PUBLIC TYPE borderStyleType BorderStyle
PUBLIC TYPE pageSetup XSSFPrintSetup
PUBLIC TYPE headerType Header
PUBLIC TYPE footerType Footer
PUBLIC TYPE pageOrderType PageOrder


#Constants that corralate to the Built-in Cell Formats
PUBLIC CONSTANT cDecimalFormat = 40
PUBLIC CONSTANT cMoneyFormat = 8
PUBLIC CONSTANT cIntegerFormat = 38
PUBLIC CONSTANT cDateFormat = 14
PUBLIC CONSTANT cTimeFormat = 20
PUBLIC CONSTANT cDatetimeFormat = 22


FUNCTION workbook_create()
    RETURN XSSFWorkbook.create()
END FUNCTION



FUNCTION workbook_writeToFile(w, filename)
DEFINE w workbookType
DEFINE filename STRING
DEFINE fo FileOutputStream

    LET fo = FileOutputStream.create(filename)
    CALL w.write(fo)
    CALL fo.close()
END FUNCTION


FUNCTION workbook_open(filename)
DEFINE filename STRING
DEFINE fi FileInputStream
DEFINE w workbookType

    LET fi = FileInputStream.create(filename)
    LET w = XSSFWorkbook.create(fi)
    RETURN w
END FUNCTION



FUNCTION workbook_createsheet(w)
DEFINE w workbookType
DEFINE s sheetType

    LET s= w.createSheet()
    RETURN s
END FUNCTION



FUNCTION sheet_createrow(s,idx)
DEFINE s sheetType
DEFINE idx INTEGER
DEFINE r rowType
    LET r = s.createRow(idx)
    RETURN r
END FUNCTION



FUNCTION sheet_autosizecolumn(s, c)
DEFINE s sheetType
DEFINE c INTEGER

    CALL s.autoSizeColumn(c)
END FUNCTION

FUNCTION sheet_columnwidth_set(s, c, wf)
DEFINE s sheetType
DEFINE c INTEGER
DEFINE wf FLOAT
DEFINE wi INTEGER

    LET wi = 256.0*wf
    CALL s.setColumnWidth(c,wi)
END FUNCTION



FUNCTION row_createcell(r,idx)
DEFINE r rowType
DEFINE idx INTEGER
DEFINE c cellType
    LET c = r.createCell(idx)
    RETURN c
END FUNCTION


FUNCTION cell_value_set(c, v)
DEFINE c cellType
DEFINE v STRING
    CALL c.setCellValue(v)
END FUNCTION



FUNCTION cell_number_set(c, v)
DEFINE c cellType
DEFINE v FLOAT
  
    CALL c.setCellValue(v)
END FUNCTION


FUNCTION cell_date_set(c, v)
DEFINE c cellType
DEFINE v STRING
DEFINE javaDate java.util.Date
DEFINE dateFormatter SimpleDateFormat
DEFINE baselineDate java.util.Date

    IF v IS NULL THEN
        CALL c.setCellValue(v)
    ELSE
        LET dateFormatter = java.text.SimpleDateFormat.create("yyyy-MM-dd")
        LET baselineDate = dateFormatter.parse("1900-01-01")
        LET javaDate = dateFormatter.parse(v)
        IF javaDate.before(baselineDate) THEN
           CALL c.setCellValue("")
        ELSE
           CALL c.setCellValue(javaDate)
        END IF
    END IF
    
END FUNCTION

FUNCTION cell_datetime_set(c, v)
DEFINE c cellType
DEFINE v STRING
DEFINE javaDatetime Calendar
DEFINE dateFormatter SimpleDateFormat

    IF v IS NULL THEN
        CALL c.setCellValue(v)
    ELSE
        LET dateFormatter = java.text.SimpleDateFormat.create("yyyy-MM-dd HH:mm:ss")
        LET javaDatetime = Calendar.getInstance()
        CALL javaDatetime.setTime(dateFormatter.parse(v))
        CALL c.setCellValue(LocalDateTime.ofInstant(javaDatetime.toInstant(), ZoneId.systemDefault()))
    END IF
    
END FUNCTION

FUNCTION cell_time_set(c, v)
DEFINE c cellType
DEFINE v STRING
DEFINE javaDatetime Calendar
DEFINE dateFormatter SimpleDateFormat

    IF v IS NULL THEN
        CALL c.setCellValue(v)
    ELSE
        LET dateFormatter = java.text.SimpleDateFormat.create("HH:mm:ss")               
        LET javaDatetime = Calendar.getInstance()
        CALL javaDatetime.setTime(dateFormatter.parse(v))
        CALL c.setCellValue(LocalDateTime.ofInstant(javaDatetime.toInstant(), ZoneId.systemDefault()))
    END IF
    
END FUNCTION

FUNCTION cell_formula_set(c, v)
DEFINE c cellType
DEFINE v STRING
    CALL c.setCellFormula(v)
END FUNCTION


-- Map A to 0 B to 1, Z to 25, AA to 26, AZ to 51
FUNCTION column2row(col)
DEFINE col STRING

    CASE
        WHEN col MATCHES "[A-Z]" 
            RETURN ORD(col) - 65
        WHEN col MATCHES "[A-Z][A-Z]"
            RETURN ((ORD(col.subString(1,1))-65)*26) + (ORD(col.subString(2,2)) - 65)
    END CASE
    RETURN -1
END FUNCTION
    



FUNCTION cell_style_set(c, s)
DEFINE c cellType
DEFINE s cellStyleType
    CALL c.setCellStyle(s)
END FUNCTION

FUNCTION cell_style_format_create(w workbookType, fmt STRING) RETURNS (cellStyleType)
DEFINE c cellFormat
DEFINE s cellStyleType

	LET s = w.createCellStyle()
	LET c = w.createDataFormat()
	CALL s.setDataFormat(c.getFormat(fmt))

	RETURN s

END FUNCTION

FUNCTION cell_style_builtin_create(w workbookType, fmt SMALLINT) RETURNS (cellStyleType)
DEFINE s cellStyleType

	LET s = w.createCellStyle()
	CALL s.setDataFormat(fmt)

	RETURN s
END FUNCTION

FUNCTION font_create(w)
DEFINE w workbookType
DEFINE f fontType
    LET f = w.createFont()
    RETURN f
END FUNCTION

FUNCTION font_set(f, a, v)
DEFINE f fontType
DEFINE a STRING
DEFINE v STRING

    CASE 
        WHEN a="weight" AND v="bold"
            CALL f.setBold(true)
        WHEN a="weight" AND v="normal"
            CALL f.setBold(false)
        -- add more as required
    END CASE
END FUNCTION



FUNCTION style_create(w)
DEFINE w workbookType
DEFINE s cellStyleType
    LET s = w.createCellStyle()
    RETURN s
END FUNCTION



FUNCTION style_set(s, a, v)
DEFINE s cellStyleType
DEFINE a STRING
DEFINE v STRING

    CASE 
        WHEN a="alignment" AND v="center"
            CALL s.setWrapText(TRUE)                                       #4JS timecard open issuses 01052022 #23 header word wrapping on 01062022
            CALL s.setAlignment(HorizontalAlignment.CENTER)
            
        WHEN a="alignment" AND v="left"
            CALL s.setAlignment(HorizontalAlignment.LEFT)

        WHEN a="alignment" AND v="right"
            CALL s.setAlignment(HorizontalAlignment.RIGHT)

        WHEN a="alignment" AND v="justify"
            CALL s.setAlignment(HorizontalAlignment.JUSTIFY)

        WHEN a="alignment" AND v="general"
            CALL s.setAlignment(HorizontalAlignment.GENERAL)

        -- add more as required
    END CASE
END FUNCTION


FUNCTION style_font_set(s,f)
DEFINE s cellStyleType
DEFINE f fontType
    CALL s.setFont(f)
END FUNCTION


FUNCTION workbook_fit_to_page(old_filename STRING, new_filename STRING)
DEFINE w workbookType
DEFINE s sheetType

DEFINE ps PrintSetup
CONSTANT SHORT_ONE SMALLINT = 1

    LET w = workbook_open(old_filename)
    LET s = w.getSheetAt(0)
    CALL s.setFitToPage(true)
    CALL s.setAutobreaks(true)
    LET ps = s.getPrintSetup()
    CALL ps.setFitWidth(SHORT_ONE)
    CALL ps.setFitHeight(SHORT_ONE)

    CALL workbook_writeToFile(w, new_filename)
END FUNCTION

PUBLIC FUNCTION merge_cells(s sheetType, rowIdx INTEGER, startIdx INTEGER, endIdx INTEGER) RETURNS ()

   CALL s.addMergedRegion(cellRangeAddressType.create(rowIdx, rowIdx, startIdx, endIdx))

END FUNCTION #merge_cells

PUBLIC FUNCTION set_background_color(cellStyle cellStyleType, bgColor SMALLINT) RETURNS ()

   CALL cellStyle.setFillForegroundColor(bgColor)
   CALL cellStyle.setFillPattern(excelPatterns.SOLID_FOREGROUND)

END FUNCTION #set_background_color

PUBLIC FUNCTION get_background_color(cellStyle cellStyleType) RETURNS (SMALLINT)

   RETURN cellStyle.getFillBackgroundColor()

END FUNCTION #get_background_color

PUBLIC FUNCTION set_border_style(cellStyle cellStyleType, topBorder BOOLEAN, rightBorder BOOLEAN, bottomBorder BOOLEAN, leftBorder BOOLEAN) RETURNS ()

   IF topBorder THEN
      CALL cellStyle.setBorderTop(borderStyleType.THIN)
      CALL cellStyle.setTopBorderColor(excelColors.BLACK.index)
   END IF

   IF rightBorder THEN
      CALL cellStyle.setBorderRight(borderStyleType.THIN)
      CALL cellStyle.setRightBorderColor(excelColors.BLACK.index)
   END IF

   IF bottomBorder THEN
      CALL cellStyle.setBorderBottom(borderStyleType.THIN)
      CALL cellStyle.setBottomBorderColor(excelColors.BLACK.index)
   END IF

   IF leftBorder THEN
      CALL cellStyle.setBorderLeft(borderStyleType.THIN)
      CALL cellStyle.setLeftBorderColor(excelColors.BLACK.index)
   END IF

END FUNCTION #set_border_style

PUBLIC FUNCTION auto_size_column(s sheetType, colIdx INTEGER) RETURNS ()

   CALL s.autoSizeColumn(colIdx)

END FUNCTION #auto_size_column

PUBLIC FUNCTION freeze_rows(s sheetType, startRow INTEGER, numOfRows INTEGER) RETURNS ()

   CALL s.createFreezePane(startRow, numOfRows)

END FUNCTION #freeze_rows

PUBLIC FUNCTION workbook_createsheet_with_name(w workbookType, sheetName STRING) RETURNS (sheetType)
   DEFINE s sheetType

    LET s= w.createSheet(sheetName)
    RETURN s

END FUNCTION

PUBLIC FUNCTION set_style_format(w workbookType, s cellStyleType, fmt STRING) RETURNS ()
   DEFINE c cellFormat

	LET c = w.createDataFormat()
	CALL s.setDataFormat(c.getFormat(fmt))

END FUNCTION

PUBLIC FUNCTION set_builtin_style_format(s cellStyleType, fmt SMALLINT) RETURNS ()

	CALL s.setDataFormat(fmt)

END FUNCTION

PUBLIC FUNCTION create_style_from_style(w workbookType, s cellStyleType) RETURNS (cellStyleType)
   DEFINE s2 cellStyleType

   LET s2 = style_create(w)
   IF s IS NOT NULL THEN
     CALL s2.cloneStyleFrom(s)
   END IF

   RETURN s2

END FUNCTION

PUBLIC FUNCTION set_margins(s sheetType,
                            headerMargin FLOAT,
                            footerMargin FLOAT,
                            leftMargin FLOAT,
                            rightMargin FLOAT) RETURNS ()

   CALL s.setMargin(sheetType.HeaderMargin, headerMargin)
   CALL s.setMargin(sheetType.FooterMargin, footerMargin)

   CALL s.setMargin(sheetType.LeftMargin, leftMargin)
   CALL s.setMargin(sheetType.RightMargin, rightMargin)

END FUNCTION

PUBLIC FUNCTION set_header_footer(s sheetType,
                                  headerMargin FLOAT,
                                  footerMargin FLOAT) RETURNS ()
   DEFINE ps pageSetup

   LET ps = s.getPrintSetup()
   CALL ps.setHeaderMargin(headerMargin)
   CALL ps.setFooterMargin(footerMargin)

END FUNCTION #set_header_footer

PUBLIC FUNCTION set_sheet_header(s sheetType, 
                                 leftTitle STRING,
                                 centerTitle STRING,
                                 rightTitle STRING) RETURNS ()
   DEFINE sheetHeader headerType

   LET sheetHeader = s.getHeader()
   IF leftTitle IS NOT NULL THEN
      CALL sheetHeader.setLeft(leftTitle);
   END IF
   IF centerTitle IS NOT NULL THEN
      CALL sheetHeader.setCenter(centerTitle);
   END IF
   IF rightTitle IS NOT NULL THEN
      CALL sheetHeader.setRight(rightTitle);
   END IF

END FUNCTION #set_sheet_header

PUBLIC FUNCTION set_sheet_footer(s sheetType, footerText STRING) RETURNS ()
   DEFINE sheetFooter footerType
   CONSTANT cPageFooter = "Page &P of &N"

   LET sheetFooter = s.getFooter()
   CALL sheetFooter.setLeft(footerText)
   CALL sheetFooter.setRight(cPageFooter)

END FUNCTION #set_sheet_footer

PUBLIC FUNCTION set_print_settings(s sheetType) RETURNS ()
   DEFINE ps pageSetup

   LET ps = s.getPrintSetup()
   #CALL s.setFitToPage(TRUE)
   CALL s.setAutobreaks(TRUE)
   CALL s.setPrintGridlines(TRUE)
   CALL s.setRepeatingRows(cellRangeAddressType.valueOf("1:1"))
   CALL ps.setLandscape(TRUE)
   CALL ps.setPageOrder(pageOrderType.DOWN_THEN_OVER)
   CALL ps.setHeaderMargin(0.3)
   CALL ps.setFooterMargin(0.3)

END FUNCTION #set_print_settings
