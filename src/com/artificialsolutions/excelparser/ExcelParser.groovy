package com.artificialsolutions.excelparser

import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

class ExcelParser {
    static parse(String filename, boolean useRowsHeaders, boolean useColumnHeaders, boolean markupFormatting) {

        def wbData = [:]
        def file = FileGetter.getFile(filename)
        InputStream inp = new FileInputStream(file)
        Workbook wb = WorkbookFactory.create(inp);
        boolean isOldExcel = wb.getSpreadsheetVersion().ordinal() == 0
        if (isOldExcel && markupFormatting) {
            println('WARNING - Markup formatting not supported for *.xls files. Please save your file as *.xslx to apply markup formatting.')
        }
        wb.forEach({ sheet ->


            def sheetData = [:]
            sheet.forEach({ row ->
                {
                    def rowName;

                    if (useRowsHeaders) {
                        rowName = row.getCell(0).stringCellValue
                    } else {
                        rowName = "row" + (row.rowNum + 1)
                    }

                    def rowData = [:]
                    row.forEach({ cell ->
                        {
                            def columnName;

                            if (useColumnHeaders) {
                                columnName = sheet.getRow(0).getCell(cell.columnIndex).stringCellValue
                            } else {
                                columnName = "column" + cell.columnIndex
                            }

                            def cellData = ""
                            if (markupFormatting) {
                                def richText = cell.getRichStringCellValue()
                                int formattingRuns = richText.numFormattingRuns()
                                def plainText = richText.getString().replaceAll(/\n/, "<br/>")

                                if (formattingRuns > 0 && !isOldExcel) {
                                    for (int i = 0; i < formattingRuns; i++) {
                                        int startIdx = richText.getIndexOfFormattingRun(i)
                                        int length = richText.getLengthOfFormattingRun(i)
                                        int endIdx = startIdx + length

                                        def tranche = plainText.substring(startIdx, endIdx)
                                        Font font = richText.getFontOfFormattingRun(i)


                                        if (font.getBold()) {
                                            tranche = "<b>" + tranche + "</b>"
                                        }
                                        if (font.getItalic()) {
                                            tranche = "<i>" + tranche + "</i>"
                                        }
                                        if (font.getUnderline() > Font.U_NONE) {
                                            tranche = "<u>" + tranche + "</u>"
                                        }
                                        if (font.getStrikeout()) {
                                            tranche = "<strike>" + tranche + "</strike>"
                                        }
                                        if (font.getTypeOffset() == Font.SS_SUPER) {
                                            tranche = "<sup>" + tranche + "</sup>"
                                        }
                                        if (font.getTypeOffset() == Font.SS_SUB) {
                                            tranche = "<sub>" + tranche + "</sub>"
                                        }

                                        def color = font.getXSSFColor().getARGBHex().substring(2)
                                        if (color != "000000") {
                                            tranche = "<span style=\"color: #" + color + ";\">" + tranche + "</span>"
                                        }

                                        cellData += tranche
                                    }
                                } else {
                                    cellData = plainText
                                }

                            } else {
                                cellData = cell.stringCellValue
                            }


                            if (cell.columnIndex > 0 && useColumnHeaders || !useColumnHeaders) {
                                rowData.put(columnName, cellData)
                            }
                        }
                    })
                    if (row.rowNum > 0 && useRowsHeaders || !useRowsHeaders) {
                        sheetData.put(rowName, rowData)
                    }
                }
            })
            wbData.put(sheet.sheetName, sheetData)

        })
        return wbData
    }


}
