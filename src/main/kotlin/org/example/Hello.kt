package org.example

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream

fun main() {

    val path = "./name_match.xlsx"

    val inputStream = FileInputStream(path)
    val excelWK = XSSFWorkbook(inputStream)
    val excelSH = excelWK.getSheetAt(0)


    for (count in 0..excelSH.lastRowNum) {

        val row = excelSH.getRow(count)
        val cellOne = row.getCell(0)
        val cellTwo = row.getCell(1)

        val results = if (cellOne.cellType.toString() == cellTwo.cellType.toString()) {

                        println("${cellOne.stringCellValue}    ${cellTwo.stringCellValue}   =  Matched")

                    } else {
                        println("Not Matched")
                    }


//        compareTwoCell(firstCell, secondCell)


    }


    inputStream.close()
}

//fun compareTwoCell(cell1: XSSFCell, cell2: XSSFCell ): Boolean {
//    if ((cell1 == null) && (cell2 == null)) {
//        return true
//    } else if ((cell1 == null) || (cell2 == null)) {
//        return false
//    }
//
//    val t1 = cell1.cellType
//    val t2 = cell2.cellType
//
//    if (t1 == t2) {
//
//        var equalCell = false
//
//
//        if (cell1.cellStyle == cell2.cellStyle) {
//
//            when (cell1.cellType) {
//                CellType.STRING -> {
//                    if (cell1.cellStyle == cell2.cellStyle)
//                        equalCell = true
//                }
//                CellType.BOOLEAN -> {
//                    if (cell1.booleanCellValue == cell2.booleanCellValue)
//                        equalCell = true
//                }
//                CellType.NUMERIC -> {
//                    if (cell1.numericCellValue == cell2.numericCellValue)
//                        equalCell = true
//                }
//                else -> {
//                    println("Nothing found")
//                }
//            }
//        } else
//            return false
//
//        return equalCell
//    }
//    return true
//}


