package org.example

import org.apache.poi.EncryptedDocumentException
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

fun main () {
    val path1 = "./name_match_2019.xlsx"
    val path2 = "./2019-admitted-sample.xlsx"


    try {
        val inputStream = FileInputStream(path1)
        val excelWB = WorkbookFactory.create(inputStream)

        val inputStram_2 = FileInputStream(path2)
        val excelWB_2 = WorkbookFactory.create(inputStram_2)

        val excelSH = excelWB.getSheetAt(0)
        val c = excelWB_2.getSheetAt(0)


        for (count in 1.. excelSH.lastRowNum) {
            val nameReaderOne = excelSH.getRow(count).getCell(5).stringCellValue.split(" ", "-", "//s")
            val nameReaderTwo = c.getRow(count).getCell(5).stringCellValue.split(" ", "-", "//s")
            val nameResult = excelSH.getRow(count).createCell(7)

//            val dateReaderOne = excelSH.getRow(count).getCell(4).numericCellValue
//            val dateReaderTwo = c.getRow(count).getCell(4).numericCellValue
//            val dateResult = excelSH.getRow(count).createCell(8)

            var flag = 1
            var i = 0

            while (i < nameReaderOne.lastIndex) {
                if (nameReaderOne[i] in nameReaderTwo) {
                    flag += 1
                }
                i++
            }



            when(true) {
                nameReaderOne.containsAll(nameReaderTwo) -> nameResult.setCellValue("match")
                nameReaderOne[i] in nameReaderTwo -> nameResult.setCellValue("match by $flag")
                else -> nameResult.setCellValue("not match")
            }
//            when(true) {
//                dateReaderOne == dateReaderTwo -> dateResult.setCellValue("check date")
//                else -> dateResult.setCellValue("")
//            }

            excelSH.getRow(count).createCell(6).setCellValue(c.getRow(count).getCell(5).stringCellValue)
        }

        println("code written successfully ")
        inputStream.close()
        inputStram_2.close()


        val outputStream = FileOutputStream("./resultResult.xlsx")
        excelWB.write(outputStream)
        excelWB.close()
        outputStream.close()

    } catch (ex: IOException) {
        ex.printStackTrace()
    } catch (ex: EncryptedDocumentException) {
        ex.printStackTrace()
    }
}