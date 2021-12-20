package org.example

import org.apache.poi.EncryptedDocumentException
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

fun main () {
    val path = "./name_match.xlsx"

    try {
        val inputStream = FileInputStream(path)
        val excelWB = WorkbookFactory.create(inputStream)
        val excelSH = excelWB.getSheetAt(0)

        for (count in 1.. excelSH.lastRowNum) {
            val readerOne = excelSH.getRow(count).getCell(0).stringCellValue.split(" ", "-", "//s")
            val readerSecond = excelSH.getRow(count).getCell(1).stringCellValue.split(" ", "-", "//s")
            val readerThree = excelSH.getRow(count).createCell(2)

            var flag = 1
            var i = 0

            while (i < readerOne.lastIndex) {
                if (readerOne[i] in readerSecond) {
                    flag += 1
                }
                i++
            }

            when(true) {
                readerOne.containsAll(readerSecond) -> readerThree.setCellValue("match")
                readerOne[i] in readerSecond -> readerThree.setCellValue("match by $flag")
                else -> readerThree.setCellValue("not match")
            }
        }

        println("code written successfully ")
        inputStream.close()

        val outputStream = FileOutputStream(path)
        excelWB.write(outputStream)
        excelWB.close()
        outputStream.close()

    } catch (ex: IOException) {
        ex.printStackTrace()
    } catch (ex: EncryptedDocumentException) {
        ex.printStackTrace()
    }
}