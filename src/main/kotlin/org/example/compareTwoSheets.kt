package org.example

import org.apache.poi.EncryptedDocumentException
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

fun main () {
    val path1 = "./testFile/name_match_2019.xlsx"
    val path2 = "./testFile/2019-admitted-sample.xlsx"


    try {
        val inputStream = FileInputStream(path1)
        val excelWB = WorkbookFactory.create(inputStream)

        val inputString2 = FileInputStream(path2)
        val excelWb2 = WorkbookFactory.create(inputString2)

        val excelSH = excelWB.getSheetAt(0)
        val c = excelWb2.getSheetAt(0)


        for (count in 1.. excelSH.lastRowNum) {
            val nameReaderOne = excelSH.getRow(count).getCell(5).stringCellValue.split(" ", "-", "//s")
            val nameReaderTwo = c.getRow(count).getCell(5).stringCellValue.split(" ", "-", "//s")
            val nameResult = excelSH.getRow(count).createCell(7)


            var flag = 1
            var i = 0

            while (i < nameReaderOne.lastIndex) {
//                if (nameReaderOne[i] in nameReaderTwo) {
//                    flag += 1
//                }
                if (nameReaderOne[i] !in nameReaderTwo) {
                    flag += 1
                }
                i++
            }

//            if (excelSH.getRow(count).getCell(5).stringCellValue[i] == c.getRow(count).getCell(5).stringCellValue[i]) {
//                println("${count + 1}: match by $flag")
//            }


            when(true) {
                nameReaderOne.containsAll(nameReaderTwo) -> nameResult.setCellValue("match")
                nameReaderOne[i] in nameReaderTwo[i] -> nameResult.setCellValue("not match by ${flag+1}")
                nameReaderOne[i] != nameReaderTwo[i] -> nameResult.setCellValue("not match by $flag")
                else -> nameResult.setCellValue("empty")
            }

            excelSH.getRow(count).createCell(6).setCellValue(c.getRow(count).getCell(5).stringCellValue)
        }

        println("code written successfully ")
        inputStream.close()
        inputString2.close()


        val outputStream = FileOutputStream("./T.xlsx")
        excelWB.write(outputStream)
        excelWB.close()
        outputStream.close()

    } catch (ex: IOException) {
        ex.printStackTrace()
    } catch (ex: EncryptedDocumentException) {
        ex.printStackTrace()
    }
}