package org.example

import org.apache.poi.EncryptedDocumentException
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

fun main() {
        val path = "./name_match.xlsx"
        try {
            val inputStream = FileInputStream(path)
            val excelWK = WorkbookFactory.create(inputStream)
            val excelSH = excelWK.getSheetAt(0)

            var count = 1
            while (count < excelSH.lastRowNum) {

                val row = excelSH.getRow(count)
                val cellOne = row.getCell(0)
                val cellTwo = row.getCell(1)

                when (cellOne.cellType) {
                    CellType.STRING -> cellOne.stringCellValue
                    else -> {}
                }
                when (cellTwo.cellType) {
                    CellType.STRING -> cellTwo.stringCellValue
                    else -> {}
                }

                val firstString = cellOne.stringCellValue.toString().uppercase()
                val secondString = cellTwo.stringCellValue.toString().uppercase()
                var counter = 0

                val firstSplit = firstString.split(","," ", "-", ". ", " .")
                val secondSplit = secondString.split(","," ", "-",". ", " .")

                val createRow = excelSH.createRow(1)
                createRow.createCell(2)

                var i = 0
                while (i < firstSplit.lastIndex) {
                    counter += 1

                    if (firstString == secondString) {
                        excelSH.getRow(1).getCell(2).setCellValue("match")
//                        println("match")


                         if (firstSplit[i] == secondSplit[i]) {
                            excelSH.getRow(1).getCell(2).setCellValue(" match by $counter")
//                        println("match by $counter")
                        }
                    }
                    else
                        excelSH.getRow(1).getCell(2).setCellValue(" not match")
//                            println("not match")

                    i++
                }
                inputStream.close()
                count += 1
            }

            println("Code is correctly written")

            val outputStream = FileOutputStream("./name_match.xls")
            excelWK.write(outputStream)
            excelWK.close()
            outputStream.close()

        } catch (ex: IOException) {
            ex.printStackTrace()
        } catch (ex: EncryptedDocumentException) {
            ex.printStackTrace()
        }
    }
