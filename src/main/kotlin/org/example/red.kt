package org.example

import org.apache.poi.EncryptedDocumentException
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
                row.getCell(2)


                excelSH.getRow(count).getCell(0).stringCellValue.toString().uppercase()
                excelSH.getRow(count).getCell(1).stringCellValue.toString().uppercase()
                count += 1
            }
            val firstSplit = excelSH.getRow(count).getCell(0).stringCellValue.toString().uppercase().split(","," ", "-", ". ", " .")
            val  secondSplit = excelSH.getRow(count).getCell(0).stringCellValue.toString().uppercase().split(","," ", "-",". ", " .")


            val cellThree = excelSH.getRow(count)


            var counter = 0
            var k = 0
            while (k < cellThree.lastCellNum) {
                for (i in 0..firstSplit.lastIndex) {
                    counter += 1

                    if (firstSplit == secondSplit) {
                        excelSH.getRow(1).createCell(2).setCellValue("match")


                         if (firstSplit[i] == secondSplit[i]) {
                            excelSH.getRow(1).getCell(2).setCellValue(" match by $counter")
                        }
                    }
                    else
                        excelSH.getRow(1).getCell(2).setCellValue(" not match")
                }

                k++
            }
                inputStream.close()


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
