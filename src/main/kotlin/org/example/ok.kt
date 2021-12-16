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
        val wlBk = WorkbookFactory.create(inputStream)
        val wlSh = wlBk.getSheetAt(0)



        for (count in 1.. wlSh.lastRowNum) {
            val readerOne = wlSh.getRow(count).getCell(0).stringCellValue.split(" ")
            val readerSecond = wlSh.getRow(count).getCell(1).stringCellValue.split(" ")
            val readerThree = wlSh.getRow(count).createCell(2)

            if (readerOne == readerSecond) {
                println("match")
                readerThree.setCellValue("match")

            } else {
                println("not match")
                readerThree.setCellValue("not match")
            }



        }

        inputStream.close()

        val outputStream = FileOutputStream(path)
        wlBk.write(outputStream)
        wlBk.close()
        outputStream.close()

    } catch (ex: IOException) {
        ex.printStackTrace()
    } catch (ex: EncryptedDocumentException) {
        ex.printStackTrace()
    }
}

fun excelRead() {

}