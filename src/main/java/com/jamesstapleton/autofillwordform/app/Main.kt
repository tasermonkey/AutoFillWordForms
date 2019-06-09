package com.jamesstapleton.autofillwordform.app

import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.xmlbeans.SimpleValue
import java.io.File
import java.io.FileInputStream
import java.lang.Exception
import java.util.*

val FIELD_TYPE_XPATH = "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:fldChar/@w:fldCharType"
val FIELD_TYPE_VALUE = "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:ffData/w:name/@w:val"

fun main(args: Array<String>) {
    if (args.size < 2) {
        throw Exception("Require at least two arguments: wordDocumentFile settingsFile1 [settingsFile2 [...]]")
    }

    val documentFilename = File(args[0]).absoluteFile
    val outputFilename = File(documentFilename.parentFile, "${documentFilename.nameWithoutExtension}-filledIn.${documentFilename.extension}")
    val fillData = getFillData(args.slice(1 until args.size))

    println("Document FileName: $documentFilename\nFill Data: $fillData")

    fillForm(documentFilename, outputFilename, fillData)
}

fun getFillData(settingsFilenames: Collection<String>): Properties {
    return settingsFilenames.fold(Properties()) { acc, filename ->
        FileInputStream(filename).use { inStream ->
            acc.load(inStream)
            acc
        }
    }
}

fun fillForm(documentFilename: File, outputFilename: File, fillData: Properties) {
    XWPFDocument(documentFilename.inputStream()).use { doc ->
        doc.removeProtectionEnforcement()
        fillInDocument(doc, fillData)

        outputFilename.outputStream().use {
            doc.write(it)
        }
    }
}

fun fillInDocument(document: XWPFDocument, fillData: Properties) {
    for (tbl in document.tables) {
        for (row in tbl.rows) {
            for (col in row.tableCells) {
                scanParagraphs(col.paragraphs, fillData)
            }
        }
    }

    scanParagraphs(document.paragraphs, fillData)
}

fun scanParagraphs(
    paragraphs: Collection<XWPFParagraph>,
    fillData: Properties
) {
    // See http://officeopenxml.com/WPfields.php
    // but a 'formField' can exist between multiple paragraphs, and runs, in between a 'begin' and 'end'

    var currentFieldName: String? = null
    var fieldData: String? = null
    var addedValue = false

    ///w:document/w:body/w:tbl/w:tr/w:tc/w:p/w:r/w:fldChar/w:ffData/w:name/@w:val
    ///w:document/w:body/w:p/w:r/w:fldChar/w:ffData/w:name/@w:val

    /// A complex field is made up of:
    ///w:document/w:body/w:tbl/w:tr/w:tc/w:p/w:r/w:fldChar/w:ffData/
    /// // begin
    ///.../tc/other tags(including other p,r's)
    ///.../0 or more of t's (text)
    /// // end

    for (paragraph in paragraphs) {
        for (run in paragraph.runs) {
            val cursor = run.ctr.newCursor()
            cursor.selectPath(FIELD_TYPE_XPATH)
            while (cursor.hasNextSelection()) {
                cursor.toNextSelection()
                if (cursor.`object` !is SimpleValue) {
                    println("Object not SimpleValue: ${cursor.`object`.javaClass}")
                    continue
                }

                val obj = cursor.getObject() as SimpleValue
                if ("begin" == obj.stringValue) {
                    cursor.toParent()
                    val formField = cursor.getObject().selectPath(FIELD_TYPE_VALUE)[0] as SimpleValue

                    currentFieldName = formField.stringValue
                    fieldData = fillData.get(currentFieldName) as String?
                    if (fieldData != null) {
                        println("Starting $currentFieldName")
                    }
                } else if ("end" == obj.stringValue) {
                    if (fieldData != null) {
                        println("Ending $currentFieldName")
                    }

                    currentFieldName = null
                    fieldData = null
                    addedValue = false
                }
            }

            if (currentFieldName != null && fieldData != null && run.ctr.tList.size > 0) {
                // clear out all existing texts
                for(index in (run.ctr.tList.size - 1) downTo 0 ) {
                    run.ctr.removeT(index)
                }

                if (!addedValue) {
                    println("Assigning value: $currentFieldName = $fieldData")
                    run.ctr.addNewT().stringValue = fieldData
                    addedValue = true
                }
            }
        }
    }
}