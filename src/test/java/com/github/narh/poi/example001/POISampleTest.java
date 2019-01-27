/*
 * Copyright (c) 2018, NARH https://github.com/NARH
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 * * Redistributions of source code must retain the above copyright notice,
 *   this list of conditions and the following disclaimer.
 * * Redistributions in binary form must reproduce the above copyright notice,
 *   this list of conditions and the following disclaimer in the documentation
 *   and/or other materials provided with the distribution.
 * * Neither the name of the copyright holder nor the names of its contributors
 *   may be used to endorse or promote products derived from this software
 *   without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
 * ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

package com.github.narh.poi.example001;

import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import lombok.extern.slf4j.Slf4j;

/**
 * @author narita
 *
 */
@Slf4j
public class POISampleTest {

  public static final String TEST_XLSX_FILE_NAME = "sample001.xlsx";
  public static final String TEST_XLSX_SHEET_NAME = "Sheet1";

  @Test
  public void testXLSXRead() throws Exception {
    if(log.isInfoEnabled()) log.info("start testXLSXRead");

    //テンプレート XLSX の取得
    URL filePath = getClass().getResource(File.separator + TEST_XLSX_FILE_NAME);
    if(log.isTraceEnabled()) log.trace("target origin file path:" + filePath);

    File file = new File(filePath.toURI());
    if(null != file && file.exists()) {
      // パッケージを開く
      OPCPackage pkg = OPCPackage.open(file);
      // テンプレート Book を開く
      XSSFWorkbook origin = (XSSFWorkbook) WorkbookFactory.create(pkg);
      if(null == origin) {
        if(log.isWarnEnabled()) log.warn("book {} is null.", file.toString());
      }

      XSSFSheet originSheet = origin.getSheet(TEST_XLSX_SHEET_NAME);
      XSSFRow[] headers = new XSSFRow[2];
      headers[0] = originSheet.getRow(0);
      headers[1] = originSheet.getRow(1);

      XSSFRow[] footers = new XSSFRow[1];
      footers[0] = originSheet.getRow(originSheet.getLastRowNum());

      // 作業 XLSX として開く
      SXSSFWorkbook workbook = new SXSSFWorkbook();

      // 作業 Book を閉じる
      //workbook.close();

      // 作業シートを選択する
      SXSSFSheet sheet = workbook.createSheet(TEST_XLSX_SHEET_NAME);
      if(null == sheet) {
        if(log.isWarnEnabled()) log.warn("Book {}, Sheet {} is null.", file.toString(), TEST_XLSX_SHEET_NAME);
      }

      SXSSFRow row = null;
      SXSSFCell cell = null;
      XSSFCellStyle style = null;

      // ヘッダー行を作成する
      row = sheet.createRow(0);
      row.setRowStyle(headers[0].getRowStyle());
      for(int i = 0; i < headers[0].getLastCellNum(); i++) {
        if(null != headers[0].getCell(i)) {
          cell = row.createCell(i);
          if(null != headers[0].getCell(i).getCellStyle())
            style = headers[0].getCell(i).getCellStyle();
          cell.setCellStyle(style);

          cell.setCellType(headers[0].getCell(i).getCellType());

          switch(headers[0].getCell(i).getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:
              cell.setCellValue(headers[0].getCell(i).getNumericCellValue());
              break;
            case Cell.CELL_TYPE_STRING:
              cell.setCellValue(headers[0].getCell(i).getStringCellValue());
              break;
            case Cell.CELL_TYPE_FORMULA:
              cell.setCellFormula(headers[0].getCell(i).getCellFormula());
              cell.setCellValue(headers[0].getCell(i).getDateCellValue());
              break;
            case Cell.CELL_TYPE_BLANK:
              cell.setCellType(Cell.CELL_TYPE_BLANK);
              break;
            case Cell.CELL_TYPE_BOOLEAN:
              cell.setCellValue(headers[0].getCell(i).getBooleanCellValue());
              break;
            case Cell.CELL_TYPE_ERROR:
              cell.setCellErrorValue(headers[0].getCell(i).getErrorCellValue());
              break;
          }
        }
      }

      row = sheet.createRow(1);
      for(int i = 0; i < headers[1].getLastCellNum(); i++) {
        if(null != headers[0].getCell(i)) {
          cell = row.createCell(i);
          if(null != headers[0].getCell(i).getCellStyle())
            style = headers[0].getCell(i).getCellStyle();
          cell.setCellStyle(style);

          cell.setCellType(headers[1].getCell(i).getCellType());
          cell.setCellValue(headers[1].getCell(i).getStringCellValue());
        }
      }

      if(null == row) {
        if(log.isWarnEnabled()) log.warn("Sheet {}, 4 row is null.", TEST_XLSX_SHEET_NAME);
      }

      //ByteArrayOutputStream baos = new ByteArrayOutputStream();
      ZipSecureFile.setMinInflateRatio(0.001);
      File output = Files.createTempFile(Paths.get("/tmp"),"sample001_", ".xlsx").toFile();
      FileOutputStream fos = new FileOutputStream(output);
      try {
        workbook.write(fos);
      }
      finally {
        fos.close();
        if(output.exists()) output.renameTo(new File("/tmp/foo.xlsx"));
        // 作業 Book を閉じる
        workbook.close();
        // テンプレートを閉じる
        origin.close();
      }
    }
    else {
      if(log.isWarnEnabled()) log.warn("file {} not exists.", file.toString());
    }
    if(log.isInfoEnabled()) log.info("end testXLSXRead");
  }
}
