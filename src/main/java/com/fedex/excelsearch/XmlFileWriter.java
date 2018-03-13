package com.fedex.excelsearch;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XmlFileWriter {

	public void fileWriter(List<String> expirationDates) {
		try {
			FileInputStream inputSteam = new FileInputStream(
					new File("/home/naveenkankanala/ide/sts-bundle/workspace/excelsearch/src/main/resources/GoogleSEarch.xlsx"));
			XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(inputSteam);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<String> listIterator = expirationDates.iterator();
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext() && listIterator.hasNext()) {
				Row row = rowIterator.next();
				if (row.getRowNum() != 0)
				{
					row.createCell(1).setCellValue(listIterator.next());
				}
			}
			FileOutputStream outputStream = new FileOutputStream("/home/naveenkankanala/ide/sts-bundle/workspace/excelsearch/src/main/resources/GoogleSEarch.xlsx");
			workbook.write(outputStream);
			outputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
  }
}
	
