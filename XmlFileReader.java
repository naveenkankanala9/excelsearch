package com.fedex.excelsearch;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XmlFileReader {

	public static List<String> fileReader() {
		List<String> listXmlData = new ArrayList<String>();
		try {
			FileInputStream inputSteam = new FileInputStream(
					new File("/home/naveenkankanala/ide/sts-bundle/workspace/excelsearch/src/main/resources/GoogleSEarch.xlsx"));
			XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(inputSteam);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				if (row.getRowNum() != 0)
				{	
					Iterator<Cell> cellIterator = row.cellIterator();
					Cell cell = (Cell) cellIterator.next();
					listXmlData.add(cell.getStringCellValue());					
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println(listXmlData);
		return listXmlData;
	}
}
