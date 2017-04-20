package com.dugu.acc.dev.test;

import com.dugu.acc.dev.service.MultiParseService;

public class ExcelXmlTest {
	private static String FILE_PATH = "src/main/java/com/dugu/acc/dev/generate/files/";

	public static void main(String[] args) {

		MultiParseService service = new MultiParseService();

		service.excelToXml("employee.xlsx", FILE_PATH + "employee.xml");

		service.xmlToExcel(FILE_PATH + "employee.xlsx", FILE_PATH
				+ "employee.xml");
	}
}
