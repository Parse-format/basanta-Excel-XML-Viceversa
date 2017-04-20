package com.dugu.acc.dev.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Marshaller;
import javax.xml.bind.Unmarshaller;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dugu.acc.dev.bean.Employee;
import com.dugu.acc.dev.bean.EmployeeList;

public class ParseUtils {

	@SuppressWarnings("resource")
	public static List<Employee> excelToJavaBean(String xlxsFilePath)
			throws IOException {
		List<Employee> employees = new ArrayList<Employee>();
		Employee employee = null;
		XSSFWorkbook workBook = new XSSFWorkbook(xlxsFilePath);
		XSSFSheet sheet = workBook.getSheetAt(0);
		int totalRows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total Row find :" + totalRows);
		Row row;
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			row = (Row) sheet.getRow(i);

			int id = (int) row.getCell(0).getNumericCellValue();
			String name = row.getCell(1).getStringCellValue();
			String dept = row.getCell(2).getStringCellValue();
			double salary = row.getCell(3).getNumericCellValue();
			employee = new Employee(id, name, dept, salary);
			employees.add(employee);
		}
		return employees;
	}

	public static void convertJavaToXML(List<Employee> employees,
			String FileNameToBeGenerate) {
		EmployeeList employeeList = new EmployeeList();
		employeeList.setEmployees(employees);
		try {
			JAXBContext context = JAXBContext.newInstance(EmployeeList.class);
			Marshaller m = context.createMarshaller();
			m.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, Boolean.TRUE);
			m.marshal(employeeList, new File(FileNameToBeGenerate));
		} catch (JAXBException e) {
			e.printStackTrace();
		}
	}

	public static List<Employee> convertXMLToJava(String xmlFilePath) {
		EmployeeList employeeList = null;
		List<Employee> employees = null;
		try {
			File file = new File(xmlFilePath);
			JAXBContext jaxbContext = JAXBContext
					.newInstance(EmployeeList.class);
			Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
			employeeList = (EmployeeList) jaxbUnmarshaller.unmarshal(file);
			employees = employeeList.getEmployees();
		} catch (JAXBException e) {
			e.printStackTrace();
		}
		return employees;

	}

	@SuppressWarnings("resource")
	public static void javaToExcell(String generateExcellFilePath,
			String xmlFilePath, List<Employee> employees) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Employee List");
		sheet.setDefaultColumnWidth(30);

		// create style for header cells
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontName("Arial");
		style.setFillForegroundColor(HSSFColor.BLUE.index);
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		font.setColor(XSSFFont.COLOR_RED);
		style.setFont(font);
		// cell.setCellStyle(style);
		// create header row
		XSSFRow header = sheet.createRow(0);
		header.setRowStyle(style);
		header.createCell(0).setCellValue("EId");
		header.getCell(0).setCellStyle(style);

		header.createCell(1).setCellValue("Name");
		header.getCell(1).setCellStyle(style);

		header.createCell(2).setCellValue("Dept");
		header.getCell(2).setCellStyle(style);

		header.createCell(3).setCellValue("Salary");
		header.getCell(3).setCellStyle(style);

		int rownum = 0;
		for (Employee employee : employees) {
			XSSFRow aRow = sheet.createRow(rownum++);
			aRow.createCell(0).setCellValue(employee.getId());
			aRow.createCell(1).setCellValue(employee.getName());
			aRow.createCell(2).setCellValue(employee.getDept());
			aRow.createCell(3).setCellValue(employee.getSalary());
		}
		try (FileOutputStream out = new FileOutputStream(new File(
				generateExcellFilePath));) {
			workbook.write(out);
		}
	}

}
