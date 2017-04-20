package com.dugu.acc.dev.service;

import java.io.IOException;
import java.util.List;

import com.dugu.acc.dev.bean.Employee;
import com.dugu.acc.dev.util.ParseUtils;

/**
 * @author basanta.kumar.hota
 *
 */
public class MultiParseService {

	public void excelToXml(String xlxsFilePath, String FileNameToBeGenerate) {
		List<Employee> employees = null;
		try {
			employees = ParseUtils.excelToJavaBean(xlxsFilePath);
			ParseUtils.convertJavaToXML(employees, FileNameToBeGenerate);
			System.out.println("Generated " + FileNameToBeGenerate
					+ " file from given " + xlxsFilePath
					+ " successfull,kindly  check your Specified path");
		} catch (IOException e) {
			System.out.println(e.getLocalizedMessage());
		}
	}

	public void xmlToExcel(String generateExcellFilePath, String xmlFilePath) {
		List<Employee> employees = null;
		try {
			employees = ParseUtils.convertXMLToJava(xmlFilePath);
			ParseUtils.javaToExcell(generateExcellFilePath, xmlFilePath,
					employees);
			System.out.println("Generated " + generateExcellFilePath
					+ " from given " + xmlFilePath
					+ " successfull,kindly  check your Specified path");
		} catch (IOException e) {
			System.out.println(e.getLocalizedMessage());
		}
	}
}
