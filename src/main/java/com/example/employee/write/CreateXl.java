package com.example.employee.write;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.entity.Employee;

public class CreateXl {
	public static void createXl(Map<Integer, Object> data) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Employee data");
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("ID");
		cell = row.createCell(1);
		cell.setCellValue("Name");
		cell = row.createCell(2);
		cell.setCellValue("Location");
		int rownum = 1;
		for (Integer key : data.keySet()) {
			Row r = sheet.createRow(rownum++);
			Employee e = (Employee) data.get(key);
			int cellnum = 0;
			Cell c = r.createCell(cellnum++);
			c.setCellValue(e.getId());
			c = r.createCell(cellnum++);
			c.setCellValue(e.getName());
			c = r.createCell(cellnum++);
			c.setCellValue(e.getLocation());
		}
		FileOutputStream f = null;
		try {
			f = new FileOutputStream("WriteExcel.xlsx");
			workbook.write(f);
			f.close();
			System.out.println("=======successfull=======");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				f.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	}

}
