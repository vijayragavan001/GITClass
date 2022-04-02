package com.hexa;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Launch {
	
	public static void main(String[] args) throws IOException {
	File excelLoc = new File("C:\\Users\\vijay\\eclipse-workspace\\MavenFirst\\Exceldata\\expense.xlsx");
	
	FileInputStream fIn = new FileInputStream(excelLoc);
	Workbook w = new XSSFWorkbook(fIn);
	
	Sheet s = w.getSheet("Sheet1");
	
	
	int rows = s.getPhysicalNumberOfRows();
	System.out.println(rows);

	for (int i = 0; i < rows; i++) {
		Row noR = s.getRow(i);
		for (int j = 0; j < noR.getPhysicalNumberOfCells(); j++) {
			Cell noC = noR.getCell(j);
			System.out.print(noC+"  ");
			
		}
	System.out.println();
	}
	
	}
private void add() {
	// TODO Auto-generated method stub

}
}
