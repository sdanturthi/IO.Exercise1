package com.Excel.TestCases;

import java.io.IOException;

import org.testng.annotations.Test;

import com.Excel.Utilities.XLUtils;

public class TestData {	
	
	@Test
	public void TestData_TC001() throws IOException{
		
		String xlPath = System.getProperty("user.dir") + "/src/test/java/com/Excel/TestData/Book1.xlsx";		
		XLUtils util = new XLUtils();	
		
		System.out.println(xlPath);
		
		int lastRowValue = util.getLastRow(xlPath, "Sheet1");
		System.out.println(lastRowValue);	
		
		int lastCellValue = util.getLastCol(xlPath, "Sheet1", 1);
		System.out.println(lastCellValue);
		
	 for(int i=0; i<=lastRowValue; i++) {
			for(int j=0; j<=lastCellValue; j++) {
				String data = util.getCellData(xlPath, "Sheet1", i, j);
				System.out.println(data);
			}
		}   
		
		int updatedRowValue = lastRowValue+1;
		
		util.setCellValue(xlPath, "Sheet1", updatedRowValue, 0, "12");
		util.setCellValue(xlPath, "Sheet1", updatedRowValue, 1, "no body to hire");
	}	
}


