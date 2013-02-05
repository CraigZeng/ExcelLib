package com.supertool.excel.demo;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.List;

import com.supertool.excel.ExcelFactory;
import com.supertool.excel.ExcelOperation;
import com.supertool.excel.SheetModel;

public class Demo {
	public static void main(String[] args){
		/* Read data from Excel */
		ExcelOperation excelOp = ExcelFactory.openFile("adplacement.xlsx","adplacement.properties",ExcelFactory.EXCEL_LEVEL_NEW);
		List<SheetModel> sheets = excelOp.getFromExcel();
		System.out.println(sheets);
		excelOp.closeFile();
		
		/*Write data to Excel*/
		excelOp =  ExcelFactory.openFile("download.xlsx","adplacement.properties",ExcelFactory.EXCEL_LEVEL_NEW);
		excelOp.saveToExcel(sheets);
		
		try {
			excelOp.writeToStream(new FileOutputStream("/home/supertool/Desktop/download.xlsx"));
			excelOp.closeFile();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}
