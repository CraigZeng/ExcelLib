 package test.com.excel;
import java.io.InputStream;

import junit.framework.Assert;

import org.junit.Test;

import com.supertool.excel.ExcelFactory;

public class TestExcelFactory{	
	@Test 
	public void testOpenFileWithNullStreamAndString(){
		Assert.assertNull(ExcelFactory.openFile(null,null));
	}
	
	@Test
	public void testOpenFileWithNullStream(){
		InputStream excelStream = null;
		InputStream propStream = null;
		Assert.assertNull(ExcelFactory.openFile(excelStream,propStream,null));
	}
	
	@Test
	public void testOpenFileWithWrongPropFile(){
		InputStream excelStream = null;
		String propFile = "test.html"; 
		Assert.assertNull(ExcelFactory.openFile(excelStream,propFile,ExcelFactory.EXCEL_LEVEL_NEW));
	}
	
	@Test
	public void testOpenFileWithWrongExcelFile(){
		String propFile = null;
		Assert.assertNull(ExcelFactory.openFile("test.html",propFile,null));
	}
	
	@Test
	public void testOpenFileWithWrongPorpButRightExcel(){
		Assert.assertNull(ExcelFactory.openFile("download.xlsx","test.html",null));
	}
	
	@Test
	public void testOpenFileWithWrongExcelButRightProp(){
		Assert.assertNull(ExcelFactory.openFile("test.html","adplacement.properties",null));
	}
}
