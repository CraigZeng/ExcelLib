package com.supertool.excel;
import java.io.InputStream;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFactory {
	private static Logger logger = Logger.getLogger(ExcelFactory.class);
	public final static String EXCEL_LEVEL_OLD = "2003";
	public final static String EXCEL_LEVEL_NEW = "2007";
	public final static String EXCEL_OLD_SUFFIX = ".xls";
	public final static String EXCEL_NEW_SUFFIX = ".xlsx";
	public final static String PROPFILE_SUFFIX = ".properties"; 
	
	public static ExcelOperation openFile(InputStream excelStream,InputStream propStream,final String excelCate){
		Workbook xwk = null;
    	ExcelOperation excelOp = null;
    	if(excelStream !=null && propStream != null){
	    	try{
		        if(excelCate!=null){
		    		if(excelCate.equals(EXCEL_LEVEL_OLD)){
		    		   xwk = new HSSFWorkbook(excelStream);
		    		 }else{
		    		   xwk = new XSSFWorkbook(excelStream);
		    		 }
		    	}else{
		   		     xwk = new XSSFWorkbook(excelStream);
		        }
		        excelOp=new BaseExcelOperation();
		   		excelOp.setXWK(xwk);
		   		excelOp.setPropertyFile(propStream);
		   		excelOp.initExcel();
		        return excelOp;
	    	}catch (Exception e) {
				logger.error("Load file error:\n"+e.toString());
				return null;
			}
    	}
    	return excelOp;
	}
	
	public static ExcelOperation openFile(InputStream excelStream,String propFile){
		return openFile(excelStream, propFile, null);
	}
	
	public static ExcelOperation openFile(InputStream excelStream,String propFile,final String excelCate){
		InputStream propStream = getStreamFromFile(propFile,false);
		return openFile(excelStream, propStream, excelCate);
	}
	
	public static ExcelOperation openFile(String excelFile,InputStream propStream,final String excelCate){
		InputStream excelStream = getStreamFromFile(excelFile,true);
		return openFile(excelStream, propStream, excelCate);
	}

	public static ExcelOperation openFile(String excelFile,String propFile,final String excelCate){
		InputStream excelStream = getStreamFromFile(excelFile,true);
		InputStream propStream = getStreamFromFile(propFile,false);
		return openFile(excelStream, propStream, excelCate);
	}
	
	private static InputStream getStreamFromFile(String filePath,boolean isExcel){
		if((isExcel && validataExcel(filePath)) || (!isExcel && validataProp(filePath))){
			InputStream stream = ExcelFactory.class.getClassLoader().getResourceAsStream(filePath);
			return stream;
		}else{
			return null;
		}
	}
	
	private static boolean validataExcel(String fileName){
		boolean result = false;
		if(fileName != null){
			if(fileName.trim().endsWith(EXCEL_NEW_SUFFIX)||fileName.trim().endsWith(EXCEL_OLD_SUFFIX)){
				result = true;
			}else{
				logger.error("Wrong excel file name!\n");
			}
		}else{
			logger.error("Excel file name cannot be NULL!\n");
		}
		return result;
	}
	
	private static boolean validataProp(String fileName){
		boolean result = false;
		if(fileName != null && fileName.toString().endsWith(PROPFILE_SUFFIX)){
			result = true;
		}else{
			if(fileName == null){
				logger.error("Properties file cannot be NULL!\n");
			}else{
				logger.error("Wrong properties file name!\n");
			}
		}
		return result;
	}
}
