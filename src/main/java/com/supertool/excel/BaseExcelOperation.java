package com.supertool.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class BaseExcelOperation implements ExcelOperation {

	private static Logger logger = Logger.getLogger(BaseExcelOperation.class);
    private Workbook xwk=null;
    private InputStream excelStream=null;
    private InputStream propertyStream=null;
    private List<ExCell> head=null;
    private Map<String, ExCell> body=null;
    private List<String> fields=null;
    private int sheetNumbers=0;
    private SheetOperation sheetOp=null;
    private PropertyManager propManager=null;
    
    private void setFields(List<String> fields){
    	this.fields=fields;
    	this.sheetOp.setFields(this.fields);
    }
    
    public void setXWK(Workbook xwk){
    	this.xwk = xwk;
    }
	
	@Override
	public boolean closeFile() {
		try {
			if(this.excelStream!=null){
			   this.excelStream.close();
			}
			if(this.propertyStream!=null){
			   this.propertyStream.close();
			}
			return true;
		} catch (IOException e) {
			logger.error("Close file error!\n");
			return false;
		}
	}
	
	private void setHead(List<ExCell> head) {
		this.head=head;
		this.sheetOp.setHead(this.head);
	}

	private boolean setBody(Map<String, ExCell> body) {
		this.body=body;
		this.sheetOp.setBody(this.body);
		return false;
	}

	@Override
	public boolean saveToExcel(List<SheetModel> excelData) {
		if(this.sheetNumbers==-1){
			for(int j=1;j<excelData.size();j++){
				this.xwk.cloneSheet(0);
			}
		}
		for(int i=0;i<excelData.size();i++){
		   this.sheetOp.setSheet(this.xwk.getSheetAt(i));
		   if(this.head!=null){
		      this.sheetOp.setHead(this.head);
		      this.sheetOp.saveHeadToExcel(excelData.get(i).getHead());
		   }
		   this.sheetOp.setBody(this.body);
		   this.xwk.setSheetName(i,excelData.get(i).getSheetname());	  
		   this.sheetOp.saveBodyToExcel(excelData.get(i).getBody());
		}
		return true;
	}

	@Override
	public List<SheetModel> getFromExcel() {
		SheetModel sheetDataModel=null;
		List<SheetModel> sheets=new ArrayList<SheetModel>();
		Sheet sheet=null;
		if(sheetNumbers==-1) this.sheetNumbers=this.xwk.getNumberOfSheets();
		for(int i=0;i<this.sheetNumbers;i++){
			sheetDataModel=new SheetModel();
			sheet=this.xwk.getSheetAt(i);
		    this.sheetOp.setSheet(sheet);
		    sheetDataModel.setBody(this.sheetOp.getBodyFromExcel());
			sheetDataModel.setHead(this.sheetOp.getHeadFromExcel());
			sheetDataModel.setSheetname(sheet.getSheetName());
			sheets.add(sheetDataModel);
		}
		return sheets;
	}
    
	
	private void setSheetNumbers(int sheetNumbers) {
		this.sheetNumbers=sheetNumbers;
	} 
	
	
	public void writeToStream(OutputStream os){
		try {
			this.xwk.write(os);
			os.close();
		} catch (IOException e) {
			logger.error("Write back the excel error!");
		}
		
	}

	@Override
	public void deleteColumn(int rowNum, int columnNum) {
	   int sheetNumbers = this.xwk.getNumberOfSheets();
	   for(int i=0;i<sheetNumbers;i++){
		   this.sheetOp.setSheet(this.xwk.getSheetAt(i));
		   this.sheetOp.deleteColumn(rowNum, columnNum);
	   }
	}

	@Override
	public void setExcelFile(InputStream is) {
		this.excelStream = is;
	}


	@Override
	public void setPropertyFile(InputStream is) {
		this.propertyStream = is;
	}
	
	@Override
	public void initExcel(){
		this.sheetOp = new BaseSheetOperation();
		this.propManager = new PropertyManager();
		this.propManager.openFile(this.propertyStream);
		this.setFields(this.propManager.getKeys().get("area1"));
	    this.setHead(this.propManager.getCells());
	    this.setBody(this.propManager.getData());
	    this.setSheetNumbers(this.propManager.getSheetNumbers());
	}
}
