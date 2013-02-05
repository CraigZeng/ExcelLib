package com.supertool.excel;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelOperation {
  public void initExcel();
  public void setXWK(Workbook xwk);		
  public void setExcelFile(InputStream is);
  public void setPropertyFile(InputStream is);
  public boolean closeFile();
  public boolean saveToExcel(List<SheetModel> excelData);
  public List<SheetModel> getFromExcel();
  public void writeToStream(OutputStream os);
  public void deleteColumn(int rowNum,int columnNum);
}
