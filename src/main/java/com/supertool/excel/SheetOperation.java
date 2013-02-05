package com.supertool.excel;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public interface SheetOperation {
	  public void setSheet(Sheet sheet);
	  public void setHead(List<ExCell> head);
	  public void setBody(Map<String, ExCell> body);
	  public void setFields(List<String> fields);
	  public boolean saveHeadToExcel(Map<String,Object> head);
	  public boolean saveBodyToExcel(List<Map<String,Object>> body);
	  public Map<String,Object> getHeadFromExcel();
	  public List<Map<String,Object>> getBodyFromExcel();
	  public void mapToRow(Map<String,Object> map,Row row);
	  public Map<String,Object> rowToMap(Row row);
	  public void deleteColumn(int rowNum,int columnNum);
}
