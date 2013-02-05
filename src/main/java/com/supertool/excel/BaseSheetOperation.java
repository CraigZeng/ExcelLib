package com.supertool.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
public class BaseSheetOperation implements SheetOperation {

	static Logger logger = Logger.getLogger(BaseSheetOperation.class);
	private Sheet sheet=null;
	private List<ExCell> head=null;
	private Map<String,ExCell> body=null;
	
	private List<String> fields=null;
	private final String START_POINT="startpoint";
	private final String END_POINT="endpoint";
	
	@Override
	public void setSheet(Sheet sheet) {
		this.sheet = sheet;
	}

	@Override
	public void setHead(List<ExCell> head) {
		this.head = head;
	}

	@Override
	public void setBody(Map<String, ExCell> body) {
		this.body = body;
	}

	@Override
	public void setFields(List<String> fields) {
		this.fields = fields;
	}

	@Override
	public boolean saveHeadToExcel(Map<String, Object> head) {
		for(int i=0;i<this.head.size();i++){
			ExCell exCell=this.head.get(i);
			String value=head.get(exCell.getKey().toString()).toString();
            this.sheet.getRow(exCell.getX()).getCell(exCell.getY()).setCellValue(value);			
		}
		return true;
	}

	@Override
	public boolean saveBodyToExcel(List<Map<String, Object>> body) {
		 int startRow=this.body.get(START_POINT).getX()+1;
		 int rowNums = startRow + 1;
		 if(body.size()>(rowNums-startRow)){
			 this.createRows(rowNums-1,(body.size()-(rowNums-startRow)));
		 }
		 rowNums=this.sheet.getPhysicalNumberOfRows();
		 for(int i=startRow,j=0;j<body.size();i++,j++){
		   this.mapToRow(body.get(j),this.sheet.getRow(i));
   	     }
		 return true;
	}

	@Override
	public Map<String, Object> getHeadFromExcel() {
		Map<String, Object> head=new HashMap<String,Object>();
		ExCell headExCell=null;
		for(int i=0;i<this.head.size();i++){
			headExCell = this.head.get(i);
		    head.put(headExCell.getKey(), this.getDataFromExCell(headExCell));
		}
		return head;
	}

	@Override
	public List<Map<String, Object>> getBodyFromExcel() {
		List<Map<String,Object>> bodyData=new ArrayList<Map<String,Object>>();
		ExCell startPoint=this.body.get(START_POINT);
		ExCell endPoint=this.body.get(END_POINT);
		Map<String,Object> rowMap=null;
		if(startPoint.getX()!=endPoint.getX()){
			return null;
		}
		else{
		    int rowStart=startPoint.getX()+1;
		    int rowEnd=this.sheet.getPhysicalNumberOfRows();
		    for(int i=rowStart;i<rowEnd;i++){
				Row row=this.sheet.getRow(i);
				rowMap=this.rowToMap(row);
				if(rowMap!=null){
					bodyData.add(rowMap);
				}
		    }
		}
		return bodyData;
	}

	@Override
	public void mapToRow(Map<String, Object> map, Row row) {
		if(row==null || map == null) return;
		ExCell startPoint=this.body.get(START_POINT);
		ExCell endPoint=this.body.get(END_POINT);
		int cellNumStart=startPoint.getY();
		int cellNumEnd=endPoint.getY();
		for(int i=cellNumStart,j=0;i<=cellNumEnd;i++,j++){
			row.getCell(i).setCellValue(map.get(this.fields.get(j)).toString());
		}
	}

	@Override
	public Map<String, Object> rowToMap(Row row) {
		if(row==null) return null;
		Map<String,Object> rowMap=new HashMap<String,Object>();
		ExCell startPoint=this.body.get(START_POINT);
		ExCell endPoint=this.body.get(END_POINT);
		int cellNumStart=startPoint.getY();
		int cellNumEnd=endPoint.getY();
		boolean isAllNull = true;
		Object obj = null;
		for(int i=cellNumStart,j=0;i<=cellNumEnd;i++,j++){
			obj = this.dataTypeConvert(row.getCell(i));
			rowMap.put(this.fields.get(j),obj);
			if(obj!=null){
				isAllNull = false;
			}	
		}
		if(isAllNull){
			rowMap=null;
		}
		return rowMap;
	}

	@Override
	public void deleteColumn(int rowNum, int columnNum) {
		Row row = null;
		for(int i=rowNum;i<this.sheet.getPhysicalNumberOfRows()+1;i++){
			row = this.sheet.getRow(i);
			row.removeCell(row.getCell(columnNum));
		}
	}
	
	private void createRows(int start,int nums){
		ExCell startPoint=this.body.get(START_POINT);
		ExCell endPoint=this.body.get(END_POINT);
		
		int cellNumStart=startPoint.getY();
		int cellNumEnd=endPoint.getY();
		for(int i=1;i<nums+2;i++){
			Row xRow=this.sheet.createRow(start+i);
			Cell cell = null;
			CellStyle cellUpStyle = null;
			for(int j=cellNumStart;j<=cellNumEnd;j++){
			  cell = xRow.createCell(j,Cell.CELL_TYPE_STRING);
			  cellUpStyle = this.sheet.getRow(start+i-1).getCell(j).getCellStyle();
			  cell.setCellStyle(cellUpStyle);
			}
		}
	}

	private Object getDataFromExCell(ExCell excell){
		int x=excell.getX();
		int y=excell.getY();
		Cell cell=this.sheet.getRow(x).getCell(y);
		return this.dataTypeConvert(cell);
	}
	
	private Object dataTypeConvert(Cell cell){
		Object obj=null;
		if(cell!=null){
			cell.setCellType(Cell.CELL_TYPE_STRING);
			obj = cell.getStringCellValue();
		}
		return obj;
	}
}
