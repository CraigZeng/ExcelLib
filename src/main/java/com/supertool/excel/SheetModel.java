package com.supertool.excel;

import java.util.List;
import java.util.Map;

public class SheetModel {
	private String sheetname=null;
	private Map<String, Object> head=null;
	private List<Map<String, Object>> body=null;
    public String getSheetname() {
		return sheetname;
	}
	public Map<String, Object> getHead() {
		return head;
	}
	public List<Map<String, Object>> getBody() {
		return body;
	}
	public void setSheetname(String sheetname) {
		this.sheetname = sheetname;
	}
	public void setHead(Map<String, Object> head) {
		this.head = head;
	}
	public void setBody(List<Map<String, Object>> body) {
		this.body = body;
	}   
}
