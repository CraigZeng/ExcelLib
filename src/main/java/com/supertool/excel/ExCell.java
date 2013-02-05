package com.supertool.excel;

public class ExCell {
	private Integer x;
	private Integer y;
	private String key;
	private Object value;
	public String getKey() {
		return key;
	}
	public Object getValue() {
		return value;
	}
	public void setKey(String key) {
		this.key = key;
	}
	public void setValue(Object value) {
		this.value = value;
	}
    public Integer getX() {
		return x;
	}
	public Integer getY() {
		return y;
	}
	public void setX(Integer x) {
		this.x = x;
	}
	public void setY(Integer y) {
		this.y = y;
	}
	
}
