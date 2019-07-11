package com.excel.util.excel.part;

import java.util.List;
public class SpecialSheet implements ISheet{
	
	private String sheetName;
	
	private List<Row> rows;
	
	public SpecialSheet(){
		
	}
	
	public SpecialSheet(String sheetName){
		this.setSheetName(sheetName);
	}
	
	public List<Row> getRows() {
		return rows;
	}

	public void setRows(List<Row> rows) {
		this.rows = rows;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

}
