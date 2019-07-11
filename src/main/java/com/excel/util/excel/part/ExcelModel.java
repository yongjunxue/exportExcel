package com.excel.util.excel.part;

import java.util.ArrayList;
import java.util.List;

public class ExcelModel {
	
	private String excelName;
	
	private List<ISheet> sheets=new ArrayList<>();
	
	private int sheetNo;//工作表序号
	
	public ExcelModel(){
	}
	
	public ExcelModel(String excelName){
		this.setExcelName(excelName);
	}
	
	public Sheet createSheet(){
		Sheet s=new Sheet();
		++sheetNo;
		s.setSheetName("sheet"+sheetNo);
		sheets.add(s);
		return s;
	}
	
	public Sheet createSheet(String sheetName){
		Sheet s=new Sheet(sheetName);
		++sheetNo;
		sheets.add(s);
		return s;
	}
	
	public SpecialSheet createSpecialSheet(){
		SpecialSheet s=new SpecialSheet();
		++sheetNo;
		s.setSheetName("sheet"+sheetNo);
		sheets.add(s);
		return s;
	}
	
	public SpecialSheet createSpecialSheet(String sheetName){
		SpecialSheet s=new SpecialSheet(sheetName);
		++sheetNo;
		sheets.add(s);
		return s;
	}
	
	public List<ISheet> getSheets(){
		return sheets;
	}
	
	public String getExcelName() {
		return excelName;
	}

	public ExcelModel setExcelName(String excelName) {
		this.excelName = excelName;
		return this;
	}
}
