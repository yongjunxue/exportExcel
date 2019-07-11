package com.excel.util.excel.part;

import java.util.ArrayList;
import java.util.List;

public class Row {
	private List<Cell> cells=new ArrayList<>();
	
	private Style style;
	
	public Row(){
	}
	public List<Cell> getCells() {
		return cells;
	}
	public Row setCells(List<Cell> cells) {
		this.cells = cells;
		return this;
	}
	
	public Cell createCell(){
		Cell c=new Cell();
		cells.add(c);
		return c;
	}
	public Cell createCell(String cellValue){
		Cell c=new Cell(cellValue);
		cells.add(c);
		return c;
	}
	public Style getStyle() {
		return style;
	}
	public Row setStyle(Style style) {
		this.style = style;
		return this;
	}
}
