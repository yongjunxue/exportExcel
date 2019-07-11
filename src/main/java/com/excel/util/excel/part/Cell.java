package com.excel.util.excel.part;


/**
 * 目前
 * @author xueyongjun
 *
 */
public class Cell {
	private String value;
	
	private Style style;
	
//	private int mergeCount;
	private boolean merge=false;
	
//	private int firstRow;
//	private int lastRow;
	private int firstCol;
	private int lastCol;
	
	private int width;
//	private int height;
	
	public Cell(){
		
	}
	public Cell(String value){
		this.setValue(value);
	}
	public String getValue() {
		return value;
	}
	public Cell setValue(String value) {
		this.value = value;
		return this;
	}
	public Style getStyle() {
		return style;
	}
	public Cell setStyle(Style style) {
		this.style = style;
		return this;
	}
//	public Cell setMergeRange(int firstRow,int lastRow,int firstCol,int lastCol){
	public Cell setMergeRange(int firstCol,int lastCol){
		if(
//				lastRow<firstRow || firstRow<0 || 
				lastCol<firstCol || firstCol<0){
			throw new RuntimeException("合并单元格参数错误");
		}
//		this.firstRow=firstRow;
//		this.lastRow=lastRow;
		this.firstCol=firstCol;
		this.lastCol=lastCol;
		merge=true;
		return this;
	}
	public Cell setWidth(int width){
		this.width=width;
//		this.height=height;
		return this;
	}
//	public int getFirstRow() {
//		return firstRow;
//	}
//	public int getLastRow() {
//		return lastRow;
//	}
	public int getFirstCol() {
		return firstCol;
	}
	public int getLastCol() {
		return lastCol;
	}
	public boolean isMerge() {
		return merge;
	}
	public int getWidth() {
		return width;
	}
//	public int getHeight() {
//		return height;
//	}
//	public int getMergeCount() {
//		return mergeCount;
//	}
//	public Cell setMergeCount(int mergeCount) {
//		this.mergeCount = mergeCount;
//		return this;
//	}
}
