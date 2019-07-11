package com.excel.util.excel.part;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.excel.util.CopyUtil;
@SuppressWarnings("rawtypes")
public class Block {
	private Cell title;
	
//	private List<Row> specialHeader;//特殊表头。一般情况下不会使用
	
	private List<Cell> header;
	private List<String> headerKeys;//和header一一对应

	private List<Row> rows;
	private List<Map<String,Object>> data;
	
	private Cell footer;
	
	private boolean showRowNo=false;//是否显示序号
	private String rowNoName="序号";
	
	
	private Style titleStyle;//这个优先级低于title.getStyle的优先级
//	private Style specialHeaderStyle;
	private Style headerStyle;//这个优先级低于header.getIndex(i).getStyle的优先级
	private Style rowsStyle;//表示这个block下的所有行，这个优先级低于rows.getIndex(i).getStyle的优先级
	private Style footerStyle;//这个优先级低于footer.getStyle的优先级
	
	public Block(){
		this.titleStyle=new Style().setSize((short) 20).setBold(true).setAlign(HorizontalAlignment.CENTER); //默认
		this.headerStyle=new Style().setBold(true).setBackColor(Color.LIGHT_TURQUOISE).setAlign(HorizontalAlignment.CENTER);//默认
		this.rowsStyle=new Style().setAlign(HorizontalAlignment.CENTER);
		this.footerStyle=new Style().setAlign(HorizontalAlignment.CENTER);
	}
	
	public Cell getTitle(){
		return title;
	}
	public Block setTitle(Cell title){
		this.title=title;
		return this;
	}
	public Block setTitle(String titleName){
		title=new Cell(titleName);
		return this;
	}
	
	public List<Cell> getHeader() {
		return header;
	}
	public Block setHeader(String ... headerNames) {
		header=new ArrayList<>();
		for(String headerName : headerNames){
			header.add(new Cell(headerName));
		}
		return this;
	}
	
	public List<String> getHeaderKeys() {
		return headerKeys;
	}

	public Block setHeaderKeys(String ... headerKeys){
		this.headerKeys=new ArrayList<>();
		for(String headerKey : headerKeys){
			this.headerKeys.add(headerKey);
		}
		return this;
	}
	
	public List<Row> getRows() {
		return rows;
	}
	public Block setRows(List<Row> rows){
		this.rows=rows;
		return this;
	}
	
	public List<Map<String,Object>> getData(){
		return data;
	}
	
	public Block setData(List<Map<String,Object>> data){
		this.data=data; 
		return this;
	}
	
	public Block setFooter(String footerName){
		this.footer=new Cell(footerName);
		return this;
	}
	
	public Block setFooter(Cell footer){
		this.footer=footer;
		return this;
	}
	
	public Cell getFooter(){
		return this.footer;
	}
	
	public boolean isShowRowNo() {
		return showRowNo;
	}
	public Block setShowRowNo(boolean showRowNo) {
		this.showRowNo = showRowNo;
		return this;
	}

	public Style getTitleStyle() {
		return titleStyle;
	}

	public Block setTitleStyle(Style titleStyle) {
		if(this.titleStyle==null){
			this.titleStyle=new Style();
		}
		setStyle(this.titleStyle,titleStyle);
		return this;
	}

	public Style getHeaderStyle() {
		return headerStyle;
	}

	public Block setHeaderStyle(Style headerStyle) {
		if(this.headerStyle==null){
			this.headerStyle=new Style();
		}
		setStyle(this.headerStyle,headerStyle);
		return this;
	}

	public Style getRowsStyle() {
		return rowsStyle;
	}

	public Block setRowsStyle(Style rowStyle) {
		if(this.rowsStyle==null){
			this.rowsStyle=new Style();
		}
		setStyle(this.rowsStyle,rowStyle);
		return this;
	}

	public Style getFooterStyle() {
		return footerStyle;
	}

	public Block setFooterStyle(Style footerStyle) {
		if(this.footerStyle==null){
			this.footerStyle=new Style();
		}
		setStyle(this.footerStyle,footerStyle);
		return this;
	}
	
	public int getColCount(){
		int rowCount=0;
		if(this.getHeader()!=null){
			rowCount=this.getHeader().size();
			if(this.isShowRowNo()){
				rowCount++;
			}
		}else{
			throw new RuntimeException("表头不能为空");
		}
		return rowCount;
	}

	public String getRowNoName() {
		return rowNoName;
	}

	public Block setRowNoName(String rowNoName) {
		this.rowNoName = rowNoName;
		return this;
	}
	
	private Block setStyle(Style toStyle, Style fromStyle) {
		CopyUtil.copyIgnoreNull(fromStyle,toStyle);
//		if(fromStyle.getBackColor() != null){
//			toStyle.setBackColor(fromStyle.getBackColor());
//		}
//		if(fromStyle.getFontColor()!=null){
//			toStyle.setFontColor(fromStyle.getFontColor());
//		}
//		if(fromStyle.getSize()!=null){
//			toStyle.setSize(fromStyle.getSize());
//		}
//		if(fromStyle.getBold() != null){
//			toStyle.setBold(fromStyle.getBold());
//		}
//		if(fromStyle.getAlign() != null){
//			toStyle.setAlign(fromStyle.getAlign());
//		}
//		if(fromStyle.getFontName() != null){
//			toStyle.setFontName(fromStyle.getFontName());
//		}
		return this;
	}

//	public List<Row> getSpecialHeader() {
//		return specialHeader;
//	}
//
//	public void setSpecialHeader(List<Row> specialHeader) {
//		this.specialHeader = specialHeader;
//	}
//
//	public Style getSpecialHeaderStyle() {
//		return specialHeaderStyle;
//	}
//
//	public void setSpecialHeaderStyle(Style specialHeaderStyle) {
//		this.specialHeaderStyle = specialHeaderStyle;
//	}

}
