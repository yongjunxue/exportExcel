package com.excel.util.excel.part;

import org.apache.poi.ss.usermodel.HorizontalAlignment;


public class Style {
	
	private Short backColor;//背景色
	
	private Short fontColor;//字体颜色
	
	private Short size;//字体大小
	
	private Boolean bold;//加粗
	
	private String fontName="微软雅黑";//字体名字
	
	private HorizontalAlignment align=HorizontalAlignment.CENTER;//对齐方式
	
	public Short getBackColor() {
		return backColor;
	}

	public Style setBackColor(Short backColor) {
		this.backColor = backColor;
		return this;
	}

	public Short getFontColor() {
		return fontColor;
	}

	public Style setFontColor(Short fontColor) {
		this.fontColor = fontColor;
		return this;
	}

	public Short getSize() {
		return size;
	}

	public Style setSize(Short size) {
		this.size = size;
		return this;
	}
	
	public Boolean getBold() {
		return bold;
	}

	public Style setBold(Boolean bold) {
		this.bold = bold;
		return this;
	}

	public HorizontalAlignment getAlign() {
		return align;
	}

	public Style setAlign(HorizontalAlignment align) {
		this.align = align;
		return this;
	}

	public String getFontName() {
		return fontName;
	}

	public Style setFontName(String fontName) {
		this.fontName = fontName;
		return this;
	}
}
