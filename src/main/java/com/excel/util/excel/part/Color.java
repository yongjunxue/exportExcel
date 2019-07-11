package com.excel.util.excel.part;

import org.apache.poi.hssf.util.HSSFColor;

@SuppressWarnings("deprecation")
public interface Color {
	public static final short BLACK=new HSSFColor.BLACK().getIndex();
	public static final short WHITE=new HSSFColor.WHITE().getIndex();
	public static final short RED=new HSSFColor.RED().getIndex();
	public static final short GREEN=new HSSFColor.GREEN().getIndex();
	public static final short BLUE=new HSSFColor.BLUE().getIndex();
	
	public static final short LIGHT_TURQUOISE=new HSSFColor.LIGHT_TURQUOISE().getIndex();
	public static final short LIGHT_ORANGE=new HSSFColor.LIGHT_ORANGE().getIndex();
	public static final short LIGHT_GREEN=new HSSFColor.LIGHT_GREEN().getIndex();
	public static final short LIGHT_CORNFLOWER_BLUE=new HSSFColor.LIGHT_CORNFLOWER_BLUE().getIndex();
}
