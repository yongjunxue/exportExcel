package com.excel.util.excel.part;

import java.util.ArrayList;
import java.util.List;

public class Sheet implements ISheet{
	
	private String sheetName;
	
	private List<Block> blocks=new ArrayList<>();
	
	private int blockSpace=1;//块与块之间的间隔,默认为1个单元格的间隔
	
	private int nextRowNo=0;//下一行
	
	public Sheet(){
		
	}
	
	public Sheet(String sheetName){
		this.setSheetName(sheetName);
	}
	
	public Block createBlock(){
		Block b=new Block();
		blocks.add(b);
		return b;
	}

	public List<Block> getBlocks(){
		return blocks;
	}
	
	public String getSheetName() {
		return sheetName;
	}

	public Sheet setSheetName(String sheetName) {
		this.sheetName = sheetName;
		return this;
	}

	public int getBlockSpace() {
		return blockSpace;
	}

	public Sheet setBlockSpace(int blockSpace) {
		this.blockSpace = blockSpace;
		return this;
	}
	
	public int getMaxColCount(){
		int maxColCount=0;
		for(Block b : blocks){
			if(b.getColCount()>maxColCount){
				maxColCount=b.getColCount();
			}
		}
		return maxColCount;
	}

	public int nextRowNo() {
		return nextRowNo++;
	}

	public int getCurrentRowNo() {
		return nextRowNo-1;
	}

}
