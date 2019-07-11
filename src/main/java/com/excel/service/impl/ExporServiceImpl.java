package com.excel.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.boot.configurationprocessor.json.JSONObject;
import org.springframework.stereotype.Service;

import com.excel.service.ExportService;
import com.excel.util.excel.part.Block;
import com.excel.util.excel.part.Cell;
import com.excel.util.excel.part.Color;
import com.excel.util.excel.part.ExcelModel;
import com.excel.util.excel.part.Row;
import com.excel.util.excel.part.Sheet;
import com.excel.util.excel.part.SpecialSheet;
import com.excel.util.excel.part.Style;

@Service
public class ExporServiceImpl implements ExportService{

	@Override
	public ExcelModel getExcelModel() {
		ExcelModel em=new ExcelModel("导出测试");
		
		//-------导出1，数据区统一使用一种样式----------
		List<Map<String,Object>> list = getTestData();
        Sheet s= em.createSheet();
        s.setBlockSpace(3);
        Block b=s.createBlock();
        b.setTitle("aaa").setTitleStyle((new Style().setFontColor(Color.GREEN)));
        b.setHeader("名字","手机号","身份证");
        b.setHeaderKeys("name","phonenumber","idcard");
        b.setData(list).setRowsStyle((new Style().setFontColor(Color.LIGHT_ORANGE)));//****************************
        b.setFooter("我是表尾");
        b.setShowRowNo(true);
        
        //-------导出2，样式可以具体到单元格-----------
        JSONObject arguments=new JSONObject();
        Block b2=s.createBlock();
        b2.setTitle("bbb").setTitleStyle((new Style().setFontColor(Color.GREEN)));
        b2.setHeader("名字","手机号","身份证","邮箱").setHeaderStyle((new Style().setFontColor(Color.RED)).setBold(true));
        b2.setHeaderKeys("name","phonenumber","idcard","email");
        b2.setRowsStyle((new Style().setBackColor(Color.LIGHT_ORANGE)));
        setRows(b2,list);//可以设置某一行或者某一个单元格的样式
        b2.setShowRowNo(true);//显示序号
        
      //-------导出3.无规则数据的导出-----------
        SpecialSheet ss=em.createSpecialSheet("合并单元格测试");
        Map<String,Object> m=new HashMap<>();
        m.put("name", "任务名称：我的任务");
        m.put("date", "执行时间：2018-09-26 16:49:33----2018-09-27 16:49:35");
        m.put("userName", "创建人：王五");
        m.put("timeLimit", "时间限制：22");
        m.put("totalScore", "得分：60");
        m.put("count", "子任务数：3");
        setSpecialSheet(ss,m);
		
		return em;
	}

	private void setRows(Block b, List<Map<String,Object>> mList) {
		List<Row> rs=new ArrayList<>();
		for(Map rowMap : mList){
			Row r=new Row();
			for(String key : b.getHeaderKeys()){
				if(rowMap.get(key) != null){
					Cell c=r.createCell(rowMap.get(key).toString());
					if(key.equals("name")){
						if(rowMap.get(key)!=null && rowMap.get(key).toString().contains("2")){
							r.setStyle((new Style().setBackColor(Color.BLUE)));//如果名字中包含“2”,则将整行设置为这种样式
						}
						if(rowMap.get(key)!=null && rowMap.get(key).toString().contains("8")){
							c.setStyle(new Style().setBackColor(Color.LIGHT_CORNFLOWER_BLUE).setSize((short) 15));//如果名字中包含“8”,则将这个单元格设为这种样式
						}
					}else if(key.equals("code")){
						if(rowMap.get(key)!=null && rowMap.get(key).toString().contains("34")){
							c.setStyle(new Style().setBackColor(Color.LIGHT_CORNFLOWER_BLUE).setBold(true));
						}
					}
				}
				
			}
			rs.add(r);
		}
		b.setRows(rs);
	}
	
	private void setSpecialSheet(SpecialSheet ss, Map m) {
		List<Row> rs=new ArrayList<>();
		Row r=new Row();
		r.createCell("标题").setMergeRange(0, 3).setWidth(50000).setStyle(new Style().setSize((short) 20).setBold(true).setAlign(HorizontalAlignment.CENTER));
		rs.add(r);
		r=new Row();
		short size=12;
		
		Style style=new Style().setSize(size ).setAlign(HorizontalAlignment.LEFT);
		r.createCell(m.get("name").toString()).setMergeRange(0, 1).setWidth(15000).setStyle(style);
		r.createCell(m.get("date").toString()).setMergeRange(2, 3).setWidth(15000).setStyle(style);
		rs.add(r);
		r=new Row();
		r.createCell(m.get("userName").toString()).setWidth(10000).setStyle(style );
		r.createCell(m.get("timeLimit").toString()).setWidth(10000).setStyle(style);
		r.createCell(m.get("totalScore").toString()).setWidth(10000).setStyle(style );
		r.createCell(m.get("count").toString()).setWidth(10000).setStyle(style);
		rs.add(r);
		ss.setRows(rs);
	}

	private List<Map<String,Object>> getTestData() {
		List<Map<String,Object>> list =new ArrayList<>();
		Map<String,Object> item=new HashMap<>();
		item.put("name", "张三");//"name","phonenumber","idcard","email"
		item.put("phonenumber", "11111111111");
		item.put("idcard", "1");
		item.put("email", "123@qq.com");
		list.add(item);
		
		item=new HashMap<>();
		item.put("name", "李四");//"name","phonenumber","idcard","email"
		item.put("phonenumber", "22222222222");
		item.put("idcard", "2");
		item.put("email", "22@qq.com");
		list.add(item);
		
		item=new HashMap<>();
		item.put("name", "李8");//"name","phonenumber","idcard","email"
		item.put("phonenumber", "444444");
		item.put("idcard", "34");
		item.put("email", "22er@qq.com");
		list.add(item);
		
		return list;
	}

	@Override
	public List<Map<String, Object>> getUserList(Map<String, Object> params) {
		return getTestData();
	}

	@Override
	public List<Map<String, Object>> getProvinceList(Map<String, Object> params) {
		List<Map<String,Object>> list =new ArrayList<>();
		Map<String,Object> item=new HashMap<>();
		item.put("code", "001");
		item.put("name", "北京");
		list.add(item);
		
		item=new HashMap<>();
		item.put("code", "002");
		item.put("name", "天津");
		list.add(item);
		
		item=new HashMap<>();
		item.put("code", "003");
		item.put("name", "上海");
		list.add(item);
		
		return list;
	}

}
