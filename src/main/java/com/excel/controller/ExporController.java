package com.excel.controller;

import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import com.excel.service.ExportService;
import com.excel.util.ExcelUtil;
import com.excel.util.excel.part.Block;
import com.excel.util.excel.part.ExcelModel;
import com.excel.util.excel.part.Sheet;

@Controller
public class ExporController {
	
	@Autowired
	ExportService exportService;
	
	@RequestMapping(value = "/export")
    public void export(HttpServletResponse response){
		ExcelModel em=exportService.getExcelModel();
		ExcelUtil.exportXlsx(em, response);
	}
	
	@RequestMapping(value = "/export2")
    public void export2(@RequestParam Map<String, Object> params,HttpServletResponse response){
    	ExcelModel em=new ExcelModel("导出测试");

    	//------工作表1
    	Sheet sheet1 = em.createSheet("用户和省份");
    	sheet1.setBlockSpace(3);//同一个sheet中各个block的间隔为3行
    	
    	List<Map<String,Object>> userList = exportService.getUserList(params);
    	Block block1 = sheet1.createBlock();
    	block1.setHeader("名字","手机号","身份证","邮箱");
    	block1.setHeaderKeys("name","phonenumber","idcard","email");
    	block1.setData(userList);
    	
    	List<Map<String,Object>> provinceList = exportService.getProvinceList(params);
    	Block block2 = sheet1.createBlock();
    	block2.setHeader("省市编号","省份名称");
    	block2.setHeaderKeys("code","name");
    	block2.setData(provinceList);
    	block2.setShowRowNo(true);//控制是否显示序号
    	
    	//------工作表2
    	Sheet sheet2 = em.createSheet("省份");
    	Block block3 = sheet2.createBlock();
    	block3.setTitle("标题");
    	block3.setHeader("省市编号","省份名称");
    	block3.setHeaderKeys("code","name");
    	block3.setData(provinceList);
    	
    	ExcelUtil.exportXlsx(em, response);
	}
}
