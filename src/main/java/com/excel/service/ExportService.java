package com.excel.service;

import java.util.List;
import java.util.Map;

import com.excel.util.excel.part.ExcelModel;

public interface ExportService {

	ExcelModel getExcelModel();

	List<Map<String, Object>> getUserList(Map<String, Object> params);

	List<Map<String, Object>> getProvinceList(Map<String, Object> params);

}
