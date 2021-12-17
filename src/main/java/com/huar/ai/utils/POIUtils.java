package com.huar.ai.utils;


import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description
 * @Author zhang
 * @Date 2021/10/23 11:09
 **/
public class POIUtils {
	private final static String xls = "xls";
	private final static String xlsx = "xlsx";
	private final static String DATE_FORMAT = "yyyy/MM/dd";



	public static List<List<Map<String, Object>>> getExcelDataValidations(MultipartFile file) throws IOException {
		//检查文件
		checkFile(file);
		//获得Workbook工作薄对象
		Workbook workbook = getWorkBook(file);
		//创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回
		List<List<Map<String, Object>>> list = new ArrayList<>();
		if (workbook != null) {
			for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
				List<Map<String, Object>> sheetList = new ArrayList<>();
				//获得当前sheet工作表
				Sheet sheet = workbook.getSheetAt(sheetNum);
				if (sheet == null) {
					continue;
				}
				//获得当前sheet的开始行
				int firstRowNum = sheet.getFirstRowNum();
				String sheetName = sheet.getSheetName();
				//获得当前sheet的结束行
				int lastRowNum = sheet.getLastRowNum();
				List<? extends DataValidation> dataValidations = sheet.getDataValidations();
				for (DataValidation validation : dataValidations) {
					Map<String, Object> map = new LinkedHashMap<>();
					map.put("sheetName", sheetName);
					CellRangeAddressList addressList = validation.getRegions();
					//空值判断
					if (null == addressList || addressList.getSize() == 0) {
						continue;
					}
					//获取单元格行位置
					int rowNum = addressList.getCellRangeAddress(0).getFirstRow() + 1;
					//获取单元格列位置
					int columnNum = addressList.getCellRangeAddress(0).getFirstColumn() + 1;
					String columnString = CodeConvert(columnNum);
					String cellPosition = columnString + rowNum;
					map.put("cellPosition", cellPosition);
					//根据位置信息判断是不是自己想要获取的单元格位置，比如我的单元格是A1，则对应的坐标为1，1
					DataValidationConstraint constraint = validation.getValidationConstraint();
					// String valueRange = constraint.getFormula1();
					// map.put("valueRange", valueRange);
					//获取单元格数组
					String[] strs = constraint.getExplicitListValues();
					map.put("valueRange", StringUtils.join(strs,","));
					Row row = sheet.getRow(rowNum - 1);
					Cell cell = row.getCell(columnNum - 1);
					String cellValue = getCellValue(cell);
					map.put("cellValue", cellValue);
					sheetList.add(map);
				}
				list.add(sheetList);
			}
			workbook.close();
		}
		return list;
	}

	public static String CodeConvert(int n) {
		StringBuilder s = new StringBuilder();
		while (n > 0) {
			int m = n % 26;
			if (m == 0) m = 26;
			s.insert(0, (char) (m + 64));
			n = (n - m) / 26;
		}
		return s.toString();
	}

	//校验文件是否合法
	public static void checkFile(MultipartFile file) throws IOException {
		//判断文件是否存在
		if (null == file) {
			throw new FileNotFoundException("文件不存在！");
		}
		//获得文件名
		String fileName = file.getOriginalFilename();
		//判断文件是否是excel文件
		assert fileName != null;
		if (!fileName.endsWith(xls) && !fileName.endsWith(xlsx)) {
			throw new IOException(fileName + "不是excel文件");
		}
	}

	public static Workbook getWorkBook(MultipartFile file) {
		//获得文件名
		String fileName = file.getOriginalFilename();
		//创建Workbook工作薄对象，表示整个excel
		Workbook workbook = null;
		try {
			//获取excel文件的io流
			InputStream is = file.getInputStream();
			//根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
			assert fileName != null;
			if (fileName.endsWith(xls)) {
				//2003
				workbook = new HSSFWorkbook(is);
			} else if (fileName.endsWith(xlsx)) {
				//2007
				workbook = new XSSFWorkbook(is);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		return workbook;
	}

	public static String getCellValue(Cell cell) {
		String cellValue = "";
		if (cell == null) {
			return cellValue;
		}
		//如果当前单元格内容为日期类型，需要特殊处理
		String dataFormatString = cell.getCellStyle().getDataFormatString();
		if (dataFormatString.equals("m/d/yy")) {
			cellValue = new SimpleDateFormat(DATE_FORMAT).format(cell.getDateCellValue());
			return cellValue;
		}
		//把数字当成String来读，避免出现1读成1.0的情况
		if (cell.getCellType() == CellType.NUMERIC.getCode()) {
			cell.setCellType(CellType.STRING);
		}
		//判断数据的类型
		switch (cell.getCellTypeEnum()) {
			case NUMERIC: //数字
				cellValue = String.valueOf(cell.getNumericCellValue());
				break;
			case STRING: //字符串
				cellValue = String.valueOf(cell.getStringCellValue());
				break;
			case BOOLEAN: //Boolean
				cellValue = String.valueOf(cell.getBooleanCellValue());
				break;
			case FORMULA: //公式
				cellValue = String.valueOf(cell.getCellFormula());
				break;
			case BLANK: //空值
				cellValue = "";
				break;
			case ERROR: //故障
				cellValue = "非法字符";
				break;
			default:
				cellValue = "未知类型";
				break;
		}
		return cellValue;
	}
}

