package com.huar.ai;

import com.huar.ai.utils.POIUtils;
import org.apache.http.entity.ContentType;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * @author zhang
 * @description
 * @date 2021/12/17 09:54
 **/
public class DataValidity {

  public static void main(String[] args) throws IOException {

	  String rootPath = System.getProperty("user.dir");
	  String filePath = rootPath + "/src/main/resources/workbook.xlsx";
	  File file = new File(filePath);
	  FileInputStream fileInputStream = new FileInputStream(file);
	  MultipartFile multipartFile = new MockMultipartFile(file.getName(), file.getName(),
			  ContentType.APPLICATION_OCTET_STREAM.toString(), fileInputStream);

	  List<List<Map<String, Object>>> mapList = POIUtils.getExcelDataValidations(multipartFile);
	  mapList.forEach(list -> {
		  list.forEach(map -> {
			  map.forEach((k, v) -> {
				  System.out.println(k + ":" + v);
			  });
			  System.out.println("===============================================");
		  });
	  });


  }
}

