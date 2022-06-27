package com.heifan.common.utils;

import cn.hutool.poi.excel.ExcelFileUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import lombok.extern.slf4j.Slf4j;

import java.io.*;
import java.util.List;
import java.util.Map;

/**
 * 　Excel工具类
 *
 * @author HiF
 * @date 2022/5/9 9:24
 */
@Slf4j
public class HeiFanExcelUtils {

    public static void main(String[] args) throws FileNotFoundException {
        parseExcel("C:\\Users\\HeiFan\\Desktop\\78.xlsx", 0);
    }

    static void parseExcel(String filePath, Integer sheetIndex) throws FileNotFoundException {
        File file = new File(filePath);
        FileInputStream fileInputStream = new FileInputStream(file);
        ExcelReader reader = ExcelUtil.getReader(fileInputStream, sheetIndex);
        List<Map<String, Object>> maps = reader.readAll();
        for (Map<String, Object> data : maps) {
            Object name = data.get("姓名");
            Object age = data.get("年龄");
            log.info("name:{}", name);
            log.info("age:{}", age);
        }
    }
}
