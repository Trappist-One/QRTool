package com.hkfs.utils;

import com.sun.media.sound.InvalidFormatException;
import jdk.nashorn.internal.runtime.NumberToString;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * @Description:
 * @Author: Bruce Lee
 * @Date: 2021/11/2 11:51
 */
public class ExcelUtil {

    public static List<String> readeExcelHeader(InputStream excelInputSteam,
                                                int sheetNumber,
                                                int headerNumber,
                                                int rowStart) throws IOException, InvalidFormatException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        //要返回的数据
        List<String> headers = new ArrayList<String>();
        //生成工作表
        Workbook workbook = WorkbookFactory.create(excelInputSteam);
        Sheet sheet = workbook.getSheetAt(sheetNumber);
        Row header = sheet.getRow(headerNumber);
        DataFormatter dataFormatter = new DataFormatter();
        for (int i = 0; i < header.getLastCellNum(); i++) {
            //获取单元格
            Cell cell = header.getCell(i);
            headers.add(dataFormatter.formatCellValue(cell));
        }
        return headers;
    }

    public static List<Map<String, Object>> readeExcelData(InputStream excelInputSteam,
                                                           int sheetNumber,
                                                           int headerNumber,
                                                           int rowStart) throws IOException, InvalidFormatException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        //需要的变量以及要返回的数据
        List<Map<String, Object>> result = new ArrayList<Map<String, Object>>();
        List<String> headers = new ArrayList<String>();
        //生成工作表
        Workbook workbook = WorkbookFactory.create(excelInputSteam);
        Sheet sheet = workbook.getSheetAt(sheetNumber);
        Row header = sheet.getRow(headerNumber);
        //最后一行数据
        int rowEnd = sheet.getLastRowNum();
        DataFormatter dataFormatter = new DataFormatter();
        //获取标题信息
        for (int i = 0; i < header.getLastCellNum(); ++i) {
            Cell cell = header.getCell(i);
            headers.add(dataFormatter.formatCellValue(cell));
        }
        //获取内容信息
        for (int i = rowStart; i <= rowEnd; ++i) {
            Row currentRow = sheet.getRow(i);
            if (Objects.isNull(currentRow)) {
                continue;
            }
            Map<String, Object> dataMap = new HashMap<>();
            for (int j = 0; j < currentRow.getLastCellNum(); ++j) {
                //将null转化为Blank
                Cell data = currentRow.getCell(j, Row.CREATE_NULL_AS_BLANK);
                if (Objects.isNull(data)) {     //感觉这个if有点多余
                    dataMap.put(headers.get(j), null);
                } else {
                    switch (data.getCellType()) {   //不同的类型分别进行存储
                        case Cell.CELL_TYPE_STRING:
                            dataMap.put(headers.get(j), data.getRichStringCellValue().getString());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            if (DateUtil.isCellDateFormatted(data)) {
                                dataMap.put(headers.get(j), data.getDateCellValue());
                            } else {
                                dataMap.put(headers.get(j), NumberToString.stringFor(data.getNumericCellValue()));
                            }
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            dataMap.put(headers.get(j), data.getCellFormula());
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            dataMap.put(headers.get(j), data.getBooleanCellValue());
                            break;
                        default:
                            dataMap.put(headers.get(j), null);
                    }
                }
            }
            result.add(dataMap);
        }
        return result;
    }


}
