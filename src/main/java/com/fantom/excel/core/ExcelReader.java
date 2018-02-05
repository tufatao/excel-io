package com.fantom.excel.core;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelReader {
    private Workbook workbook;

    public ExcelReader read(File file) throws IOException, InvalidFormatException {
        return read(file, null);
    }

    public ExcelReader read(File file, String password) throws IOException, InvalidFormatException {
        this.workbook = WorkbookFactory.create(file, password);
        return this;
    }

    public ExcelReader read(InputStream inputStream) throws IOException, InvalidFormatException {
        return read(inputStream, null);
    }

    public ExcelReader read(InputStream inputStream, String password) throws IOException, InvalidFormatException {
        this.workbook = WorkbookFactory.create(inputStream, password);
        return this;
    }

    public List<Map<String, Object>> toMaplist(String sheetName, int titleRowIndex) {
        return toMaplist(sheetName, 0, titleRowIndex);
    }

    public List<Map<String, Object>> toMaplist(int sheetIndex, int titleRowIndex) {
        return toMaplist(null, sheetIndex, titleRowIndex);
    }

    public List<Map<String, Object>> toMaplist() {
        return toMaplist(null, 0, 0);
    }

    /**
     * Workbook 中 指定Sheet 转换为 List<Map<String, Object>>
     *
     * @param sheetName     sheet名字
     * @param sheetIndex    sheet位置
     * @param titleRowIndex titleRow位置
     * @return
     * @throws IOException
     */
    public List<Map<String, Object>> toMaplist(String sheetName, int sheetIndex, int titleRowIndex) {
        List<Map<String, Object>> result = new ArrayList<>();
        Sheet sheet;
        if (sheetName != null && sheetName.length() > 0) {
            sheet = this.workbook.getSheet(sheetName);
        } else {
            sheet = this.workbook.getSheetAt(sheetIndex);
        }
        int lastRowNum = sheet.getLastRowNum();
        Row titleRow = null;
        int lastCellNum = 0;
        String[] titles = new String[0];
        for (int j = titleRowIndex; j < lastRowNum; j++) {
            Map<String, Object> map = new HashMap<>();
            if (j == titleRowIndex) {
                titleRow = sheet.getRow(j);
                lastCellNum = titleRow.getLastCellNum();
                titles = new String[lastCellNum];
            }
            Row dataRow = sheet.getRow(j + 1);
            for (int k = 0; k < lastCellNum; k++) {
                if (j == titleRowIndex) {
                    titles[k] = titleRow.getCell(k).getStringCellValue();
                }
                map.put(titles[k], dataRow.getCell(k).toString());
            }
            result.add(map);
        }
        return result;
    }

    public Workbook getWorkbook() {
        return workbook;
    }
}
