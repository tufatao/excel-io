package com.fantom.excel.core;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

public class ExcelReader {
    private Workbook workbook;

    public ExcelReader read(File file) throws IOException, InvalidFormatException {
        return read(file, null);
    }

    public ExcelReader read(File file, String password) throws IOException, InvalidFormatException {
        this.workbook = WorkbookFactory.create(file, password);
        System.out.println("Read excel as Workbook");
        return this;
    }

    public ExcelReader read(InputStream inputStream) throws IOException, InvalidFormatException {
        return read(inputStream, null);
    }

    public ExcelReader read(InputStream inputStream, String password) throws IOException, InvalidFormatException {
        this.workbook = WorkbookFactory.create(inputStream, password);
        System.out.println("Read excel as Workbook");
        return this;
    }

    public List<Map<String, Object>> toMaplist(String sheetName, int titleRowIndex) {
        System.out.printf("Extract Maplist from Workbook with " +
                "sheetName(%s) titleRowIndex(%d)\n", sheetName, titleRowIndex);
        List<Map<String, Object>> result = null;
        return result;
    }

    public List<Map<String, Object>> toMaplist() {
        return toMaplist(0, 0);
    }

    public List<Map<String, Object>> toMaplist(int sheetIndex, int titleRowIndex) {
        System.out.printf("Extract Maplist from Workbook with " +
                "sheetIndex(%d) titleRowIndex(%d)\n", sheetIndex, titleRowIndex);
        List<Map<String, Object>> result = null;
        return result;
    }
}
