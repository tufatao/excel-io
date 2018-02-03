package com.fantom.excel.core;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public class ExcelWriter {
    private Workbook workbook;

    public ExcelWriter generate(String maplistJson, String sheetName) {
        System.out.println("Generate maplist to Workbook");
        this.workbook = null;
        return this;
    }

    public ExcelWriter generate(List<Map<String, Object>> maplist, String sheetName) {
        System.out.println("Generate maplist to Workbook");
        this.workbook = null;
        return this;
    }

    public void write(File file) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            this.workbook.write(fos);
            System.out.println("Write excel to file");
        }
    }

    public void write(OutputStream outputStream) throws IOException {
        try (OutputStream os = outputStream) {
            this.workbook.write(os);
            System.out.println("Write excel to outputStream");
        }
    }
}
