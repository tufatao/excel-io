package com.fantom.excel.core;

import com.fantom.excel.support.JsonUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class ExcelWriter {
    private Workbook workbook;

    public ExcelWriter generate(String mapListJson, String sheetName) {
        System.out.println("Generate mapList to Workbook");
        List<Map<String, Object>> mapList = JsonUtil.json2Obj(mapListJson, List.class);
        return generate(mapList, sheetName);
    }

    /**
     * 生成 <code>Workbook</code>
     * Generate <code>Workbook</code>
     *
     * @param mapList 数据集 List<Map<String, Object>>
     *                Attention: mapList中Object, 数字类型(int, float, double) 已针对处理, 其他类型请先自行处理, 或默认按toString()取值
     * @return <code>ExcelWriter</code>
     * @see Workbook
     * @see HSSFWorkbook
     */
    public ExcelWriter generate(List<Map<String, Object>> mapList, String sheetName) {
        if (mapList == null) {
            throw new NullPointerException("mapList can not be null !");
        }
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);
        if (mapList.size() > 0) {
            Row firstRow = null;
            for (int i = 0; i < mapList.size(); i++) {
                Map<String, Object> map = mapList.get(i);
                Row row = sheet.createRow(i + 1);
                if (i == 0) {
                    firstRow = sheet.createRow(0);
                }
                Iterator<Map.Entry<String, Object>> entryIterator = map.entrySet().iterator();
                int j = 0;
                while (entryIterator.hasNext()) {
                    Map.Entry<String, Object> entry = entryIterator.next();
                    if (i == 0) {
                        String keyObj = entry.getKey();
                        String fieldName = keyObj == null ? "" : keyObj;
                        firstRow.createCell(j).setCellValue(fieldName);
                    }
                    Object valObj = entry.getValue();
                    if (valObj != null) {
                        if (valObj instanceof Integer) {
                            row.createCell(j).setCellValue(((Integer) valObj).intValue());
                        } else if (valObj instanceof Double) {
                            row.createCell(j).setCellValue(((Double) valObj).doubleValue());
                        } else if (valObj instanceof Float) {
                            row.createCell(j).setCellValue(((Float) valObj).floatValue());
                        } else {
                            row.createCell(j).setCellValue(valObj.toString());
                        }
                    } else {
                        row.createCell(j).setCellValue("");
                    }
                    j++;
                }
            }
        }
        this.workbook = workbook;
        return this;
    }

    /**
     * 将 workbook 输出到 文件
     * to file.
     *
     * @param file
     * @throws IOException
     */
    public void write(File file) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            this.workbook.write(fos);
        }
    }

    /**
     * 将 workbook 输出到 outputStream
     * to outputStream.
     *
     * @param outputStream
     * @throws IOException
     */
    public void write(OutputStream outputStream) throws IOException {
        try (OutputStream os = outputStream) {
            this.workbook.write(os);
        }
    }

    public Workbook getWorkbook() {
        return workbook;
    }
}
