package com.fantom.excel;

import com.fantom.excel.core.ExcelReader;
import com.fantom.excel.core.ExcelWriter;

public class ExcelIO {
    public static ExcelReader getReader() {
        return new ExcelReader();
    }

    public static ExcelWriter getWriter() {
        return new ExcelWriter();
    }
}
