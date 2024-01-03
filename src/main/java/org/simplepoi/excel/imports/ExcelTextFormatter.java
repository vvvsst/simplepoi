package org.simplepoi.excel.imports;

import org.apache.poi.ss.format.CellFormatPart;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.format.CellTextFormatter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

public class ExcelTextFormatter {
    static final private Logger logger = LoggerFactory.getLogger(ExcelTextFormatter.class);
    static final private Method getCellFormatType;

    static private final ExcelTextFormatter INSTANCE =  new ExcelTextFormatter();

    static {
        try {
            getCellFormatType = CellFormatPart.class.getDeclaredMethod("getCellFormatType");
        } catch (NoSuchMethodException e) {
            throw new RuntimeException(e);
        }
        getCellFormatType.setAccessible(true);
    }

    private ExcelTextFormatter() {
    }

    public static ExcelTextFormatter getInstance() {
        return INSTANCE;
    }

    // value should be String, otherwise call toString
    public String format(String value, String dataFormatString) //  cell.getCellStyle().getDataFormatString()
            throws InvocationTargetException, IllegalAccessException {
        CellFormatPart cellFormatPart = null;
        CellFormatType typeResult = null;
        try {
            if (!dataFormatString.toLowerCase().contains("general"))
                cellFormatPart = new CellFormatPart(dataFormatString);
        } catch (Exception e) {
            logger.warn(e.getMessage());
        }
        if (cellFormatPart != null) {
            typeResult = (CellFormatType) getCellFormatType.invoke(cellFormatPart);
            if (typeResult == CellFormatType.TEXT) return new CellTextFormatter(dataFormatString).format(value);
        }
        return value;
    }


}
