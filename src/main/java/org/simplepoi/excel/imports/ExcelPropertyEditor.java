package org.simplepoi.excel.imports;

public interface ExcelPropertyEditor {

    boolean supports(Class<?> sourceType, Class<?> targetType);

    Object convert(Object obj, Class<?> targetType);

}
