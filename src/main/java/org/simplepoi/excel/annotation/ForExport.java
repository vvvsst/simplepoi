package org.simplepoi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指明这个字段的设置只会用于导出
 * */
@Retention(RetentionPolicy.CLASS) // add some prefix(export/import) to distinguish used in export / import  todo
@Target(ElementType.METHOD)
public @interface ForExport {
}
