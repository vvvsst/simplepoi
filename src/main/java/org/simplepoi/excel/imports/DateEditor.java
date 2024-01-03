package org.simplepoi.excel.imports;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.ResolverStyle;
import java.util.Date;
import java.util.Locale;

import static org.simplepoi.excel.ExelCommonUtil.trimAllWhitespace;

public class DateEditor implements ExcelPropertyEditor {
    private String[] dateFormats = new String[0];

    public DateEditor(String[] dateFormats) {
        if (dateFormats != null) this.dateFormats = dateFormats;
    }

    public DateEditor(String dateFormat) {
        if (dateFormat != null) this.dateFormats = new String[]{dateFormat};
    }

    @Override
    public boolean supports(Class<?> sourceType, Class<?> targetType) {
        if (sourceType == Date.class || sourceType == String.class)
            return targetType == Date.class || targetType == LocalDate.class || targetType == LocalDateTime.class;
        return false;
    }

    @Override
    public Object convert(Object obj, Class<?> targetType) {
        if (obj.getClass() == Date.class) return convert((Date) obj, targetType);
        if (obj.getClass() == String.class) {
            try {
                return convert((String) obj, targetType);
            } catch (Exception e) {
                e.printStackTrace();
                return null;
            }
        }
        return null;
    }

    private Object convert(Date date, Class<?> targetType) {
        if (targetType == Date.class) return date;
        if (targetType == LocalDate.class) return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        if (targetType == LocalDateTime.class) return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
        return null;
    }

    private Object convert(String dateStr, Class<?> targetType) {
        dateStr = trimAllWhitespace(dateStr);
        DateTimeFormatterBuilder formatterBuilder = new DateTimeFormatterBuilder();
        for (String dateFormat : dateFormats) {
            formatterBuilder.appendOptional(DateTimeFormatter.ofPattern(dateFormat));
        }
        DateTimeFormatter dateTimeFormatter = formatterBuilder.toFormatter(Locale.ENGLISH).withResolverStyle(ResolverStyle.LENIENT);
        if (targetType == LocalDateTime.class)
            return LocalDateTime.parse(dateStr, dateTimeFormatter);
        if (targetType == Date.class)
            return Date.from(LocalDateTime.parse(dateStr, dateTimeFormatter).atZone(ZoneId.systemDefault()).toInstant());
        if (targetType == LocalDate.class)
            return LocalDate.parse(dateStr, dateTimeFormatter);
        return null;
    }
}
