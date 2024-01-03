package org.simplepoi.excel.imports;


import org.apache.commons.lang3.StringUtils;

import java.math.BigDecimal;
import java.math.BigInteger;

public class NumberEditor implements ExcelPropertyEditor {

    @Override
    public boolean supports(Class<?> sourceType, Class<?> targetType) {
        if (sourceType == Double.class || sourceType == String.class)
            return targetType == BigDecimal.class || targetType == Integer.class
                    || targetType == Double.class;
        return false;
    }

    @Override
    public Object convert(Object value, Class<?> targetClass) {
        try {
            if (value.getClass() == Double.class) return convert((Double) value, targetClass);
            if (value.getClass() == String.class) return convert((String) value, targetClass);
        } catch (Exception e){
            e.printStackTrace();
            return null;
        }
        return null;
    }

    private Object convert(Double doubleNum, Class<?> targetClass) {
        if (Short.class == targetClass) {
            return doubleNum.shortValue();
        } else if (Integer.class == targetClass) {
            return doubleNum.intValue();
        } else if (Long.class == targetClass) {
            return doubleNum.longValue();
        } else if (BigInteger.class == targetClass) {
            return new BigDecimal(doubleNum);
        } else if (Float.class == targetClass) {
            return doubleNum.floatValue();
        } else if (Double.class == targetClass) {
            return doubleNum;
        } else if (BigDecimal.class == targetClass || Number.class == targetClass) {
            return new BigDecimal(doubleNum);
        } else {
            throw new IllegalArgumentException(
                    "Cannot convert String [" + doubleNum + "] to target class [" + targetClass.getName() + "]");
        }
    }

    private Object convert(String text, Class<?> targetClass) {
        String trimmed = trimAllWhitespace(text);
        if (StringUtils.isEmpty(trimmed)) return null;
        if (Byte.class == targetClass) {
            return (isHexNumber(trimmed) ? Byte.decode(trimmed) : Byte.valueOf(trimmed));
        } else if (Short.class == targetClass) {
            return (isHexNumber(trimmed) ? Short.decode(trimmed) : Short.valueOf(trimmed));
        } else if (Integer.class == targetClass) {
            return (isHexNumber(trimmed) ? Integer.decode(trimmed) : Integer.valueOf(trimmed));
        } else if (Long.class == targetClass) {
            return (isHexNumber(trimmed) ? Long.decode(trimmed) : Long.valueOf(trimmed));
        } else if (BigInteger.class == targetClass) {
            return (new BigInteger(trimmed));
        } else if (Float.class == targetClass) {
            return Float.valueOf(trimmed);
        } else if (Double.class == targetClass) {
            return Double.valueOf(trimmed);
        } else if (BigDecimal.class == targetClass || Number.class == targetClass) {
            return new BigDecimal(trimmed);
        } else {
            throw new IllegalArgumentException(
                    "Cannot convert String [" + text + "] to target class [" + targetClass.getName() + "]");
        }
    }

    public String trimAllWhitespace(String str) {
        if (!hasLength(str)) {
            return str;
        }

        int len = str.length();
        StringBuilder sb = new StringBuilder(str.length());
        for (int i = 0; i < len; i++) {
            char c = str.charAt(i);
            if (!Character.isWhitespace(c)) {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    public boolean hasLength(String str) {
        return (str != null && !str.isEmpty());
    }

    private boolean isHexNumber(String value) {
        int index = (value.startsWith("-") ? 1 : 0);
        return (value.startsWith("0x", index) || value.startsWith("0X", index) || value.startsWith("#", index));
    }

}
