package org.simplepoi.test;

import org.junit.Test;

public class ExcelTest {

    @Test
    public void test1(){
        String userHome = System.getProperty("user.home");
        String desktopPath = userHome + "\\Desktop";
        System.out.println(desktopPath);
        System.out.println("ok");
    }

    @Test
    public void whiteCodePointTest(){
//        String.valueOf(result).replaceAll("\\u00A0", " ").trim()
        String s = trimAllWhitespace("aaaa\u00A0dffg");
        String s2 = trimAllWhitespace("aaaa dffg");
        String s3 = trimAllWhitespace("aaaa\rdffg");
        System.out.println(s);
        System.out.println(s2);
        System.out.println(s3);
        String trim = s3.replaceAll("\\u00A0", " ").trim();
        System.out.println(trim);
        System.out.println("ok");

    }

    public String trimAllWhitespace(String str) {
        if (str == null) return null;
        if (str.length() == 0) return str;
        int len = str.length();
        StringBuilder sb = new StringBuilder(str.length());
        for (int i = 0; i < len; i++) {
            char c = str.charAt(i);
            if (!Character.isWhitespace(c) && c != '\u00A0') { // " " \r is included
                sb.append(c);
            }
        }
        return sb.toString();
    }

    @Test
    public void whiteCodePointTest2(){
//        String.valueOf(result).replaceAll("\\u00A0", " ").trim()
        String s = trimAllWhitespace2("\u00A0aaaa\u00A0dffg");
        String s2 = trimAllWhitespace2(" aaaa dffg");
        String s3 = trimAllWhitespace2("\naaaa\rdffg");
        System.out.println("\u00A0aaaa\u00A0dffg");
        System.out.println(s);
        System.out.println(s2);
        System.out.println(s3);
        String trim = s3.replaceAll("\\u00A0", " ").trim();
        System.out.println(trim);
        System.out.println("ok");

    }

    public static String  trimAllWhitespace2(String str) {
        int len = str.length();
        int st = 0;
        while ((st < len) && (Character.isWhitespace(str.charAt(st)) || str.charAt(st) == '\u00A0' )) {
            st++;
        }
        while ((st < len) && (Character.isWhitespace(str.charAt(st)) || str.charAt(st) == '\u00A0' )) {
            len--;
        }
        return ((st > 0) || (len < str.length())) ? str.substring(st, len) : str;
    }
}
