package org.simplepoi.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.net.URISyntaxException;

public abstract class ExelCommonUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExelCommonUtil.class);
    /**
     * @param photoByte
     * @return
     */
    public static String getFileExtendName(byte[] photoByte) {
        String strFileExtendName = "JPG";
        if ((photoByte[0] == 71) && (photoByte[1] == 73) && (photoByte[2] == 70) && (photoByte[3] == 56) && ((photoByte[4] == 55) || (photoByte[4] == 57)) && (photoByte[5] == 97)) {
            strFileExtendName = "GIF";
        } else if ((photoByte[6] == 74) && (photoByte[7] == 70) && (photoByte[8] == 73) && (photoByte[9] == 70)) {
            strFileExtendName = "JPG";
        } else if ((photoByte[0] == 66) && (photoByte[1] == 77)) {
            strFileExtendName = "BMP";
        } else if ((photoByte[1] == 80) && (photoByte[2] == 78) && (photoByte[3] == 71)) {
            strFileExtendName = "PNG";
        }
        return strFileExtendName;
    }


    public static String getWebRootPath(String filePath) {
        try {
            String path = null;
            try {
                path = ReflectionUtil.class.getClassLoader().getResource("").toURI().getPath();
            } catch (URISyntaxException e) {
                //e.printStackTrace();
                //update-begin-author:taoyan date:20211116 for: JAR包分离 发布出空指针 https://gitee.com/jeecg/jeecg-boot/issues/I4CMHK
            } catch (NullPointerException e) {
                path = ReflectionUtil.class.getProtectionDomain().getCodeSource().getLocation().getPath();
            }
            //update-end-author:taoyan date:20211116 for: JAR包分离 发布出空指针 https://gitee.com/jeecg/jeecg-boot/issues/I4CMHK
            //update-begin--Author:zhangdaihao  Date:20190424 for：解决springboot 启动模式，上传路径获取为空问题---------------------
            if (path == null || path == "") {
                //解决springboot 启动模式，上传路径获取为空问题
                path = ReflectionUtil.class.getClassLoader().getResource("").getPath();
            }
            //update-end--Author:zhangdaihao  Date:20190424 for：解决springboot 启动模式，上传路径获取为空问题----------------------
            LOGGER.debug("--- getWebRootPath ----filePath--- " + path);
            path = path.replace("WEB-INF/classes/", "");
            path = path.replace("file:/", "");
            LOGGER.debug("--- path---  " + path);
            LOGGER.debug("--- filePath---  " + filePath);
            return path + filePath;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static String trimAllWhitespace(String str) {
        int len = str.length();
        int st = 0;
        while ((st < len) && (Character.isWhitespace(str.charAt(st)) || str.charAt(st) == '\u00A0')) {
            st++;
        }
        while ((st < len) && (Character.isWhitespace(str.charAt(st)) || str.charAt(st) == '\u00A0')) {
            len--;
        }
        return ((st > 0) || (len < str.length())) ? str.substring(st, len) : str;
    }
}
