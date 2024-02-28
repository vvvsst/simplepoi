package org.simplepoi.excel.export;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.simplepoi.excel.constant.ExcelType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.security.SecureRandom;
import java.security.cert.CertificateException;
import java.security.cert.X509Certificate;

import static org.simplepoi.excel.ExelCommonUtil.getFileExtendName;
import static org.simplepoi.excel.ExelCommonUtil.getWebRootPath;

public abstract class ImageCellCreateSupport {
    private static final Logger LOGGER = LoggerFactory.getLogger(ImageCellCreateSupport.class);
    protected static ExcelType type = ExcelType.HSSF;

    static public void setImageCellType(ExcelType typeIpt) {
        type = typeIpt;
    }

    static public void createImageCell(Drawing patriarch, Row row, int i, String imagePath) throws Exception {
//        row.setHeight((short) (50 * entity.getHeight()));
        row.createCell(i);
        ClientAnchor anchor;
        if (type.equals(ExcelType.HSSF)) {
            anchor = new HSSFClientAnchor(0, 0, 0, 0, (short) i, row.getRowNum(), (short) (i + 1), row.getRowNum() + 1);
        } else {
            anchor = new XSSFClientAnchor(0, 0, 0, 0, (short) i, row.getRowNum(), (short) (i + 1), row.getRowNum() + 1);
        }

        if (StringUtils.isEmpty(imagePath)) {
            return;
        }

        //update-beign-author:taoyan date:20200302 for:【多任务】online 专项集中问题 LOWCOD-159
//        int imageType = entity.getExportImageType();
        int imageType = -1;
        byte[] value = null;
        if (imageType == 2) {
            //原来逻辑 2 // need entity's method to obtain binary byte data
//            value = (byte[]) (entity.getMethods() != null ? getFieldBySomeMethod(entity.getMethods(), obj) : entity.getMethod().invoke(obj, new Object[]{}));
        } else if (imageType == 4 || imagePath.startsWith("http")) {
            //新增逻辑 网络图片4
            try {
                if (imagePath.contains(",")) {
                    if (imagePath.startsWith(",")) {
                        imagePath = imagePath.substring(1);
                    }
                    String[] images = imagePath.split(",");
                    imagePath = images[0];
                }
                if (imagePath.startsWith("https")) {
                    value = getImageDataByHttps(imagePath);
                } else {
                    value = getImageDataByHttp(imagePath);
                }
            } catch (Exception exception) {
                LOGGER.warn(exception.getMessage());
                //exception.printStackTrace();
            }
        } else {
            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            BufferedImage bufferImg;
            String path = null;
            if (imageType == 1) {
                //原来逻辑 1
                path = getWebRootPath(imagePath);
                LOGGER.debug("--- createImageCell getWebRootPath ----filePath--- " + path);
                path = path.replace("WEB-INF/classes/", "");
                path = path.replace("file:/", "");
            }
//            else if (imageType == 3) {
//                //新增逻辑 本地图片3
//                //begin-------author：liusq---data：2021-01-27----for：本地图片ImageBasePath为空报错的问题
//                if (StringUtils.isNotBlank(entity.getImageBasePath())) {
//                    if (!entity.getImageBasePath().endsWith(File.separator) && !imagePath.startsWith(File.separator)) {
//                        path = entity.getImageBasePath() + File.separator + imagePath;
//                    } else {
//                        path = entity.getImageBasePath() + imagePath;
//                    }
//                } else {
//                    path = imagePath;
//                }
//                //end-------author：liusq---data：2021-01-27----for：本地图片ImageBasePath为空报错的问题
//            }
            try {
                bufferImg = ImageIO.read(new File(path));
                //update-begin-author:taoYan date:20211203 for: Excel 导出图片的文件带小数点符号 导出报错 https://gitee.com/jeecg/jeecg-boot/issues/I4JNHR
                ImageIO.write(bufferImg, imagePath.substring(imagePath.lastIndexOf(".") + 1, imagePath.length()), byteArrayOut);
                //update-end-author:taoYan date:20211203 for: Excel 导出图片的文件带小数点符号 导出报错 https://gitee.com/jeecg/jeecg-boot/issues/I4JNHR
                value = byteArrayOut.toByteArray();
            } catch (Exception e) {
                LOGGER.error(e.getMessage());
            }
        }
        if (value != null) { // if the value has been obtained, set the value into the image cell by addPicture()
            patriarch.createPicture(anchor, row.getSheet().getWorkbook().addPicture(value, getImageType(value)));
        }
        //update-end-author:taoyan date:20200302 for:【多任务】online 专项集中问题 LOWCOD-159


    }

    /**
     * 通过https地址获取图片数据
     *
     * @param imagePath
     * @return
     * @throws Exception
     */
    static private byte[] getImageDataByHttps(String imagePath) throws Exception {
        SSLContext sslcontext = SSLContext.getInstance("SSL", "SunJSSE");
        sslcontext.init(null, new TrustManager[]{new MyX509TrustManager()}, new SecureRandom());
        URL url = new URL(imagePath);
        HttpsURLConnection conn = (HttpsURLConnection) url.openConnection();
        conn.setSSLSocketFactory(sslcontext.getSocketFactory());
        conn.setRequestMethod("GET");
        conn.setConnectTimeout(5 * 1000);
        InputStream inStream = conn.getInputStream();
        byte[] value = readInputStream(inStream);
        return value;
    }

    public static class MyX509TrustManager implements X509TrustManager {

        @Override
        public void checkClientTrusted(X509Certificate[] chain, String authType) throws CertificateException {
            // TODO Auto-generated method stub
        }

        @Override
        public void checkServerTrusted(X509Certificate[] chain, String authType) throws CertificateException {
            // TODO Auto-generated method stub

        }

        @Override
        public X509Certificate[] getAcceptedIssuers() {
            // TODO Auto-generated method stub
            return null;
        }

    }

    /**
     * inStream读取到字节数组
     *
     * @param inStream
     * @return
     * @throws Exception
     */
    static private byte[] readInputStream(InputStream inStream) throws Exception {
        if (inStream == null) {
            return null;
        }
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len = 0;
        //每次读取的字符串长度，如果为-1，代表全部读取完毕
        while ((len = inStream.read(buffer)) != -1) {
            outStream.write(buffer, 0, len);
        }
        inStream.close();
        return outStream.toByteArray();
    }

    /**
     * 通过http地址获取图片数据
     *
     * @param imagePath
     * @return
     * @throws Exception
     */
    static private byte[] getImageDataByHttp(String imagePath) throws Exception {
        URL url = new URL(imagePath);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setConnectTimeout(5 * 1000);
        InputStream inStream = conn.getInputStream();
        byte[] value = readInputStream(inStream);
        return value;
    }

    static public int getImageType(byte[] value) {
        String type = getFileExtendName(value);
        if (type.equalsIgnoreCase("JPG")) {
            return Workbook.PICTURE_TYPE_JPEG;
        } else if (type.equalsIgnoreCase("PNG")) {
            return Workbook.PICTURE_TYPE_PNG;
        }
        return Workbook.PICTURE_TYPE_JPEG;
    }

}

