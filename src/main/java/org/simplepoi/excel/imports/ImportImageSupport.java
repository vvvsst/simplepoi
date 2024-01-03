package org.simplepoi.excel.imports;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.simplepoi.excel.ImportParams;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.simplepoi.excel.ExelCommonUtil.getFileExtendName;
import static org.simplepoi.excel.ExelCommonUtil.getWebRootPath;

public abstract class ImportImageSupport {
    /**
     * @param object
     * @param picId
     * @param excelParams
     * @param titleString
     * @param pictures
     * @param params
     * @throws Exception
     */
    public static void saveImage(Object object, String picId, Map<String, ExcelImportServer.ExcelImportEntity> excelParams, String titleString,
                                 Map<String, PictureData> pictures, ImportParams params) throws Exception {
        if (pictures == null || pictures.get(picId) == null) {
            return;
        }
        PictureData image = pictures.get(picId);
        byte[] data = image.getData();
        String fileName = "pic" + Math.round(Math.random() * 100000000000L);
        fileName += "." + getFileExtendName(data);
        //update-beign-author:taoyan date:20200302 for:【多任务】online 专项集中问题 LOWCOD-159
        int saveType = excelParams.get(titleString).getSaveType();
        if (saveType == 1) {
            String path = getWebRootPath(getSaveUrl(excelParams.get(titleString), object));
            File savefile = new File(path);
            if (!savefile.exists()) {
                savefile.mkdirs();
            }
            savefile = new File(path + "/" + fileName);
            FileOutputStream fos = new FileOutputStream(savefile);
            fos.write(data);
            fos.close();
            excelParams.get(titleString).getMethod().invoke(object, getSaveUrl(excelParams.get(titleString), object) + "/" + fileName);
        } else if (saveType == 2) {
            excelParams.get(titleString).getMethod().invoke(object, data);
        }
        //update-end-author:taoyan date:20200302 for:【多任务】online 专项集中问题 LOWCOD-159
    }



    /**
     * 获取保存的真实路径
     *
     * @param excelImportEntity
     * @param object
     * @return
     * @throws Exception
     */
    private static String getSaveUrl(ExcelImportServer.ExcelImportEntity excelImportEntity, Object object) throws Exception {
        String url = "";
        if (excelImportEntity.getSaveUrl().equals("upload")) {
            url = object.getClass().getName().split("\\.")[object.getClass().getName().split("\\.").length - 1];
            return excelImportEntity.getSaveUrl() + "/" + url.substring(0, url.lastIndexOf("Entity"));
        }
        return excelImportEntity.getSaveUrl();
    }


    /**
     * 获取Excel2007图片
     *
     * @param sheet
     *            当前sheet对象
     * @param workbook
     *            工作簿对象
     * @return Map key:图片单元格索引（1_1）String，value:图片流PictureData
     */
    public static Map<String, PictureData> getSheetPictrues07(XSSFSheet sheet, XSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        for (POIXMLDocumentPart dr : sheet.getRelations()) {
            if (dr instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) dr;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    String picIndex = ctMarker.getRow() + "_" + ctMarker.getCol();
                    sheetIndexPicMap.put(picIndex, pic.getPictureData());
                }
            }
        }
        return sheetIndexPicMap;
    }

    /**
     * 获取Excel2003图片
     *
     * @param sheet
     *            当前sheet对象
     * @param workbook
     *            工作簿对象
     * @return Map key:图片单元格索引（1_1）String，value:图片流PictureData
     */
    public static Map<String, PictureData> getSheetPictrues03(HSSFSheet sheet, HSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (!pictures.isEmpty()) {
            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    int pictureIndex = pic.getPictureIndex() - 1;
                    HSSFPictureData picData = pictures.get(pictureIndex);
                    String picIndex = String.valueOf(anchor.getRow1()) + "_" + String.valueOf(anchor.getCol1());
                    sheetIndexPicMap.put(picIndex, picData);
                }
            }
            return sheetIndexPicMap;
        } else {
            return null;
        }
    }

}
