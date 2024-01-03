
package org.simplepoi.excel;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.simplepoi.excel.constant.ExcelType;

/**
 * Excel 导出参数
 * 
 * @author JEECG
 * @version 1.0 2013年8月24日
 */
public class ExportParams   {

	/**
	 * 表格名称
	 */
	private String title;

	/**
	 * 表格名称
	 */
	private short titleHeight = 10;

	/**
	 * 第二行名称
	 */
	private String secondTitle;

	/**
	 * 表格名称
	 */
	private short secondTitleHeight = 8;
	/**
	 * sheetName
	 */
	private String sheetName;


	/**
	 * 是否添加需要需要
	 */
	private String indexName = "序号";
	/**
	 * 冰冻列
	 */
	private int freezeCol;


	/**
	 * Excel 导出版本
	 */
	private ExcelType type = ExcelType.HSSF;

	/**
	 * 是否创建表头
	 */
	private boolean isCreateHeadRows = true;


//update-begin---author:liusq  Date:20220104  for：[LOWCOD-2521]【autopoi】大数据导出方法【全局】----

	/**
	 * 单sheet最大值
	 * 03版本默认6W行,07默认100W
	 */
	private int     maxNum           = 0;
	/**
	 * 导出时在excel中每个列的高度 单位为字符，一个汉字=2个字符
	 * 全局设置,优先使用
	 */
	private short height = 0;


//update-end---author:liusq  Date:20220104  for：[LOWCOD-2521]【autopoi】大数据导出方法【全局】----
	public ExportParams() {

	}

	public ExportParams(String title, String sheetName) {
		this.title = title;
		this.sheetName = sheetName;
	}

	public ExportParams(String title, String sheetName, ExcelType type) {
		this.title = title;
		this.sheetName = sheetName;
		this.type = type;
	}

	public ExportParams(String title, String secondTitle, String sheetName) {
		this.title = title;
		this.secondTitle = secondTitle;
		this.sheetName = sheetName;
	}


	public String getSecondTitle() {
		return secondTitle;
	}

	public short getSecondTitleHeight() {
		return (short) (secondTitleHeight * 50);
	}

	public String getSheetName() {
		return sheetName;
	}

	public String getTitle() {
		return title;
	}

	public short getTitleHeight() {
		return (short) (titleHeight * 50);
	}

	public void setSecondTitle(String secondTitle) {
		this.secondTitle = secondTitle;
	}

	public void setSecondTitleHeight(short secondTitleHeight) {
		this.secondTitleHeight = secondTitleHeight;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public void setTitleHeight(short titleHeight) {
		this.titleHeight = titleHeight;
	}

	public ExcelType getType() {
		return type;
	}

	public void setType(ExcelType type) {
		this.type = type;
	}

	public String getIndexName() {
		return indexName;
	}

	public void setIndexName(String indexName) {
		this.indexName = indexName;
	}

	public int getFreezeCol() {
		return freezeCol;
	}

	public void setFreezeCol(int freezeCol) {
		this.freezeCol = freezeCol;
	}

	public boolean isCreateHeadRows() {
		return isCreateHeadRows;
	}

	public void setCreateHeadRows(boolean isCreateHeadRows) {
		this.isCreateHeadRows = isCreateHeadRows;
	}


	public int getMaxNum() {
		return maxNum;
	}

	public void setMaxNum(int maxNum) {
		this.maxNum = maxNum;
	}

	public short getHeight() {
		return height == -1 ? -1 : (short) (height * 50);
	}

	public void setHeight(short height) {
		this.height = height;
	}
}
