
package org.simplepoi.excel;


import java.util.List;

/**
 * 导入参数设置
 * 
 * @author JEECG
 * @date 2013-9-24
 * @version 1.0
 */
public class ImportParams  {
	/**
	 * 表格标题行数,默认0
	 */
	private int titleRows = 0;
	/**
	 * 表头行数,默认1
	 */
	private int headRows = 1;
	private int headColumns = 1;

 //update-begin-author:liusq date:20220605 for:https://gitee.com/jeecg/jeecg-boot/issues/I57UPC 导入 ImportParams 中没有startSheetIndex参数
	/**
	 * 开始读取的sheet位置,默认为0
	 */
	private int                 startSheetIndex  = 0;
	//update-end-author:liusq date:20220605 for:https://gitee.com/jeecg/jeecg-boot/issues/I57UPC 导入 ImportParams 中没有startSheetIndex参数

	//update-begin-author:taoyan date:20211210 for:https://gitee.com/jeecg/jeecg-boot/issues/I45C32 导入空白sheet报错
	/**
	 * 上传表格需要读取的sheet 数量,默认为0
	 */
	private int sheetNum = 0;
	//update-end-author:taoyan date:20211210 for:https://gitee.com/jeecg/jeecg-boot/issues/I45C32 导入空白sheet报错

	public ImportParams( ) {
	}
	public ImportParams(int headRows,int headColumns) {
		this.headRows = headRows;
		this.headColumns = headColumns;
	}

	/**
	 * 最后的无效行数
	 */
	private int lastOfInvalidRow = 0;

	public int getHeadRows() {
		return headRows;

	}
	public int getHeadColumns() {
		return headColumns;
	}

	public int getSheetNum() {
		return sheetNum;
	}


	public int getTitleRows() {
		return titleRows;
	}

	public void setHeadRows(int headRows) {
		this.headRows = headRows;
	}

	public void setSheetNum(int sheetNum) {
		this.sheetNum = sheetNum;
	}

	public void setTitleRows(int titleRows) {
		this.titleRows = titleRows;
	}


	public int getLastOfInvalidRow() {
		return lastOfInvalidRow;
	}

	public void setLastOfInvalidRow(int lastOfInvalidRow) {
		this.lastOfInvalidRow = lastOfInvalidRow;
	}


	public int getStartSheetIndex() {
		return startSheetIndex;
	}

	public void setStartSheetIndex(int startSheetIndex) {
		this.startSheetIndex = startSheetIndex;
	}

}
