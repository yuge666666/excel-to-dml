package com.wangyu.bigdata.application.poi;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ReadExcel {


    private static final Logger logger = LoggerFactory.getLogger(ReadExcel.class);


    //判断指定的单元格是否是合并单元格
    private static boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    // 获取合并单元格的值
    public static CellRegion getMergedRegionValue(Sheet sheet, int row, int column) {
        CellRegion cellregion = new CellRegion();
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {

                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    cellregion.setStartrownum(firstRow);
                    cellregion.setEndrownum(lastRow);
                    cellregion.setValue(getCellValue(fCell));
                    return cellregion;
                }
            }
        }

        return null;
    }


    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {

        if (cell == null) {
            return "";
        }


        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

            return cell.getStringCellValue();

        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

            return String.valueOf(cell.getBooleanCellValue());

        } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

            return String.valueOf(cell.getNumericCellValue());

        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

            return String.valueOf(cell.getNumericCellValue());

        }
        return "";
    }

    //求单元格或者合并单元格的高度
    public static int getheight(Sheet sheet, int rownum, int colnum) {
        if (isMergedRegion(sheet, rownum, colnum)) {
            CellRegion cellresult = getMergedRegionValue(sheet, rownum, colnum);
            if (cellresult != null) {
                return cellresult.getEndrownum() - cellresult.getStartrownum() + 1;
            }
            return 0;
        } else {
            return 1;
        }

    }

    //获取合并或未合并的任意单元格
    public static String getvalue(Sheet sheet, int rownum, int colnum) {
        if (isMergedRegion(sheet, rownum, colnum)) {
            CellRegion cellresults = getMergedRegionValue(sheet, rownum, colnum);
            if (cellresults != null) {
                return cellresults.getValue();
            }
            return "";
        } else {
            Row row = sheet.getRow(rownum);
            Cell cell = row.getCell(colnum);
            return getCellValue(cell);
        }
    }

    public static void main(String[] args) {
        Workbook workbook = null;
        POIFSFileSystem fs = null;
        //设置要读取的文件路径
        String filepath = args[0];
        if (filepath == null) {
            logger.error("输入文件路径为空");
            System.exit(-1);
        }
        try {
            String filetring = filepath.substring(filepath.lastIndexOf("."));
            InputStream instream = new FileInputStream(filepath);
            //HSSFWorkbook相当于一个excel文件，HSSFWorkbook是解析excel2007之前的版本（xls）
            //之后版本使用XSSFWorkbook（xlsx）
            if (".xls".equals(filetring)) {
                workbook = new HSSFWorkbook(instream);
            } else if (".xlsx".equals(filetring)) {
                workbook = new XSSFWorkbook(instream);
            }
            if (workbook != null) {
                for (int sheetat = 0; sheetat < workbook.getNumberOfSheets(); sheetat++) {
                    //获得sheet工作簿
                    Sheet sheet = workbook.getSheetAt(sheetat);
                    for (int rownum = 0; rownum < sheet.getPhysicalNumberOfRows(); rownum++) {
                        Row row = sheet.getRow(rownum);
                        for (int colnumindex = 0; colnumindex < row.getLastCellNum(); colnumindex++) {
                            String value = getvalue(sheet, rownum, colnumindex);
                            System.out.println(value);
                        }
                    }
                }
            }
            instream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
