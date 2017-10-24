package com.github.doghere.excel;

import java.io.*;

import java.util.Iterator;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class MergeExcel {

    public static void main(String[]args) throws IOException {
        File dir = new File("job/佣金结算/out");
        //将所有类型的尽调excel文件合并成一个excel文件
        XSSFWorkbook newExcelCreat = new XSSFWorkbook();
        for (File file1 : dir.listFiles(new FileFilter() {
            @Override
            public boolean accept(File file) {
                return file.getName().endsWith(".xlsx");
            }
        })) {
            String s = file1.getAbsolutePath() ;
//遍历每个源excel文件，fileNameList为源文件的名称集合
            InputStream in = new FileInputStream(s);
            XSSFWorkbook fromExcel = null;
            try {
                fromExcel = new XSSFWorkbook(in);
            } catch (IOException e) {

            }
            for (int i = 0; i < fromExcel.getNumberOfSheets(); i++) {//遍历每个sheet
                XSSFSheet oldSheet = fromExcel.getSheetAt(i);
                XSSFSheet newSheet = newExcelCreat.createSheet(oldSheet.getSheetName());
                MergeExcel.copySheet(newExcelCreat, oldSheet, newSheet);
            }
        }

        String allFileName="/tmp/qwe.xlsx";
        FileOutputStream fileOut = new FileOutputStream(allFileName);
        newExcelCreat.write(fileOut);
        fileOut.flush();
        fileOut.close();

    }


    public static void mergeExcel(String toFile, String ... fromFile) throws IOException {
        //将所有类型的尽调excel文件合并成一个excel文件
        XSSFWorkbook newExcelCreat = new XSSFWorkbook();
        for (String file : fromFile) {
            //遍历每个源excel文件，fileNameList为源文件的名称集合
            InputStream in = new FileInputStream(file);
            XSSFWorkbook fromExcel = null;
            try {
                fromExcel = new XSSFWorkbook(in);
            } catch (IOException e) {
                e.printStackTrace();
            }
            assert fromExcel != null;
            for (int i = 0; i < fromExcel.getNumberOfSheets(); i++) {//遍历每个sheet
                XSSFSheet oldSheet = fromExcel.getSheetAt(i);
                XSSFSheet newSheet = newExcelCreat.createSheet(oldSheet.getSheetName());
                MergeExcel.copySheet(newExcelCreat, oldSheet, newSheet);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(toFile);
        newExcelCreat.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

    private static void copyCellStyle(XSSFCellStyle fromStyle, XSSFCellStyle toStyle) {

        toStyle.cloneStyleFrom(fromStyle);//此一行代码搞定

    }
    private static void mergeSheetAllRegion(XSSFSheet fromSheet, XSSFSheet toSheet) {//合并单元格
        int num = fromSheet.getNumMergedRegions();
        CellRangeAddress cellR = null;
        for (int i = 0; i < num; i++) {
            cellR = fromSheet.getMergedRegion(i);
            toSheet.addMergedRegion(cellR);
        }
    }

    private static void copyCell(XSSFWorkbook wb, XSSFCell fromCell, XSSFCell toCell) {
        XSSFCellStyle newstyle=wb.createCellStyle();
        copyCellStyle(fromCell.getCellStyle(), newstyle);
        //toCell.setEncoding(fromCell.getEncoding());
        //样式
        toCell.setCellStyle(newstyle);
        if (fromCell.getCellComment() != null) {
            toCell.setCellComment(fromCell.getCellComment());
        }
        // 不同数据类型处理
        int fromCellType = fromCell.getCellType();
        toCell.setCellType(fromCellType);
        if (fromCellType == XSSFCell.CELL_TYPE_NUMERIC) {
            if (DateUtil.isCellDateFormatted(fromCell)) {
                toCell.setCellValue(fromCell.getDateCellValue());
            } else {
                toCell.setCellValue(fromCell.getNumericCellValue());
            }
        } else if (fromCellType == XSSFCell.CELL_TYPE_STRING) {
            toCell.setCellValue(fromCell.getRichStringCellValue());
        } else if (fromCellType == XSSFCell.CELL_TYPE_BLANK) {
            // nothing21
        } else if (fromCellType == XSSFCell.CELL_TYPE_BOOLEAN) {
            toCell.setCellValue(fromCell.getBooleanCellValue());
        } else if (fromCellType == XSSFCell.CELL_TYPE_ERROR) {
            toCell.setCellErrorValue(fromCell.getErrorCellValue());
        } else if (fromCellType == XSSFCell.CELL_TYPE_FORMULA) {
            toCell.setCellFormula(fromCell.getCellFormula());
        } else { // nothing29
        }

    }

    private static void copyRow(XSSFWorkbook wb, XSSFRow oldRow, XSSFRow toRow){
        toRow.setHeight(oldRow.getHeight());
        for (Iterator cellIt = oldRow.cellIterator(); cellIt.hasNext();) {
            XSSFCell tmpCell = (XSSFCell) cellIt.next();
            XSSFCell newCell = toRow.createCell(tmpCell.getColumnIndex());
            copyCell(wb,tmpCell, newCell);
        }
    }
    private static void copySheet(XSSFWorkbook wb, XSSFSheet fromSheet, XSSFSheet toSheet) {
        mergeSheetAllRegion(fromSheet, toSheet);
        //设置列宽
        for(int i=0;i<=fromSheet.getRow(fromSheet.getFirstRowNum()).getLastCellNum();i++){
            toSheet.setColumnWidth(i,fromSheet.getColumnWidth(i));
        }
        for (Iterator rowIt = fromSheet.rowIterator(); rowIt.hasNext();) {
            XSSFRow oldRow = (XSSFRow) rowIt.next();
            XSSFRow newRow = toSheet.createRow(oldRow.getRowNum());
            copyRow(wb,oldRow,newRow);
        }
    }


}