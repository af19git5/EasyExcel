package io.github.af19git5.service;

import io.github.af19git5.entity.ExcelCell;
import io.github.af19git5.entity.ExcelSheet;
import io.github.af19git5.entity.ExcelStyle;
import io.github.af19git5.exception.ExcelException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 讀取excel服務
 *
 * @author Jimmy Kang
 */
public class ReadExcelService {

    /**
     * 讀取excel資料
     *
     * @param excelFilePath excel檔案路徑
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public List<ExcelSheet> read(String excelFilePath) throws ExcelException {
        return read(new File(excelFilePath));
    }

    /**
     * 讀取excel資料
     *
     * @param excelFile excel檔案
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public List<ExcelSheet> read(File excelFile) throws ExcelException {
        List<ExcelSheet> excelSheetList = new ArrayList<>();
        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            if (workbook instanceof XSSFWorkbook) {
                XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
                for (int sheetNum = 0; sheetNum < xssfWorkbook.getNumberOfSheets(); sheetNum++) {
                    XSSFSheet sheet = xssfWorkbook.getSheetAt(sheetNum);
                    excelSheetList.add(readSheet(xssfWorkbook, sheet));
                }
            } else if (workbook instanceof HSSFWorkbook) {
                HSSFWorkbook hssfWorkbook = (HSSFWorkbook) workbook;
                for (int sheetNum = 0; sheetNum < hssfWorkbook.getNumberOfSheets(); sheetNum++) {
                    HSSFSheet sheet = hssfWorkbook.getSheetAt(sheetNum);
                    excelSheetList.add(readSheet(hssfWorkbook, sheet));
                }
            } else {
                throw new ExcelException("File is not excel.");
            }
        } catch (IOException e) {
            throw new ExcelException(e.getMessage());
        }
        return excelSheetList;
    }

    /**
     * 讀取工作表資料
     *
     * @param workbook excel資料
     * @param sheet 工作表資料
     * @return 工作表資料
     */
    private ExcelSheet readSheet(HSSFWorkbook workbook, HSSFSheet sheet) {
        ExcelSheet excelSheet = new ExcelSheet(sheet.getSheetName(), new ArrayList<>());
        excelSheet.setCellRangeAddressList(sheet.getMergedRegions());
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            HSSFRow row = sheet.getRow(rowNum);
            if (null == row) continue;
            for (int columnNum = 0; columnNum < row.getLastCellNum(); columnNum++) {
                HSSFCell cell = row.getCell(columnNum);
                if (cell == null) continue;
                excelSheet
                        .getCellList()
                        .add(
                                new ExcelCell(
                                        cell.toString(),
                                        rowNum,
                                        columnNum,
                                        cell.getCellType(),
                                        new ExcelStyle(workbook, cell.getCellStyle())));
            }
        }
        return excelSheet;
    }

    /**
     * 讀取工作表資料
     *
     * @param workbook excel資料
     * @param sheet 工作表資料
     * @return 工作表資料
     */
    private ExcelSheet readSheet(XSSFWorkbook workbook, XSSFSheet sheet) {
        ExcelSheet excelSheet = new ExcelSheet(sheet.getSheetName(), new ArrayList<>());
        excelSheet.setCellRangeAddressList(sheet.getMergedRegions());
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            XSSFRow row = sheet.getRow(rowNum);
            if (null == row) continue;
            for (int columnNum = 0; columnNum < row.getLastCellNum(); columnNum++) {
                XSSFCell cell = row.getCell(columnNum);
                if (cell == null) continue;
                excelSheet
                        .getCellList()
                        .add(
                                new ExcelCell(
                                        cell.toString(),
                                        rowNum,
                                        columnNum,
                                        cell.getCellType(),
                                        new ExcelStyle(cell.getCellStyle())));
            }
        }
        return excelSheet;
    }
}
