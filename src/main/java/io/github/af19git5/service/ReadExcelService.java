package io.github.af19git5.service;

import io.github.af19git5.entity.ExcelCell;
import io.github.af19git5.entity.ExcelMergedRegion;
import io.github.af19git5.entity.ExcelSheet;
import io.github.af19git5.entity.ExcelStyle;
import io.github.af19git5.exception.ExcelException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 讀取excel服務
 *
 * @author Jimmy Kang
 */
public class ReadExcelService {

    private final String dateTimeFormat = "yyyy-MM-dd HH:mm:ss";

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
     * @param excelFilePath excel檔案路徑
     * @param password 密碼
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public List<ExcelSheet> read(String excelFilePath, String password) throws ExcelException {
        return read(new File(excelFilePath), password);
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
            excelSheetList.addAll(read(workbook));
        } catch (IOException e) {
            throw new ExcelException(e.getMessage());
        }
        return excelSheetList;
    }

    /**
     * 讀取excel資料
     *
     * @param excelFile excel檔案
     * @param password 密碼
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public List<ExcelSheet> read(File excelFile, String password) throws ExcelException {
        List<ExcelSheet> excelSheetList = new ArrayList<>();
        try (Workbook workbook = WorkbookFactory.create(excelFile, password)) {
            excelSheetList.addAll(read(workbook));
        } catch (IOException e) {
            throw new ExcelException(e.getMessage());
        }
        return excelSheetList;
    }

    /**
     * 讀取excel資料
     *
     * @param inputStream InputStream
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public List<ExcelSheet> read(InputStream inputStream) throws ExcelException {
        List<ExcelSheet> excelSheetList = new ArrayList<>();
        try (Workbook workbook = WorkbookFactory.create(inputStream)) {
            excelSheetList.addAll(read(workbook));
        } catch (IOException e) {
            throw new ExcelException(e.getMessage());
        }
        return excelSheetList;
    }

    /**
     * 讀取excel資料
     *
     * @param inputStream InputStream
     * @param password 密碼
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public List<ExcelSheet> read(InputStream inputStream, String password) throws ExcelException {
        List<ExcelSheet> excelSheetList = new ArrayList<>();
        try (Workbook workbook = WorkbookFactory.create(inputStream, password)) {
            excelSheetList.addAll(read(workbook));
        } catch (IOException e) {
            throw new ExcelException(e.getMessage());
        }
        return excelSheetList;
    }

    /**
     * 讀取excel資料
     *
     * @param workbook Workbook
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    private List<ExcelSheet> read(Workbook workbook) throws ExcelException {
        List<ExcelSheet> excelSheetList = new ArrayList<>();
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
        excelSheet.setMergedRegionList(new ArrayList<>());
        for (CellRangeAddress cellAddresses : sheet.getMergedRegions()) {
            excelSheet
                    .getMergedRegionList()
                    .add(
                            new ExcelMergedRegion(
                                    cellAddresses.getFirstRow(),
                                    cellAddresses.getLastRow(),
                                    cellAddresses.getFirstColumn(),
                                    cellAddresses.getLastColumn()));
        }
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            HSSFRow row = sheet.getRow(rowNum);
            if (null == row) continue;
            for (int columnNum = 0; columnNum < row.getLastCellNum(); columnNum++) {
                HSSFCell cell = row.getCell(columnNum);
                if (cell == null) {
                    excelSheet
                            .getCellList()
                            .add(new ExcelCell("", rowNum, columnNum, CellType.STRING));
                } else {
                    excelSheet
                            .getCellList()
                            .add(
                                    new ExcelCell(
                                            readHSSFCellValue(cell),
                                            rowNum,
                                            columnNum,
                                            cell.getCellType(),
                                            new ExcelStyle(workbook, cell.getCellStyle())));
                }
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
        excelSheet.setMergedRegionList(new ArrayList<>());
        for (CellRangeAddress cellAddresses : sheet.getMergedRegions()) {
            excelSheet
                    .getMergedRegionList()
                    .add(
                            new ExcelMergedRegion(
                                    cellAddresses.getFirstRow(),
                                    cellAddresses.getLastRow(),
                                    cellAddresses.getFirstColumn(),
                                    cellAddresses.getLastColumn()));
        }
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            XSSFRow row = sheet.getRow(rowNum);
            if (null == row) continue;
            for (int columnNum = 0; columnNum < row.getLastCellNum(); columnNum++) {
                XSSFCell cell = row.getCell(columnNum);
                if (cell == null) {
                    excelSheet
                            .getCellList()
                            .add(new ExcelCell("", rowNum, columnNum, CellType.STRING));
                } else {
                    excelSheet
                            .getCellList()
                            .add(
                                    new ExcelCell(
                                            readXSSFCellValue(cell),
                                            rowNum,
                                            columnNum,
                                            cell.getCellType(),
                                            new ExcelStyle(cell.getCellStyle())));
                }
            }
        }
        return excelSheet;
    }

    /**
     * 讀取XSSFCell數值
     *
     * @param cell XSSFCell
     * @return 數值
     */
    private String readHSSFCellValue(HSSFCell cell) {
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            // 判斷欄位是否為日期格式
            Date date = cell.getDateCellValue();
            return new SimpleDateFormat(dateTimeFormat).format(date);
        } else {
            return cell.toString();
        }
    }

    /**
     * 讀取XSSFCell數值
     *
     * @param cell XSSFCell
     * @return 數值
     */
    private String readXSSFCellValue(XSSFCell cell) {
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            // 判斷欄位是否為日期格式
            Date date = cell.getDateCellValue();
            return new SimpleDateFormat(dateTimeFormat).format(date);
        } else {
            return cell.toString();
        }
    }
}
