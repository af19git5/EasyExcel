package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelCell;
import io.github.af19git5.entity.ExcelSheet;
import io.github.af19git5.exception.ExcelException;

import lombok.NonNull;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * excel寫出建構器
 *
 * @author Jimmy Kang
 */
public class ExcelWriteBuilder {

    private final List<ExcelSheet> sheetList;

    public ExcelWriteBuilder() {
        this.sheetList = new ArrayList<>();
    }

    public ExcelWriteBuilder addSheet(@NonNull ExcelSheet sheet) {
        this.sheetList.add(sheet);
        return this;
    }

    public ExcelWriteBuilder addSheet(@NonNull ExcelSheet... sheets) {
        this.sheetList.addAll(Arrays.asList(sheets));
        return this;
    }

    public ExcelWriteBuilder addSheet(@NonNull List<@NonNull ExcelSheet> sheetList) {
        this.sheetList.addAll(sheetList);
        return this;
    }

    public ExcelWriteBuilder clearSheet() {
        this.sheetList.clear();
        return this;
    }

    /**
     * 建立HSSFWorkBook(xls)
     *
     * @return HSSFWorkbook
     */
    private HSSFWorkbook buildHSSFWorkbook() {
        // 新建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        for (ExcelSheet sheet : sheetList) {
            // 建立工作表
            HSSFSheet hssfSheet = workbook.createSheet(sheet.getName());
            Map<Integer, HSSFRow> rowMap = new HashMap<>();
            for (ExcelCell cell : sheet.getCellList()) {
                HSSFRow row;
                if (null == rowMap.get(cell.getRow())) {
                    row = hssfSheet.createRow(cell.getRow());
                    rowMap.put(cell.getRow(), row);
                } else {
                    row = rowMap.get(cell.getRow());
                }
                HSSFCell hssfCell = row.createCell(cell.getColumn());
                switch (cell.getCellType()) {
                    case FORMULA:
                        hssfCell.setCellFormula(cell.getValue());
                        break;
                    case BOOLEAN:
                        hssfCell.setCellValue(Boolean.parseBoolean(cell.getValue()));
                        break;
                    case NUMERIC:
                        hssfCell.setCellType(cell.getCellType());
                        try {
                            hssfCell.setCellValue(Double.parseDouble(cell.getValue()));
                        } catch (NumberFormatException e) {
                            hssfCell.setCellValue(0);
                        }
                    default:
                        hssfCell.setCellType(cell.getCellType());
                        hssfCell.setCellValue(cell.getValue());
                        break;
                }
                if (null != cell.getStyle()) {
                    hssfCell.setCellStyle(cell.getStyle().toHSSCellStyle(workbook));
                }
                hssfSheet.autoSizeColumn(cell.getColumn(), true);
            }
            // 處理表格欄位合併
            for (CellRangeAddress cellAddresses : sheet.getCellRangeAddressList()) {
                hssfSheet.addMergedRegionUnsafe(cellAddresses);
            }
        }
        return workbook;
    }

    /**
     * 建立XSSFWorkBook(xlsx)
     *
     * @return HSSFWorkbook
     */
    private XSSFWorkbook buildXSSFWorkbook() {
        // 新建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        for (ExcelSheet sheet : sheetList) {
            // 建立工作表
            XSSFSheet xssfSheet = workbook.createSheet(sheet.getName());
            Map<Integer, XSSFRow> rowMap = new HashMap<>();
            for (ExcelCell cell : sheet.getCellList()) {
                XSSFRow row;
                if (null == rowMap.get(cell.getRow())) {
                    row = xssfSheet.createRow(cell.getRow());
                    rowMap.put(cell.getRow(), row);
                } else {
                    row = rowMap.get(cell.getRow());
                }
                XSSFCell xssfCell = row.createCell(cell.getColumn());
                switch (cell.getCellType()) {
                    case FORMULA:
                        xssfCell.setCellFormula(cell.getValue());
                        break;
                    case BOOLEAN:
                        xssfCell.setCellValue(Boolean.parseBoolean(cell.getValue()));
                        break;
                    case NUMERIC:
                        xssfCell.setCellType(cell.getCellType());
                        try {
                            xssfCell.setCellValue(Double.parseDouble(cell.getValue()));
                        } catch (NumberFormatException e) {
                            xssfCell.setCellValue(0);
                        }
                    default:
                        xssfCell.setCellType(cell.getCellType());
                        xssfCell.setCellValue(cell.getValue());
                        break;
                }
                if (null != cell.getStyle()) {
                    xssfCell.setCellStyle(cell.getStyle().toXSSCellStyle(workbook));
                }
                xssfSheet.autoSizeColumn(cell.getColumn(), true);
            }
            // 處理表格欄位合併
            for (CellRangeAddress cellAddresses : sheet.getCellRangeAddressList()) {
                xssfSheet.addMergedRegionUnsafe(cellAddresses);
            }
        }
        return workbook;
    }

    /**
     * 輸出xls
     *
     * @return byte陣列
     */
    public byte[] outputXls() throws ExcelException {
        byte[] bytes;
        try (HSSFWorkbook workbook = buildHSSFWorkbook();
                ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            workbook.write(out);
            bytes = out.toByteArray();
        } catch (IOException e) {
            throw new ExcelException(e);
        }
        return bytes;
    }

    /**
     * 輸出xls
     *
     * @param filePath 儲存檔案位置
     */
    public void outputXls(String filePath) throws ExcelException {
        outputXls(new File(filePath));
    }

    /**
     * 輸出xls
     *
     * @param file 儲存檔案
     */
    public void outputXls(File file) throws ExcelException {
        try (HSSFWorkbook workbook = buildHSSFWorkbook();
                FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    /**
     * 輸出xlsx
     *
     * @return byte陣列
     */
    public byte[] outputXlsx() throws ExcelException {
        byte[] bytes;
        try (XSSFWorkbook workbook = buildXSSFWorkbook();
                ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            workbook.write(out);
            bytes = out.toByteArray();
        } catch (IOException e) {
            throw new ExcelException(e);
        }
        return bytes;
    }

    /**
     * 輸出xlsx
     *
     * @param filePath 儲存檔案位置
     */
    public void outputXlsx(String filePath) throws ExcelException {
        outputXlsx(new File(filePath));
    }

    /**
     * 輸出xlsx
     *
     * @param file 儲存檔案
     */
    public void outputXlsx(File file) throws ExcelException {
        try (XSSFWorkbook workbook = buildXSSFWorkbook();
                FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }
}
