package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelCell;
import io.github.af19git5.entity.ExcelMergedRegion;
import io.github.af19git5.entity.ExcelSheet;
import io.github.af19git5.entity.ExcelStyle;
import io.github.af19git5.exception.ExcelException;

import lombok.NonNull;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * excel寫出建構器
 *
 * @author Jimmy Kang
 */
public class ExcelWriteBuilder {

    private final List<ExcelSheet> sheetList;

    private final Map<ExcelStyle, HSSFCellStyle> hssfCellStyleMap = new HashMap<>();
    private final Map<ExcelStyle, XSSFCellStyle> xssfCellStyleMap = new HashMap<>();

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
        HSSFPalette palette = workbook.getCustomPalette();
        for (ExcelSheet sheet : sheetList) {
            // 建立工作表
            HSSFSheet hssfSheet = workbook.createSheet(sheet.getName());
            // 處理表格欄位合併
            for (ExcelMergedRegion mergedRegion : sheet.getMergedRegionList()) {
                if (mergedRegion.getFirstRow().equals(mergedRegion.getLastRow())
                        && mergedRegion.getFirstColumn().equals(mergedRegion.getLastColumn())) {
                    // 欄位合併只有一格時略過處理
                    continue;
                }
                CellRangeAddress cellAddresses =
                        new CellRangeAddress(
                                mergedRegion.getFirstRow(),
                                mergedRegion.getLastRow(),
                                mergedRegion.getFirstColumn(),
                                mergedRegion.getLastColumn());
                hssfSheet.addMergedRegion(cellAddresses);
            }
            // 處理欄位資料
            Map<Integer, HSSFRow> rowMap = new HashMap<>();
            int maxColumnNum = 0;
            for (ExcelCell cell : sheet.getCellList()) {
                HSSFRow row;
                if (null == rowMap.get(cell.getRow())) {
                    row = hssfSheet.createRow(cell.getRow());
                    rowMap.put(cell.getRow(), row);
                } else {
                    row = rowMap.get(cell.getRow());
                }
                maxColumnNum = Math.max(maxColumnNum, cell.getColumn());
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
                        break;
                    default:
                        hssfCell.setCellType(cell.getCellType());
                        hssfCell.setCellValue(cell.getValue());
                        break;
                }
                if (null != cell.getStyle()) {
                    HSSFCellStyle hssfCellStyle = hssfCellStyleMap.get(cell.getStyle());
                    if (null == hssfCellStyle) {
                        hssfCellStyle = cell.getStyle().toHSSCellStyle(workbook);
                        hssfCellStyleMap.put(cell.getStyle(), hssfCellStyle);
                    }
                    hssfCell.setCellStyle(hssfCellStyle);
                }
            }
            // 處理表格欄位合併邊線顏色
            for (ExcelMergedRegion mergedRegion : sheet.getMergedRegionList()) {
                HSSFCellStyle cellStyle = null;
                for (int rowNum = mergedRegion.getFirstRow();
                        rowNum <= mergedRegion.getLastRow();
                        rowNum++) {
                    HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                    if (null == hssfRow) {
                        hssfRow = hssfSheet.createRow(rowNum);
                    }
                    for (int columnNum = mergedRegion.getFirstColumn();
                            columnNum <= mergedRegion.getLastColumn();
                            columnNum++) {
                        HSSFCell hssfCell = hssfRow.getCell(columnNum);
                        if (null == hssfCell) {
                            hssfCell = hssfRow.createCell(columnNum);
                        }
                        if (rowNum == mergedRegion.getFirstRow()
                                && columnNum == mergedRegion.getFirstColumn()) {
                            cellStyle = workbook.createCellStyle();
                            cellStyle.cloneStyleFrom(hssfCell.getCellStyle());
                            if (!mergedRegion.getBorderTop().equals(BorderStyle.NONE)) {
                                cellStyle.setBorderTop(mergedRegion.getBorderTop());
                                if (null != mergedRegion.getBorderTopColor()) {
                                    Color rgbColor = Color.decode(mergedRegion.getBorderTopColor());
                                    HSSFColor color =
                                            palette.findSimilarColor(
                                                    (byte) rgbColor.getRed(),
                                                    (byte) rgbColor.getGreen(),
                                                    (byte) rgbColor.getBlue());
                                    cellStyle.setTopBorderColor(color.getIndex());
                                }
                            }
                            if (!mergedRegion.getBorderBottom().equals(BorderStyle.NONE)) {
                                cellStyle.setBorderBottom(mergedRegion.getBorderBottom());
                                if (null != mergedRegion.getBorderBottomColor()) {
                                    Color rgbColor =
                                            Color.decode(mergedRegion.getBorderBottomColor());
                                    HSSFColor color =
                                            palette.findSimilarColor(
                                                    (byte) rgbColor.getRed(),
                                                    (byte) rgbColor.getGreen(),
                                                    (byte) rgbColor.getBlue());
                                    cellStyle.setBottomBorderColor(color.getIndex());
                                }
                            }
                            if (!mergedRegion.getBorderLeft().equals(BorderStyle.NONE)) {
                                cellStyle.setBorderLeft(mergedRegion.getBorderLeft());
                                if (null != mergedRegion.getBorderLeftColor()) {
                                    Color rgbColor =
                                            Color.decode(mergedRegion.getBorderLeftColor());
                                    HSSFColor color =
                                            palette.findSimilarColor(
                                                    (byte) rgbColor.getRed(),
                                                    (byte) rgbColor.getGreen(),
                                                    (byte) rgbColor.getBlue());
                                    cellStyle.setLeftBorderColor(color.getIndex());
                                }
                            }
                            if (!mergedRegion.getBorderRight().equals(BorderStyle.NONE)) {
                                cellStyle.setBorderRight(mergedRegion.getBorderRight());
                                if (null != mergedRegion.getBorderRightColor()) {
                                    Color rgbColor =
                                            Color.decode(mergedRegion.getBorderRightColor());
                                    HSSFColor color =
                                            palette.findSimilarColor(
                                                    (byte) rgbColor.getRed(),
                                                    (byte) rgbColor.getGreen(),
                                                    (byte) rgbColor.getBlue());
                                    cellStyle.setRightBorderColor(color.getIndex());
                                }
                            }
                        }
                        if (null != cellStyle) {
                            hssfCell.setCellStyle(cellStyle);
                        }
                    }
                }
            }
            // 處理行列隱藏
            for (Integer rowNum : sheet.getHiddenRowNumSet()) {
                hssfSheet.getRow(rowNum).setZeroHeight(true);
            }
            for (Integer columnNum : sheet.getHiddenColumnNumSet()) {
                hssfSheet.setColumnHidden(columnNum, true);
            }
            // 處理自動適應寬度
            for (int columnNum = 0; columnNum <= maxColumnNum; columnNum++) {
                hssfSheet.autoSizeColumn(columnNum);
            }
            // 處理欄位覆寫寬度
            sheet.getOverrideColumnWidthMap().forEach(hssfSheet::setColumnWidth);
            // 處理滾動凍結欄位
            hssfSheet.createFreezePane(sheet.getFreezeColumnNum(), sheet.getFreezeRowNum());
            // 保護資料表
            if (sheet.getIsProtect()) {
                hssfSheet.protectSheet(sheet.getPassword());
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
            // 處理表格欄位合併
            for (ExcelMergedRegion mergedRegion : sheet.getMergedRegionList()) {
                if (mergedRegion.getFirstRow().equals(mergedRegion.getLastRow())
                        && mergedRegion.getFirstColumn().equals(mergedRegion.getLastColumn())) {
                    // 欄位合併只有一格時略過處理
                    continue;
                }
                CellRangeAddress cellAddresses =
                        new CellRangeAddress(
                                mergedRegion.getFirstRow(),
                                mergedRegion.getLastRow(),
                                mergedRegion.getFirstColumn(),
                                mergedRegion.getLastColumn());
                xssfSheet.addMergedRegionUnsafe(cellAddresses);
            }
            // 處理欄位資料
            Map<Integer, XSSFRow> rowMap = new HashMap<>();
            int maxColumnNum = 0;
            for (ExcelCell cell : sheet.getCellList()) {
                XSSFRow row;
                if (null == rowMap.get(cell.getRow())) {
                    row = xssfSheet.createRow(cell.getRow());
                    rowMap.put(cell.getRow(), row);
                } else {
                    row = rowMap.get(cell.getRow());
                }
                maxColumnNum = Math.max(maxColumnNum, cell.getColumn());
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
                        break;
                    default:
                        xssfCell.setCellType(cell.getCellType());
                        xssfCell.setCellValue(cell.getValue());
                        break;
                }
                if (null != cell.getStyle()) {
                    XSSFCellStyle xssfCellStyle = xssfCellStyleMap.get(cell.getStyle());
                    if (null == xssfCellStyle) {
                        xssfCellStyle = cell.getStyle().toXSSCellStyle(workbook);
                        xssfCellStyleMap.put(cell.getStyle(), xssfCellStyle);
                    }
                    xssfCell.setCellStyle(xssfCellStyle);
                }
            }

            // 處理表格欄位合併邊線顏色
            for (ExcelMergedRegion mergedRegion : sheet.getMergedRegionList()) {
                XSSFCellStyle cellStyle = null;
                for (int rowNum = mergedRegion.getFirstRow();
                        rowNum <= mergedRegion.getLastRow();
                        rowNum++) {
                    XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                    if (null == xssfRow) {
                        xssfRow = xssfSheet.createRow(rowNum);
                    }
                    for (int columnNum = mergedRegion.getFirstColumn();
                            columnNum <= mergedRegion.getLastColumn();
                            columnNum++) {
                        XSSFCell xssfCell = xssfRow.getCell(columnNum);
                        if (null == xssfCell) {
                            xssfCell = xssfRow.createCell(columnNum);
                        }
                        if (rowNum == mergedRegion.getFirstRow()
                                && columnNum == mergedRegion.getFirstColumn()) {
                            cellStyle = xssfCell.getCellStyle().copy();
                            if (!mergedRegion.getBorderTop().equals(BorderStyle.NONE)) {
                                cellStyle.setBorderTop(mergedRegion.getBorderTop());
                                if (null != mergedRegion.getBorderTopColor()) {
                                    XSSFColor color =
                                            XSSFColor.from(
                                                    CTColor.Factory.newInstance(),
                                                    new DefaultIndexedColorMap());
                                    color.setARGBHex(mergedRegion.getBorderTopColor().substring(1));
                                    cellStyle.setTopBorderColor(color);
                                }
                            }
                            if (!mergedRegion.getBorderBottom().equals(BorderStyle.NONE)) {
                                cellStyle.setBorderBottom(mergedRegion.getBorderBottom());
                                if (null != mergedRegion.getBorderBottomColor()) {
                                    XSSFColor color =
                                            XSSFColor.from(
                                                    CTColor.Factory.newInstance(),
                                                    new DefaultIndexedColorMap());
                                    color.setARGBHex(
                                            mergedRegion.getBorderBottomColor().substring(1));
                                    cellStyle.setBottomBorderColor(color);
                                }
                            }
                            if (!mergedRegion.getBorderLeft().equals(BorderStyle.NONE)) {
                                cellStyle.setBorderLeft(mergedRegion.getBorderLeft());
                                if (null != mergedRegion.getBorderLeftColor()) {
                                    XSSFColor color =
                                            XSSFColor.from(
                                                    CTColor.Factory.newInstance(),
                                                    new DefaultIndexedColorMap());
                                    color.setARGBHex(
                                            mergedRegion.getBorderLeftColor().substring(1));
                                    cellStyle.setLeftBorderColor(color);
                                }
                            }
                            if (!mergedRegion.getBorderRight().equals(BorderStyle.NONE)) {
                                cellStyle.setBorderRight(mergedRegion.getBorderRight());
                                if (null != mergedRegion.getBorderRightColor()) {
                                    XSSFColor color =
                                            XSSFColor.from(
                                                    CTColor.Factory.newInstance(),
                                                    new DefaultIndexedColorMap());
                                    color.setARGBHex(
                                            mergedRegion.getBorderRightColor().substring(1));
                                    cellStyle.setRightBorderColor(color);
                                }
                            }
                        }
                        if (null != cellStyle) {
                            xssfCell.setCellStyle(cellStyle);
                        }
                    }
                }
            }
            // 處理行列隱藏
            for (Integer rowNum : sheet.getHiddenRowNumSet()) {
                xssfSheet.getRow(rowNum).setZeroHeight(true);
            }
            for (Integer columnNum : sheet.getHiddenColumnNumSet()) {
                xssfSheet.setColumnHidden(columnNum, true);
            }
            // 處理自動適應寬度
            for (int columnNum = 0; columnNum <= maxColumnNum; columnNum++) {
                xssfSheet.autoSizeColumn(columnNum);
            }
            // 處理欄位覆寫寬度
            sheet.getOverrideColumnWidthMap().forEach(xssfSheet::setColumnWidth);
            // 處理滾動凍結欄位
            xssfSheet.createFreezePane(sheet.getFreezeColumnNum(), sheet.getFreezeRowNum());
            // 保護資料表
            if (sheet.getIsProtect()) {
                xssfSheet.protectSheet(sheet.getPassword());
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
    public void outputXls(@NonNull String filePath) throws ExcelException {
        outputXls(new File(filePath));
    }

    /**
     * 輸出xls
     *
     * @param file 儲存檔案
     */
    public void outputXls(@NonNull File file) throws ExcelException {
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
    public void outputXlsx(@NonNull String filePath) throws ExcelException {
        outputXlsx(new File(filePath));
    }

    /**
     * 輸出xlsx
     *
     * @param file 儲存檔案
     */
    public void outputXlsx(@NonNull File file) throws ExcelException {
        try (XSSFWorkbook workbook = buildXSSFWorkbook();
                FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }
}
