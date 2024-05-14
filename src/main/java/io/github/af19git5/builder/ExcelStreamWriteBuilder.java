package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelStreamCell;
import io.github.af19git5.entity.ExcelStreamStyle;
import io.github.af19git5.exception.ExcelException;

import lombok.NonNull;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * excel寫出建構器(資料流輸出)
 *
 * @author Jimmy Kang
 */
public class ExcelStreamWriteBuilder implements AutoCloseable {

    private final SXSSFWorkbook workbook;
    private final Map<String, SXSSFSheet> sheetMap;
    private final Map<String, List<ExcelStreamCell>> cellMap;
    private final Map<ExcelStreamStyle, CellStyle> cellStyleMap = new HashMap<>();

    public ExcelStreamWriteBuilder() {
        workbook = new SXSSFWorkbook();
        sheetMap = new LinkedHashMap<>();
        cellMap = new LinkedHashMap<>();
    }

    /**
     * 建立工作表
     *
     * @param sheetCode 工作表代碼
     * @param name 工作表名稱
     * @return 原方法
     */
    public ExcelStreamWriteBuilder createSheet(@NonNull String sheetCode, @NonNull String name) {
        sheetMap.put(sheetCode, workbook.createSheet(name));
        cellMap.put(sheetCode, new ArrayList<>());
        return this;
    }

    /**
     * 增加表格欄位合併規則
     *
     * @param sheetCode 工作表代碼
     * @param cellRangeAddresses 欄位合併規則
     * @return 原方法
     */
    public ExcelStreamWriteBuilder addCellRangeAddress(
            @NonNull String sheetCode, @NonNull CellRangeAddress... cellRangeAddresses) {
        SXSSFSheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        for (CellRangeAddress cellAddresses : cellRangeAddresses) {
            sheet.addMergedRegionUnsafe(cellAddresses);
        }
        return this;
    }

    /**
     * 增加表格欄位合併規則
     *
     * @param sheetCode 工作表代碼
     * @param cellRangeAddressList 欄位合併規則
     * @return 原方法
     */
    public ExcelStreamWriteBuilder addCellRangeAddress(
            @NonNull String sheetCode, @NonNull List<CellRangeAddress> cellRangeAddressList) {
        return addCellRangeAddress(
                sheetCode, cellRangeAddressList.toArray(new CellRangeAddress[0]));
    }

    /**
     * 增加隱藏列
     *
     * @param sheetCode 工作表代碼
     * @param hiddenRowNum 隱藏列
     * @return 原方法
     */
    public ExcelStreamWriteBuilder addHiddenRowNum(@NonNull String sheetCode, int hiddenRowNum) {
        SXSSFSheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        sheet.getRow(hiddenRowNum).setZeroHeight(true);
        return this;
    }

    /**
     * 增加隱藏行
     *
     * @param sheetCode 工作表代碼
     * @param hiddenColumnNum 隱藏行
     * @return 原方法
     */
    public ExcelStreamWriteBuilder addHiddenColumnNum(
            @NonNull String sheetCode, int hiddenColumnNum) {
        SXSSFSheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        sheet.setColumnHidden(hiddenColumnNum, true);
        return this;
    }

    /**
     * 保護工作表
     *
     * @param sheetCode 工作表代碼
     * @param password 密碼
     * @return 原方法
     */
    public ExcelStreamWriteBuilder protectSheet(
            @NonNull String sheetCode, @NonNull String password) {
        SXSSFSheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        sheet.protectSheet(password);
        return this;
    }

    /**
     * 新增欄位資料
     *
     * @param sheetCode 工作表代碼
     * @param cells 欄位資料
     * @return 原方法
     */
    public ExcelStreamWriteBuilder addCell(
            @NonNull String sheetCode, @NonNull ExcelStreamCell... cells) {
        List<ExcelStreamCell> cellList = cellMap.get(sheetCode);
        if (cellList == null) {
            return this;
        }
        cellList.addAll(Arrays.asList(cells));
        return this;
    }

    /**
     * 新增欄位資料
     *
     * @param sheetCode 工作表代碼
     * @param cellList 欄位資料
     * @return 原方法
     */
    public ExcelStreamWriteBuilder addCell(
            @NonNull String sheetCode, @NonNull List<ExcelStreamCell> cellList) {
        return addCell(sheetCode, cellList.toArray(new ExcelStreamCell[0]));
    }

    /**
     * 寫入資料流中
     *
     * @return 原方法
     */
    public ExcelStreamWriteBuilder flush(@NonNull String sheetCode) throws ExcelException {
        SXSSFSheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        List<ExcelStreamCell> cellList = cellMap.get(sheetCode);
        if (cellList == null) {
            return this;
        }
        Map<Integer, SXSSFRow> rowMap = new HashMap<>();
        for (ExcelStreamCell cell : cellList) {
            SXSSFRow row;
            if (null == rowMap.get(cell.getRow())) {
                row = sheet.createRow(cell.getRow());
                rowMap.put(cell.getRow(), row);
            } else {
                row = rowMap.get(cell.getRow());
            }
            SXSSFCell sxssfCell = row.createCell(cell.getColumn());
            switch (cell.getCellType()) {
                case FORMULA:
                    sxssfCell.setCellFormula(cell.getValue());
                    break;
                case BOOLEAN:
                    sxssfCell.setCellValue(Boolean.parseBoolean(cell.getValue()));
                    break;
                case NUMERIC:
                    sxssfCell.setCellType(cell.getCellType());
                    try {
                        sxssfCell.setCellValue(Double.parseDouble(cell.getValue()));
                    } catch (NumberFormatException e) {
                        sxssfCell.setCellValue(0);
                    }
                default:
                    sxssfCell.setCellType(cell.getCellType());
                    sxssfCell.setCellValue(cell.getValue());
                    break;
            }
            if (null != cell.getStyle()) {
                CellStyle cellStyle = cellStyleMap.get(cell.getStyle());
                if (null == cellStyle) {
                    cellStyle = cell.getStyle().toCellStyle(workbook);
                    cellStyleMap.put(cell.getStyle(), cellStyle);
                }
                sxssfCell.setCellStyle(cellStyle);
            }
        }
        try {
            sheet.flushRows();
        } catch (IOException e) {
            throw new ExcelException(e);
        }
        cellList.clear();
        return this;
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
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    /** 關閉workbook */
    @Override
    public void close() throws ExcelException {
        try {
            workbook.close();
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }
}
