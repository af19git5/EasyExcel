package io.github.af19git5.builder;


import io.github.af19git5.entity.ExcelStreamCell;
import io.github.af19git5.exception.ExcelException;

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
    public ExcelStreamWriteBuilder createSheet(String sheetCode, String name) {
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
            String sheetCode, CellRangeAddress... cellRangeAddresses) {
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
            String sheetCode, List<CellRangeAddress> cellRangeAddressList) {
        return addCellRangeAddress(
                sheetCode, cellRangeAddressList.toArray(new CellRangeAddress[0]));
    }

    /**
     * 新增欄位資料
     *
     * @param sheetCode 工作表代碼
     * @param cells 欄位資料
     * @return 原方法
     */
    public ExcelStreamWriteBuilder addCell(String sheetCode, ExcelStreamCell... cells) {
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
    public ExcelStreamWriteBuilder addCell(String sheetCode, List<ExcelStreamCell> cellList) {
        return addCell(sheetCode, cellList.toArray(new ExcelStreamCell[0]));
    }

    /**
     * 寫入資料流中
     *
     * @return 原方法
     */
    public ExcelStreamWriteBuilder flush(String sheetCode) throws ExcelException {
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
                sxssfCell.setCellStyle(cell.getStyle().toCellStyle(workbook));
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
    public void outputXlsx(String filePath) throws ExcelException {
        outputXlsx(new File(filePath));
    }

    /**
     * 輸出xlsx
     *
     * @param file 儲存檔案
     */
    public void outputXlsx(File file) throws ExcelException {
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
