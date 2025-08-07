package io.github.af19git5.builder;

import com.github.miachm.sods.*;

import io.github.af19git5.entity.ExcelStreamStyle;
import io.github.af19git5.entity.StreamCell;
import io.github.af19git5.entity.StreamMergedRegion;
import io.github.af19git5.exception.ExcelException;
import io.github.af19git5.utils.SodsUtils;

import lombok.NonNull;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.security.NoSuchAlgorithmException;
import java.util.*;

/**
 * ods寫出建構器(資料流輸出)
 *
 * @author Jimmy Kang
 */
public class OdsStreamWriteBuilder implements StreamWriteBuilder {

    private final SpreadSheet spreadSheet;
    private final Map<String, Sheet> sheetMap;
    private final Map<String, List<StreamCell>> cellMap;
    private final Map<ExcelStreamStyle, Style> cellStyleMap;
    private final File outputFile;

    public OdsStreamWriteBuilder() throws IOException {
        spreadSheet = new SpreadSheet();
        sheetMap = new HashMap<>();
        cellMap = new HashMap<>();
        cellStyleMap = new HashMap<>();
        outputFile = File.createTempFile("output", ".ods");
    }

    /**
     * 建立工作表
     *
     * @param sheetCode 工作表代碼
     * @param name 工作表名稱
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder createSheet(@NonNull String sheetCode, @NonNull String name) {
        sheetMap.put(sheetCode, new Sheet(name));
        spreadSheet.appendSheet(sheetMap.get(sheetCode));
        cellMap.put(sheetCode, new ArrayList<>());
        return this;
    }

    /**
     * 新增欄位資料
     *
     * @param sheetCode 工作表代碼
     * @param cells 欄位資料
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder cells(@NonNull String sheetCode, @NonNull StreamCell... cells) {
        List<StreamCell> cellList = cellMap.get(sheetCode);
        if (null == cellList) {
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
    @Override
    public OdsStreamWriteBuilder cellList(
            @NonNull String sheetCode, @NonNull List<StreamCell> cellList) {
        return cells(sheetCode, cellList.toArray(new StreamCell[0]));
    }

    /**
     * 增加表格欄位合併規則 FIXME: 後續須調整實務邏輯至flush()中
     *
     * @param sheetCode 工作表代碼
     * @param mergedRegions 欄位合併規則
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder mergedRegions(
            @NonNull String sheetCode, @NonNull StreamMergedRegion... mergedRegions) {
        Sheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        for (StreamMergedRegion mergedRegion : mergedRegions) {
            int numRows = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
            int numColumns = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
            Range mergeRange =
                    sheet.getRange(
                            mergedRegion.getFirstRow(),
                            mergedRegion.getFirstColumn(),
                            numRows,
                            numColumns);
            mergeRange.merge();

            // 處理表格欄位合併後樣式
            Borders borders = mergeRange.getStyle().getBorders();
            if (null != borders) {
                borders.setBorderTop(SodsUtils.isBorder(mergedRegion.getBorderTop()));
                borders.setBorderLeft(SodsUtils.isBorder(mergedRegion.getBorderLeft()));
                borders.setBorderBottom(SodsUtils.isBorder(mergedRegion.getBorderBottom()));
                borders.setBorderRight(SodsUtils.isBorder(mergedRegion.getBorderRight()));
                borders.setBorderTopProperties(
                        SodsUtils.getBorderProperties(mergedRegion, SodsUtils.BORDER_TYPE.TOP));
                borders.setBorderLeftProperties(
                        SodsUtils.getBorderProperties(mergedRegion, SodsUtils.BORDER_TYPE.LEFT));
                borders.setBorderBottomProperties(
                        SodsUtils.getBorderProperties(mergedRegion, SodsUtils.BORDER_TYPE.BOTTOM));
                borders.setBorderRightProperties(
                        SodsUtils.getBorderProperties(mergedRegion, SodsUtils.BORDER_TYPE.RIGHT));
            }
        }
        return this;
    }

    /**
     * 增加表格欄位合併規則
     *
     * @param sheetCode 工作表代碼
     * @param mergedRegions 欄位合併規則
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder mergedRegions(
            @NonNull String sheetCode, @NonNull List<StreamMergedRegion> mergedRegions) {
        return mergedRegions(sheetCode, mergedRegions.toArray(new StreamMergedRegion[0]));
    }

    /**
     * 增加隱藏列
     *
     * @param sheetCode 工作表代碼
     * @param rowNums 隱藏列
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder hiddenRowNums(
            @NonNull String sheetCode, @NonNull Integer... rowNums) {
        Sheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        for (Integer rowNum : rowNums) {
            sheet.hideRow(rowNum);
        }
        return this;
    }

    /**
     * 增加隱藏列
     *
     * @param sheetCode 工作表代碼
     * @param hiddenRowNumSet 隱藏列
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder hiddenRowNumSet(
            @NonNull String sheetCode, @NonNull Set<Integer> hiddenRowNumSet) {
        return hiddenRowNums(sheetCode, hiddenRowNumSet.toArray(new Integer[0]));
    }

    /**
     * 增加隱藏行
     *
     * @param sheetCode 工作表代碼
     * @param columnNums 隱藏行
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder hiddenColumnNums(
            @NonNull String sheetCode, @NonNull Integer... columnNums) {
        Sheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        for (Integer columnNum : columnNums) {
            sheet.hideColumn(columnNum);
        }
        return this;
    }

    /**
     * 增加隱藏行
     *
     * @param sheetCode 工作表代碼
     * @param hiddenColumnNumSet 隱藏列
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder hiddenColumnNumSet(
            @NonNull String sheetCode, @NonNull Set<Integer> hiddenColumnNumSet) {
        return hiddenColumnNums(sheetCode, hiddenColumnNumSet.toArray(new Integer[0]));
    }

    /**
     * 保護工作表
     *
     * @param sheetCode 工作表代碼
     * @param password 密碼
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder protectSheet(@NonNull String sheetCode, @NonNull String password)
            throws NoSuchAlgorithmException {
        Sheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        sheet.setPassword(password);
        return this;
    }

    /**
     * 覆寫欄位寬度
     *
     * @param sheetCode 工作表代碼
     * @param columnNum 欄位代碼
     * @param width 覆寫寬度
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder overrideColumnWidth(
            @NonNull String sheetCode, int columnNum, int width) {
        Sheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        sheet.setColumnWidth(columnNum, (double) width);
        return this;
    }

    /**
     * 寫入資料流中
     *
     * @return 原方法
     */
    @Override
    public OdsStreamWriteBuilder flush(@NonNull String sheetCode) throws IOException {
        Sheet sheet = sheetMap.get(sheetCode);
        if (null == sheet) {
            return this;
        }
        List<StreamCell> cellList = cellMap.get(sheetCode);
        if (null == cellList) {
            return this;
        }
        int maxRow = 0;
        int maxColumn = 0;
        for (StreamCell cell : cellList) {
            if (cell.getRow() > maxRow) {
                sheet.appendRows(cell.getRow() - maxRow);
                maxRow = cell.getRow();
            }
            if (cell.getColumn() > maxColumn) {
                sheet.appendColumns(cell.getColumn() - maxColumn);
                maxColumn = cell.getColumn();
            }
            Range range = sheet.getRange(cell.getRow(), cell.getColumn());
            range.setValue(cell.getValue());
            Style style = cellStyleMap.get(cell.getStyle());
            if (null == style) {
                style = SodsUtils.toSodsStyle(cell.getStyle());
                cellStyleMap.put(cell.getStyle(), style);
            }
            range.setStyle(style);
        }
        spreadSheet.save(outputFile);
        cellList.clear();
        return this;
    }

    /**
     * 輸出ods
     *
     * @param filePath 儲存檔案位置
     */
    @Override
    public void output(@NonNull String filePath) throws ExcelException {
        output(new File(filePath));
    }

    /**
     * 輸出ods
     *
     * @param file 儲存檔案
     */
    @Override
    public void output(@NonNull File file) throws ExcelException {
        try {
            Files.copy(outputFile.toPath(), file.toPath(), StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    /** 關閉workbook */
    @Override
    public void close() throws ExcelException {
        if (null != outputFile && outputFile.exists()) {
            try {
                Files.deleteIfExists(outputFile.toPath());
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }
}
