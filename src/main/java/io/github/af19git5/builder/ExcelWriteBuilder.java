package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelCell;
import io.github.af19git5.entity.ExcelMergedRegion;
import io.github.af19git5.entity.ExcelSheet;
import io.github.af19git5.entity.ExcelStyle;
import io.github.af19git5.exception.ExcelException;
import io.github.af19git5.type.ExcelType;
import io.github.af19git5.utils.ExcelColorUtils;
import io.github.af19git5.utils.HSSFColorPicker;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import lombok.Getter;
import lombok.NonNull;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

/**
 * excel寫出建構器
 *
 * <p>xlsx使用資料流(SXSSF)輸出，呼叫{@link #flush(String)}會將已加入的欄位寫出並釋放記憶體；
 * xls整份保留在記憶體中。輸出時會自動flush尚未寫出的資料，不呼叫flush也能正確輸出。
 *
 * @author Jimmy Kang
 */
public class ExcelWriteBuilder implements AutoCloseable {

    /** 讀取Excel時日期欄位的文字格式，寫出NUMERIC欄位時可辨識 */
    private static final DateTimeFormatter DATE_TIME_FORMATTER =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    /**
     * -- GETTER -- 取得輸出檔案類型
     *
     * @return 檔案類型
     */
    @Getter private final ExcelType type;

    private final Workbook workbook;
    private final HSSFColorPicker colorPicker;
    private final Map<String, SheetContext> sheetMap = new LinkedHashMap<>();
    private final Map<ExcelStyle, CellStyle> cellStyleMap = new HashMap<>();
    private final Map<String, CellStyle> mergedCellStyleMap = new HashMap<>();

    /** xls延後至輸出時才寫入的公式(確保參照的工作表皆已建立) */
    private final List<PendingFormula> pendingFormulaList = new ArrayList<>();

    /** 是否強制開檔時重新計算公式，null時依是否寫過公式自動判斷 */
    private Boolean forceFormulaRecalculation;

    private boolean hasFormula = false;
    private boolean finished = false;

    public ExcelWriteBuilder() {
        this(ExcelType.XLSX);
    }

    public ExcelWriteBuilder(@NonNull ExcelType type) {
        this.type = type;
        if (ExcelType.XLS.equals(type)) {
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
            this.workbook = hssfWorkbook;
            this.colorPicker = new HSSFColorPicker(hssfWorkbook);
        } else {
            // 只在呼叫flush時寫出資料列
            this.workbook = new SXSSFWorkbook(-1);
            this.colorPicker = null;
        }
    }

    /**
     * 建立工作表
     *
     * @param sheetCode 工作表代碼
     * @param name 工作表名稱
     * @return 原方法
     */
    public ExcelWriteBuilder createSheet(@NonNull String sheetCode, @NonNull String name) {
        checkNotFinished();
        sheetMap.put(sheetCode, new SheetContext(workbook.createSheet(name)));
        return this;
    }

    /**
     * 加入整份工作表資料，工作表代碼為工作表名稱
     *
     * @param sheet 工作表資料
     * @return 原方法
     */
    public ExcelWriteBuilder sheet(@NonNull ExcelSheet sheet) {
        String sheetCode = sheet.getName();
        createSheet(sheetCode, sheet.getName());
        cellList(sheetCode, sheet.getCellList());
        mergedRegions(sheetCode, sheet.getMergedRegionList());
        hiddenRowNumSet(sheetCode, sheet.getHiddenRowNumSet());
        hiddenColumnNumSet(sheetCode, sheet.getHiddenColumnNumSet());
        sheet.getOverrideColumnWidthMap()
                .forEach((columnNum, width) -> overrideColumnWidth(sheetCode, columnNum, width));
        if (sheet.getFreezeColumnNum() > 0 || sheet.getFreezeRowNum() > 0) {
            freezePane(sheetCode, sheet.getFreezeColumnNum(), sheet.getFreezeRowNum());
        }
        if (sheet.getIsProtect()) {
            protectSheet(sheetCode, sheet.getPassword());
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
    public ExcelWriteBuilder mergedRegions(
            @NonNull String sheetCode, @NonNull ExcelMergedRegion... mergedRegions) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        for (ExcelMergedRegion mergedRegion : mergedRegions) {
            if (mergedRegion.getFirstRow().equals(mergedRegion.getLastRow())
                    && mergedRegion.getFirstColumn().equals(mergedRegion.getLastColumn())) {
                // 欄位合併只有一格時略過處理
                continue;
            }
            context.mergedRegionList.add(mergedRegion);
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
    public ExcelWriteBuilder mergedRegions(
            @NonNull String sheetCode, @NonNull List<ExcelMergedRegion> mergedRegions) {
        return mergedRegions(sheetCode, mergedRegions.toArray(new ExcelMergedRegion[0]));
    }

    /**
     * 增加隱藏列
     *
     * @param sheetCode 工作表代碼
     * @param rowNums 隱藏列
     * @return 原方法
     */
    public ExcelWriteBuilder hiddenRowNums(@NonNull String sheetCode, @NonNull Integer... rowNums) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        for (Integer rowNum : rowNums) {
            checkRowNotFlushed(context, rowNum);
            context.hiddenRowNumSet.add(rowNum);
            Row row = context.sheet.getRow(rowNum);
            if (null != row) {
                row.setZeroHeight(true);
            }
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
    public ExcelWriteBuilder hiddenRowNumSet(
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
    public ExcelWriteBuilder hiddenColumnNums(
            @NonNull String sheetCode, @NonNull Integer... columnNums) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        for (Integer columnNum : columnNums) {
            context.sheet.setColumnHidden(columnNum, true);
        }
        return this;
    }

    /**
     * 增加隱藏行
     *
     * @param sheetCode 工作表代碼
     * @param hiddenColumnNumSet 隱藏行
     * @return 原方法
     */
    public ExcelWriteBuilder hiddenColumnNumSet(
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
    public ExcelWriteBuilder protectSheet(@NonNull String sheetCode, @NonNull String password) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        context.sheet.protectSheet(password);
        return this;
    }

    /**
     * 覆寫欄位寬度
     *
     * @param sheetCode 工作表代碼
     * @param columnNum 欄位代碼
     * @param width 覆寫寬度(單位為1/256個字元寬)
     * @return 原方法
     */
    public ExcelWriteBuilder overrideColumnWidth(
            @NonNull String sheetCode, int columnNum, int width) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        context.overrideColumnWidthMap.put(columnNum, width);
        context.sheet.setColumnWidth(columnNum, width);
        return this;
    }

    /**
     * 自動調整欄位寬度(輸出時依內容計算，覆寫欄位寬度優先)
     *
     * <p>xlsx只會計算設定之後才寫出的資料，請在第一次flush前設定
     *
     * @param sheetCode 工作表代碼
     * @param columnNums 欄位代碼
     * @return 原方法
     */
    public ExcelWriteBuilder autoSizeColumns(
            @NonNull String sheetCode, @NonNull Integer... columnNums) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        for (Integer columnNum : columnNums) {
            context.autoSizeColumnNumSet.add(columnNum);
            if (context.sheet instanceof SXSSFSheet) {
                ((SXSSFSheet) context.sheet).trackColumnForAutoSizing(columnNum);
            }
        }
        return this;
    }

    /**
     * 凍結窗格
     *
     * @param sheetCode 工作表代碼
     * @param columnNum 凍結行數
     * @param rowNum 凍結列數
     * @return 原方法
     */
    public ExcelWriteBuilder freezePane(@NonNull String sheetCode, int columnNum, int rowNum) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        context.sheet.createFreezePane(columnNum, rowNum);
        return this;
    }

    /**
     * 設定列高
     *
     * @param sheetCode 工作表代碼
     * @param rowNum 列(從0開始)
     * @param heightInPoints 列高(點)
     * @return 原方法
     */
    public ExcelWriteBuilder rowHeight(
            @NonNull String sheetCode, int rowNum, @NonNull Float heightInPoints) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        checkRowNotFlushed(context, rowNum);
        context.rowHeightMap.put(rowNum, heightInPoints);
        Row row = context.sheet.getRow(rowNum);
        if (null != row) {
            row.setHeightInPoints(heightInPoints);
        }
        return this;
    }

    /**
     * 設定預設列高
     *
     * @param sheetCode 工作表代碼
     * @param heightInPoints 列高(點)
     * @return 原方法
     */
    public ExcelWriteBuilder defaultRowHeight(
            @NonNull String sheetCode, @NonNull Float heightInPoints) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        context.sheet.setDefaultRowHeightInPoints(heightInPoints);
        return this;
    }

    /**
     * 設定顯示比例
     *
     * @param sheetCode 工作表代碼
     * @param percent 顯示比例(10~400)
     * @return 原方法
     */
    public ExcelWriteBuilder zoom(@NonNull String sheetCode, int percent) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        context.sheet.setZoom(percent);
        return this;
    }

    /**
     * 設定開檔時是否強制重新計算公式，未設定時寫過公式即自動開啟
     *
     * @param forceFormulaRecalculation 是否強制重新計算
     * @return 原方法
     */
    public ExcelWriteBuilder forceFormulaRecalculation(boolean forceFormulaRecalculation) {
        checkNotFinished();
        this.forceFormulaRecalculation = forceFormulaRecalculation;
        return this;
    }

    /**
     * 新增欄位資料
     *
     * @param sheetCode 工作表代碼
     * @param cells 欄位資料
     * @return 原方法
     */
    public ExcelWriteBuilder cells(@NonNull String sheetCode, @NonNull ExcelCell... cells) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        context.cellList.addAll(Arrays.asList(cells));
        return this;
    }

    /**
     * 新增欄位資料
     *
     * @param sheetCode 工作表代碼
     * @param cellList 欄位資料
     * @return 原方法
     */
    public ExcelWriteBuilder cellList(
            @NonNull String sheetCode, @NonNull List<ExcelCell> cellList) {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        context.cellList.addAll(cellList);
        return this;
    }

    /**
     * 將已加入的欄位資料寫出(xlsx會寫入暫存檔並釋放記憶體)
     *
     * <p>xlsx寫出後的列無法再修改，之後加入的欄位列數須大於已寫出的列
     *
     * @param sheetCode 工作表代碼
     * @return 原方法
     * @throws ExcelException Excel處理錯誤
     */
    public ExcelWriteBuilder flush(@NonNull String sheetCode) throws ExcelException {
        SheetContext context = getSheetContext(sheetCode);
        if (null == context) {
            return this;
        }
        flushSheet(context, false);
        return this;
    }

    /**
     * 輸出Excel
     *
     * @param filePath 儲存檔案位置
     * @throws ExcelException Excel處理錯誤
     */
    public void output(@NonNull String filePath) throws ExcelException {
        output(new File(filePath));
    }

    /**
     * 輸出Excel
     *
     * @param file 儲存檔案
     * @throws ExcelException Excel處理錯誤
     */
    public void output(@NonNull File file) throws ExcelException {
        // 先完成寫入，避免資料錯誤時留下空檔案
        finish();
        try (FileOutputStream fos = new FileOutputStream(file)) {
            output(fos);
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    /**
     * 輸出Excel
     *
     * @param outputStream 輸出資料流(不會自動關閉)
     * @throws ExcelException Excel處理錯誤
     */
    public void output(@NonNull OutputStream outputStream) throws ExcelException {
        finish();
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    /**
     * 輸出Excel
     *
     * @return byte陣列
     * @throws ExcelException Excel處理錯誤
     */
    public byte[] outputBytes() throws ExcelException {
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            output(out);
            return out.toByteArray();
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    /** 關閉workbook並清除暫存檔 */
    @Override
    public void close() throws ExcelException {
        try {
            workbook.close();
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    /**
     * 完成所有寫出處理(僅執行一次)
     *
     * @throws ExcelException Excel處理錯誤
     */
    private void finish() throws ExcelException {
        if (finished) {
            return;
        }
        for (SheetContext context : sheetMap.values()) {
            flushSheet(context, true);
        }
        writePendingFormulas();
        for (SheetContext context : sheetMap.values()) {
            for (Integer columnNum : context.autoSizeColumnNumSet) {
                if (!context.overrideColumnWidthMap.containsKey(columnNum)) {
                    context.sheet.autoSizeColumn(columnNum);
                }
            }
        }
        workbook.setForceFormulaRecalculation(
                null != forceFormulaRecalculation ? forceFormulaRecalculation : hasFormula);
        finished = true;
    }

    /**
     * 寫出工作表資料
     *
     * @param context 工作表處理資料
     * @param isLast 是否為輸出前最後一次寫出
     * @throws ExcelException Excel處理錯誤
     */
    private void flushSheet(SheetContext context, boolean isLast) throws ExcelException {
        checkNotFinished();
        List<ExcelCell> cellList = new ArrayList<>(context.cellList);
        context.cellList.clear();
        cellList.sort(
                Comparator.comparingInt(ExcelCell::getRow).thenComparingInt(ExcelCell::getColumn));

        XSSFEvaluationWorkbook evaluationWorkbook = null;
        for (ExcelCell cell : cellList) {
            Row row = getOrCreateRow(context, cell.getRow());
            if (null == row) {
                throw new ExcelException(
                        String.format(
                                "Row %d of sheet '%s' has already been flushed and cannot be"
                                        + " written again.",
                                cell.getRow(), context.sheet.getSheetName()));
            }
            if (CellType.FORMULA.equals(cell.getCellType())
                    && workbook instanceof SXSSFWorkbook
                    && null == evaluationWorkbook) {
                evaluationWorkbook =
                        XSSFEvaluationWorkbook.create(((SXSSFWorkbook) workbook).getXSSFWorkbook());
            }
            writeCell(context, row, cell, evaluationWorkbook);
            context.maxRowNum = Math.max(context.maxRowNum, cell.getRow());
        }

        // 建立只有設定列高或隱藏的列
        int lastConfigRowNum = isLast ? Integer.MAX_VALUE : context.maxRowNum;
        Set<Integer> configRowNumSet = new TreeSet<>(context.rowHeightMap.keySet());
        configRowNumSet.addAll(context.hiddenRowNumSet);
        for (Integer rowNum : configRowNumSet) {
            if (rowNum <= lastConfigRowNum) {
                getOrCreateRow(context, rowNum);
            }
        }

        // 處理表格欄位合併(合併範圍的列都已寫入時才套用)
        Iterator<ExcelMergedRegion> iterator = context.mergedRegionList.iterator();
        while (iterator.hasNext()) {
            ExcelMergedRegion mergedRegion = iterator.next();
            if (isLast || mergedRegion.getLastRow() <= context.maxRowNum) {
                applyMergedRegion(context, mergedRegion);
                iterator.remove();
            }
        }

        if (context.sheet instanceof SXSSFSheet sxssfSheet) {
            // 保留尚未套用合併規則的列
            int keepFromRowNum = Integer.MAX_VALUE;
            for (ExcelMergedRegion mergedRegion : context.mergedRegionList) {
                keepFromRowNum = Math.min(keepFromRowNum, mergedRegion.getFirstRow());
            }
            int remaining = 0;
            Iterator<Row> rowIterator = sxssfSheet.rowIterator();
            while (rowIterator.hasNext()) {
                if (rowIterator.next().getRowNum() >= keepFromRowNum) {
                    remaining++;
                }
            }
            try {
                sxssfSheet.flushRows(remaining);
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }

    /**
     * 寫入欄位資料
     *
     * @param context 工作表處理資料
     * @param row 列
     * @param cell 欄位資料
     * @param evaluationWorkbook xlsx公式檢查用
     * @throws ExcelException Excel處理錯誤
     */
    private void writeCell(
            SheetContext context,
            Row row,
            ExcelCell cell,
            XSSFEvaluationWorkbook evaluationWorkbook)
            throws ExcelException {
        Cell poiCell = row.getCell(cell.getColumn());
        if (null == poiCell) {
            poiCell = row.createCell(cell.getColumn());
        }
        String value = null == cell.getValue() ? "" : cell.getValue();
        CellType cellType = null == cell.getCellType() ? CellType.STRING : cell.getCellType();
        switch (cellType) {
            case FORMULA:
                String formula = normalizeFormula(value);
                if (formula.isEmpty()) {
                    poiCell.setBlank();
                    break;
                }
                hasFormula = true;
                if (null == evaluationWorkbook) {
                    // xls延後至輸出時寫入公式
                    pendingFormulaList.add(new PendingFormula(context.sheet, poiCell, formula));
                } else {
                    try {
                        FormulaParser.parse(
                                formula,
                                evaluationWorkbook,
                                FormulaType.CELL,
                                workbook.getSheetIndex(context.sheet));
                    } catch (FormulaParseException e) {
                        throw new ExcelException(
                                cellErrorMessage(
                                        context.sheet, poiCell, "Invalid formula: " + formula),
                                e);
                    }
                    poiCell.setCellFormula(formula);
                }
                break;
            case NUMERIC:
                writeNumericCell(context, poiCell, value);
                break;
            case BOOLEAN:
                poiCell.setCellValue(Boolean.parseBoolean(value.trim()));
                break;
            case BLANK:
                poiCell.setBlank();
                break;
            case ERROR:
                try {
                    poiCell.setCellErrorValue(FormulaError.forString(value.trim()).getCode());
                } catch (IllegalArgumentException e) {
                    poiCell.setCellValue(value);
                }
                break;
            default:
                poiCell.setCellValue(value);
                break;
        }
        if (null != cell.getStyle()) {
            CellStyle cellStyle = cellStyleMap.get(cell.getStyle());
            if (null == cellStyle) {
                cellStyle = cell.getStyle().toCellStyle(workbook, colorPicker);
                cellStyleMap.put(cell.getStyle(), cellStyle);
            }
            poiCell.setCellStyle(cellStyle);
        }
    }

    /**
     * 寫入數值欄位，可接受數字或yyyy-MM-dd HH:mm:ss格式的日期時間
     *
     * @param context 工作表處理資料
     * @param poiCell 欄位
     * @param value 數值文字
     * @throws ExcelException 數值格式錯誤
     */
    private void writeNumericCell(SheetContext context, Cell poiCell, String value)
            throws ExcelException {
        String text = value.trim();
        if (text.isEmpty()) {
            poiCell.setBlank();
            return;
        }
        try {
            poiCell.setCellValue(Double.parseDouble(text));
            return;
        } catch (NumberFormatException ignored) {
            // 嘗試以日期時間格式解析
        }
        try {
            poiCell.setCellValue(LocalDateTime.parse(text, DATE_TIME_FORMATTER));
        } catch (DateTimeParseException e) {
            throw new ExcelException(
                    cellErrorMessage(context.sheet, poiCell, "Invalid numeric value: " + value), e);
        }
    }

    /**
     * 去除公式開頭的等號
     *
     * @param formula 公式
     * @return 公式
     */
    private String normalizeFormula(String formula) {
        String result = formula.trim();
        while (result.startsWith("=")) {
            result = result.substring(1).trim();
        }
        return result;
    }

    /**
     * 寫入xls延後的公式，並計算公式結果供不會重新計算的檢視器顯示
     *
     * @throws ExcelException 公式錯誤
     */
    private void writePendingFormulas() throws ExcelException {
        if (pendingFormulaList.isEmpty()) {
            return;
        }
        for (PendingFormula pendingFormula : pendingFormulaList) {
            Cell cell = pendingFormula.cell;
            try {
                cell.setCellFormula(pendingFormula.formula);
            } catch (FormulaParseException e) {
                throw new ExcelException(
                        cellErrorMessage(
                                pendingFormula.sheet,
                                cell,
                                "Invalid formula: " + pendingFormula.formula),
                        e);
            }
            // 參照不存在的工作表時POI會寫成#REF!而不報錯
            if (!pendingFormula.formula.contains("#REF!")
                    && cell.getCellFormula().contains("#REF!")) {
                throw new ExcelException(
                        cellErrorMessage(
                                pendingFormula.sheet,
                                cell,
                                "Formula refers to a sheet that does not exist: "
                                        + pendingFormula.formula));
            }
        }
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        for (PendingFormula pendingFormula : pendingFormulaList) {
            try {
                evaluator.evaluateFormulaCell(pendingFormula.cell);
            } catch (RuntimeException ignored) {
                // 不支援的函式交由Excel開檔時計算
            }
        }
        pendingFormulaList.clear();
    }

    /**
     * 套用表格欄位合併規則
     *
     * @param context 工作表處理資料
     * @param mergedRegion 合併規則
     */
    private void applyMergedRegion(SheetContext context, ExcelMergedRegion mergedRegion) {
        context.sheet.addMergedRegionUnsafe(
                new CellRangeAddress(
                        mergedRegion.getFirstRow(),
                        mergedRegion.getLastRow(),
                        mergedRegion.getFirstColumn(),
                        mergedRegion.getLastColumn()));
        if (!mergedRegion.hasBorder()) {
            return;
        }
        CellStyle cellStyle = null;
        for (int rowNum = mergedRegion.getFirstRow();
                rowNum <= mergedRegion.getLastRow();
                rowNum++) {
            Row row = getOrCreateRow(context, rowNum);
            if (null == row) {
                // 已寫出的列無法再調整樣式
                continue;
            }
            for (int columnNum = mergedRegion.getFirstColumn();
                    columnNum <= mergedRegion.getLastColumn();
                    columnNum++) {
                Cell cell = row.getCell(columnNum);
                if (null == cell) {
                    cell = row.createCell(columnNum);
                }
                if (null == cellStyle) {
                    cellStyle = getMergedCellStyle(cell.getCellStyle(), mergedRegion);
                }
                cell.setCellStyle(cellStyle);
            }
        }
    }

    /**
     * 取得套用合併邊線後的樣式(相同樣式只建立一次)
     *
     * @param baseCellStyle 合併範圍左上角欄位樣式
     * @param mergedRegion 合併規則
     * @return 欄位樣式
     */
    private CellStyle getMergedCellStyle(CellStyle baseCellStyle, ExcelMergedRegion mergedRegion) {
        String key =
                String.join(
                        "|",
                        String.valueOf(baseCellStyle.getIndex()),
                        String.valueOf(mergedRegion.getBorderTop()),
                        String.valueOf(mergedRegion.getBorderBottom()),
                        String.valueOf(mergedRegion.getBorderLeft()),
                        String.valueOf(mergedRegion.getBorderRight()),
                        String.valueOf(mergedRegion.getBorderTopColor()),
                        String.valueOf(mergedRegion.getBorderBottomColor()),
                        String.valueOf(mergedRegion.getBorderLeftColor()),
                        String.valueOf(mergedRegion.getBorderRightColor()));
        CellStyle cellStyle = mergedCellStyleMap.get(key);
        if (null != cellStyle) {
            return cellStyle;
        }
        cellStyle = workbook.createCellStyle();
        cellStyle.cloneStyleFrom(baseCellStyle);
        if (!BorderStyle.NONE.equals(mergedRegion.getBorderTop())) {
            cellStyle.setBorderTop(mergedRegion.getBorderTop());
            ExcelColorUtils.setBorderColor(
                    cellStyle, BorderSide.TOP, mergedRegion.getBorderTopColor(), colorPicker);
        }
        if (!BorderStyle.NONE.equals(mergedRegion.getBorderBottom())) {
            cellStyle.setBorderBottom(mergedRegion.getBorderBottom());
            ExcelColorUtils.setBorderColor(
                    cellStyle, BorderSide.BOTTOM, mergedRegion.getBorderBottomColor(), colorPicker);
        }
        if (!BorderStyle.NONE.equals(mergedRegion.getBorderLeft())) {
            cellStyle.setBorderLeft(mergedRegion.getBorderLeft());
            ExcelColorUtils.setBorderColor(
                    cellStyle, BorderSide.LEFT, mergedRegion.getBorderLeftColor(), colorPicker);
        }
        if (!BorderStyle.NONE.equals(mergedRegion.getBorderRight())) {
            cellStyle.setBorderRight(mergedRegion.getBorderRight());
            ExcelColorUtils.setBorderColor(
                    cellStyle, BorderSide.RIGHT, mergedRegion.getBorderRightColor(), colorPicker);
        }
        mergedCellStyleMap.put(key, cellStyle);
        return cellStyle;
    }

    /**
     * 取得或建立列，建立時套用列高及隱藏設定
     *
     * @param context 工作表處理資料
     * @param rowNum 列
     * @return 列，已寫出(xlsx)時回傳null
     */
    private Row getOrCreateRow(SheetContext context, int rowNum) {
        Row row = context.sheet.getRow(rowNum);
        if (null != row) {
            return row;
        }
        if (isRowFlushed(context, rowNum)) {
            return null;
        }
        row = context.sheet.createRow(rowNum);
        Float heightInPoints = context.rowHeightMap.get(rowNum);
        if (null != heightInPoints) {
            row.setHeightInPoints(heightInPoints);
        }
        if (context.hiddenRowNumSet.contains(rowNum)) {
            row.setZeroHeight(true);
        }
        return row;
    }

    private boolean isRowFlushed(SheetContext context, int rowNum) {
        return context.sheet instanceof SXSSFSheet
                && rowNum <= ((SXSSFSheet) context.sheet).getLastFlushedRowNum();
    }

    private void checkRowNotFlushed(SheetContext context, int rowNum) {
        if (isRowFlushed(context, rowNum)) {
            throw new IllegalStateException(
                    String.format(
                            "Row %d of sheet '%s' has already been flushed and cannot be changed.",
                            rowNum, context.sheet.getSheetName()));
        }
    }

    private void checkNotFinished() {
        if (finished) {
            throw new IllegalStateException("Workbook has been output and cannot be changed.");
        }
    }

    private SheetContext getSheetContext(String sheetCode) {
        checkNotFinished();
        return sheetMap.get(sheetCode);
    }

    private String cellErrorMessage(Sheet sheet, Cell cell, String message) {
        return String.format(
                "%s (sheet: '%s', row: %d, column: %d)",
                message, sheet.getSheetName(), cell.getRowIndex(), cell.getColumnIndex());
    }

    /** 工作表處理資料 */
    private static class SheetContext {

        private final Sheet sheet;

        /** 尚未寫出的欄位資料 */
        private final List<ExcelCell> cellList = new ArrayList<>();

        /** 尚未套用的合併規則 */
        private final List<ExcelMergedRegion> mergedRegionList = new ArrayList<>();

        /** 列高 */
        private final Map<Integer, Float> rowHeightMap = new HashMap<>();

        /** 隱藏列 */
        private final Set<Integer> hiddenRowNumSet = new HashSet<>();

        /** 自動調整寬度的欄位 */
        private final Set<Integer> autoSizeColumnNumSet = new LinkedHashSet<>();

        /** 覆寫欄位寬度 */
        private final Map<Integer, Integer> overrideColumnWidthMap = new HashMap<>();

        /** 已寫入的最大列 */
        private int maxRowNum = -1;

        private SheetContext(Sheet sheet) {
            this.sheet = sheet;
        }
    }

    /** xls延後寫入的公式 */
    private record PendingFormula(Sheet sheet, Cell cell, String formula) {}
}
