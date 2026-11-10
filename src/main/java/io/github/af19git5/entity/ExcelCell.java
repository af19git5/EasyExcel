package io.github.af19git5.entity;

import io.github.af19git5.builder.ExcelCellBuilder;

import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;

import org.apache.poi.ss.usermodel.CellType;

import java.math.BigDecimal;

/**
 * Excel欄位資料
 *
 * @author Jimmy Kang
 */
@Getter
@Setter
public class ExcelCell {

    /** 欄位數值 */
    private String value;

    /** 欄位類別 */
    private CellType cellType = CellType.STRING;

    /** 橫列(從0開始) */
    private Integer row;

    /** 直行(從0開始) */
    private Integer column;

    /** 欄位樣式 */
    private ExcelStyle style;

    public ExcelCell(String value, @NonNull Integer row, @NonNull Integer column) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
    }

    public ExcelCell(
            String value,
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull CellType cellType) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
    }

    public ExcelCell(
            String value, @NonNull Integer row, @NonNull Integer column, ExcelStyle style) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.style = style;
    }

    public ExcelCell(
            String value,
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull CellType cellType,
            ExcelStyle style) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
        this.style = style;
    }

    public ExcelCell(@NonNull Integer row, @NonNull Integer column, String value) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
    }

    public ExcelCell(
            @NonNull Integer row,
            @NonNull Integer column,
            String value,
            @NonNull CellType cellType) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
    }

    public ExcelCell(
            @NonNull Integer row, @NonNull Integer column, String value, ExcelStyle style) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.style = style;
    }

    public ExcelCell(
            @NonNull Integer row,
            @NonNull Integer column,
            String value,
            @NonNull CellType cellType,
            ExcelStyle style) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
        this.style = style;
    }

    /**
     * 建立數字欄位(自動設為NUMERIC)
     *
     * @param row 橫列(從0開始)
     * @param column 直行(從0開始)
     * @param value 數值，為null時寫出空白欄位
     * @return 欄位資料
     */
    public static ExcelCell number(@NonNull Integer row, @NonNull Integer column, Number value) {
        return new ExcelCell(row, column, toNumberText(value), CellType.NUMERIC);
    }

    /**
     * 建立數字欄位(自動設為NUMERIC)
     *
     * @param row 橫列(從0開始)
     * @param column 直行(從0開始)
     * @param value 數值，為null時寫出空白欄位
     * @param style 欄位樣式
     * @return 欄位資料
     */
    public static ExcelCell number(
            @NonNull Integer row, @NonNull Integer column, Number value, ExcelStyle style) {
        return new ExcelCell(row, column, toNumberText(value), CellType.NUMERIC, style);
    }

    /**
     * 數值轉為文字(BigDecimal不使用科學記號)
     *
     * @param value 數值
     * @return 數值文字
     */
    public static String toNumberText(Number value) {
        if (null == value) {
            return "";
        }
        if (value instanceof BigDecimal) {
            return ((BigDecimal) value).toPlainString();
        }
        return value.toString();
    }

    public static ExcelCellBuilder init(
            String value, @NonNull Integer row, @NonNull Integer column) {
        return new ExcelCellBuilder(row, column, value);
    }

    public static ExcelCellBuilder init(
            @NonNull Integer row, @NonNull Integer column, String value) {
        return new ExcelCellBuilder(row, column, value);
    }
}
