package io.github.af19git5.entity;

import io.github.af19git5.builder.ExcelCellBuilder;

import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;

import org.apache.poi.ss.usermodel.CellType;

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
        this.value = value;
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

    public static ExcelCellBuilder init(
            String value, @NonNull Integer row, @NonNull Integer column) {
        return new ExcelCellBuilder(row, column, value);
    }

    public static ExcelCellBuilder init(
            @NonNull Integer row, @NonNull Integer column, String value) {
        return new ExcelCellBuilder(row, column, value);
    }
}
