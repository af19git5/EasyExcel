package io.github.af19git5.entity;

import io.github.af19git5.builder.ExcelStreamCellBuilder;

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
public class ExcelStreamCell {

    /** 欄位數值 */
    private String value;

    /** 欄位類別 */
    public CellType cellType = CellType.STRING;

    /** 橫列(從0開始) */
    private Integer row;

    /** 直行(從0開始) */
    private Integer column;

    /** 欄位樣式 */
    private ExcelStreamStyle style;

    public ExcelStreamCell(String value, @NonNull Integer row, @NonNull Integer column) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
    }

    public ExcelStreamCell(
            String value,
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull CellType cellType) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
    }

    public ExcelStreamCell(
            String value, @NonNull Integer row, @NonNull Integer column, ExcelStreamStyle style) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.style = style;
    }

    public ExcelStreamCell(
            String value,
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull CellType cellType,
            ExcelStreamStyle style) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
        this.style = style;
    }

    public ExcelStreamCell(@NonNull Integer row, @NonNull Integer column, String value) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
    }

    public ExcelStreamCell(
            @NonNull Integer row,
            @NonNull Integer column,
            String value,
            @NonNull CellType cellType) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
    }

    public ExcelStreamCell(
            @NonNull Integer row, @NonNull Integer column, String value, ExcelStreamStyle style) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.style = style;
    }

    public ExcelStreamCell(
            @NonNull Integer row,
            @NonNull Integer column,
            String value,
            @NonNull CellType cellType,
            ExcelStreamStyle style) {
        this.value = null == value ? "" : value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
        this.style = style;
    }

    public static ExcelStreamCellBuilder init(
            String value, @NonNull Integer row, @NonNull Integer column) {
        return new ExcelStreamCellBuilder(row, column, value);
    }

    public static ExcelStreamCellBuilder init(
            @NonNull Integer row, @NonNull Integer column, String value) {
        return new ExcelStreamCellBuilder(row, column, value);
    }
}
