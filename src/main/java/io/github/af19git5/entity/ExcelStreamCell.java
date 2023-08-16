package io.github.af19git5.entity;

import lombok.Data;
import lombok.NonNull;

import org.apache.poi.ss.usermodel.CellType;

/**
 * Excel欄位資料
 *
 * @author Jimmy Kang
 */
@Data
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

    public ExcelStreamCell(@NonNull String value, @NonNull Integer row, @NonNull Integer column) {
        this.value = value;
        this.row = row;
        this.column = column;
    }

    public ExcelStreamCell(
            @NonNull String value,
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull CellType cellType) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
    }

    public ExcelStreamCell(
            @NonNull String value,
            @NonNull Integer row,
            @NonNull Integer column,
            ExcelStreamStyle style) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.style = style;
    }

    public ExcelStreamCell(
            @NonNull String value,
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull CellType cellType,
            ExcelStreamStyle style) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
        this.style = style;
    }

    public ExcelStreamCell(@NonNull Integer row, @NonNull Integer column, @NonNull String value) {
        this.value = value;
        this.row = row;
        this.column = column;
    }

    public ExcelStreamCell(
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull String value,
            @NonNull CellType cellType) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
    }

    public ExcelStreamCell(
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull String value,
            ExcelStreamStyle style) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.style = style;
    }

    public ExcelStreamCell(
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull String value,
            @NonNull CellType cellType,
            ExcelStreamStyle style) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
        this.style = style;
    }
}
