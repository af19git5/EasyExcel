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
public class ExcelCell {

    /** 欄位數值 */
    private String value;

    /** 欄位類別 */
    public CellType cellType = CellType.STRING;

    /** 橫列(從0開始) */
    private Integer row;

    /** 直行(從0開始) */
    private Integer column;

    /** 欄位樣式 */
    private ExcelStyle style;

    public ExcelCell(@NonNull String value, @NonNull Integer row, @NonNull Integer column) {
        this.value = value;
        this.row = row;
        this.column = column;
    }

    public ExcelCell(
            @NonNull String value,
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull CellType cellType) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
    }

    public ExcelCell(
            @NonNull String value,
            @NonNull Integer row,
            @NonNull Integer column,
            ExcelStyle style) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.style = style;
    }

    public ExcelCell(
            @NonNull String value,
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull CellType cellType,
            ExcelStyle style) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
        this.style = style;
    }

    public ExcelCell(@NonNull Integer row, @NonNull Integer column, @NonNull String value) {
        this.value = value;
        this.row = row;
        this.column = column;
    }

    public ExcelCell(
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull String value,
            @NonNull CellType cellType) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
    }

    public ExcelCell(
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull String value,
            ExcelStyle style) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.style = style;
    }

    public ExcelCell(
            @NonNull Integer row,
            @NonNull Integer column,
            @NonNull String value,
            @NonNull CellType cellType,
            ExcelStyle style) {
        this.value = value;
        this.row = row;
        this.column = column;
        this.cellType = cellType;
        this.style = style;
    }
}
