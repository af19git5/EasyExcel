package io.github.af19git5.entity;

import io.github.af19git5.builder.ExcelMergedRegionBuilder;

import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * Excel合併欄位規則
 *
 * @author Jimmy Kang
 */
@Getter
@Setter
public class ExcelMergedRegion {

    /** 起始列 */
    private Integer firstRow;

    /** 結束列 */
    private Integer lastRow;

    /** 起始行 */
    private Integer firstColumn;

    /** 結束行 */
    private Integer lastColumn;

    /** 上邊線樣式 */
    private BorderStyle borderTop = BorderStyle.NONE;

    /** 下邊線樣式 */
    private BorderStyle borderBottom = BorderStyle.NONE;

    /** 左邊線樣式 */
    private BorderStyle borderLeft = BorderStyle.NONE;

    /** 右邊線樣式 */
    private BorderStyle borderRight = BorderStyle.NONE;

    /** 上邊線顏色 */
    private String borderTopColor;

    /** 下邊線顏色 */
    private String borderBottomColor;

    /** 左邊線顏色 */
    private String borderLeftColor;

    /** 右邊線顏色 */
    private String borderRightColor;

    public ExcelMergedRegion(int firstRow, int lastRow, int firstColumn, int lastColumn) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstColumn = firstColumn;
        this.lastColumn = lastColumn;
    }

    public ExcelMergedRegion(
            int firstRow, int lastRow, int firstColumn, int lastColumn, BorderStyle borderStyle) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstColumn = firstColumn;
        this.lastColumn = lastColumn;
        this.borderTop = borderStyle;
        this.borderBottom = borderStyle;
        this.borderLeft = borderStyle;
        this.borderRight = borderStyle;
    }

    public static ExcelMergedRegionBuilder init(
            int firstRow, int lastRow, int firstColumn, int lastColumn) {
        return new ExcelMergedRegionBuilder(firstRow, lastRow, firstColumn, lastColumn);
    }

    public void setAllBorder(BorderStyle borderStyle) {
        this.borderTop = borderStyle;
        this.borderBottom = borderStyle;
        this.borderLeft = borderStyle;
        this.borderRight = borderStyle;
    }

    public void setBorderTopColor(@NonNull String colorHex) {
        if (isValidColorHex(colorHex)) {
            this.borderTopColor = colorHex;
        }
    }

    public void setBorderBottomColor(@NonNull String colorHex) {
        if (isValidColorHex(colorHex)) {
            this.borderBottomColor = colorHex;
        }
    }

    public void setBorderLeftColor(@NonNull String colorHex) {
        if (isValidColorHex(colorHex)) {
            this.borderLeftColor = colorHex;
        }
    }

    public void setBorderRightColor(@NonNull String colorHex) {
        if (isValidColorHex(colorHex)) {
            this.borderRightColor = colorHex;
        }
    }

    public void setAllBorderColor(@NonNull String colorHex) {
        if (isValidColorHex(colorHex)) {
            this.borderTopColor = colorHex;
            this.borderBottomColor = colorHex;
            this.borderLeftColor = colorHex;
            this.borderRightColor = colorHex;
        }
    }

    /**
     * 檢查是否符合16進位色碼
     *
     * @param colorHex 16進位色碼
     * @return 是否符合格式
     */
    private boolean isValidColorHex(@NonNull String colorHex) {
        return colorHex.matches("^#(?:[0-9a-fA-F]{3}){1,2}$");
    }
}
