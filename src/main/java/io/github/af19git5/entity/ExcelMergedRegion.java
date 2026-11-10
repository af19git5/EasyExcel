package io.github.af19git5.entity;

import io.github.af19git5.builder.ExcelMergedRegionBuilder;
import io.github.af19git5.utils.ExcelColorUtils;

import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

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
        String hex = ExcelColorUtils.normalizeHex(colorHex);
        if (null != hex) {
            this.borderTopColor = hex;
        }
    }

    public void setBorderTopColor(@NonNull IndexedColors indexedColor) {
        this.borderTopColor = ExcelColorUtils.toHex(indexedColor);
    }

    public void setBorderBottomColor(@NonNull String colorHex) {
        String hex = ExcelColorUtils.normalizeHex(colorHex);
        if (null != hex) {
            this.borderBottomColor = hex;
        }
    }

    public void setBorderBottomColor(@NonNull IndexedColors indexedColor) {
        this.borderBottomColor = ExcelColorUtils.toHex(indexedColor);
    }

    public void setBorderLeftColor(@NonNull String colorHex) {
        String hex = ExcelColorUtils.normalizeHex(colorHex);
        if (null != hex) {
            this.borderLeftColor = hex;
        }
    }

    public void setBorderLeftColor(@NonNull IndexedColors indexedColor) {
        this.borderLeftColor = ExcelColorUtils.toHex(indexedColor);
    }

    public void setBorderRightColor(@NonNull String colorHex) {
        String hex = ExcelColorUtils.normalizeHex(colorHex);
        if (null != hex) {
            this.borderRightColor = hex;
        }
    }

    public void setBorderRightColor(@NonNull IndexedColors indexedColor) {
        this.borderRightColor = ExcelColorUtils.toHex(indexedColor);
    }

    public void setAllBorderColor(@NonNull String colorHex) {
        String hex = ExcelColorUtils.normalizeHex(colorHex);
        if (null != hex) {
            this.borderTopColor = hex;
            this.borderBottomColor = hex;
            this.borderLeftColor = hex;
            this.borderRightColor = hex;
        }
    }

    public void setAllBorderColor(@NonNull IndexedColors indexedColor) {
        String hex = ExcelColorUtils.toHex(indexedColor);
        this.borderTopColor = hex;
        this.borderBottomColor = hex;
        this.borderLeftColor = hex;
        this.borderRightColor = hex;
    }

    /**
     * 是否有設定任一邊線
     *
     * @return 是否有邊線
     */
    public boolean hasBorder() {
        return !BorderStyle.NONE.equals(borderTop)
                || !BorderStyle.NONE.equals(borderBottom)
                || !BorderStyle.NONE.equals(borderLeft)
                || !BorderStyle.NONE.equals(borderRight);
    }
}
