package io.github.af19git5.entity;

import io.github.af19git5.builder.ExcelStreamMergedRegionBuilder;

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
public class ExcelStreamMergedRegion {

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
    private IndexedColors borderTopColor;

    /** 下邊線顏色 */
    private IndexedColors borderBottomColor;

    /** 左邊線顏色 */
    private IndexedColors borderLeftColor;

    /** 右邊線顏色 */
    private IndexedColors borderRightColor;

    public ExcelStreamMergedRegion(int firstRow, int lastRow, int firstColumn, int lastColumn) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstColumn = firstColumn;
        this.lastColumn = lastColumn;
    }

    public ExcelStreamMergedRegion(
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

    public static ExcelStreamMergedRegionBuilder init(
            int firstRow, int lastRow, int firstColumn, int lastColumn) {
        return new ExcelStreamMergedRegionBuilder(firstRow, lastRow, firstColumn, lastColumn);
    }

    public void setAllBorder(BorderStyle borderStyle) {
        this.borderTop = borderStyle;
        this.borderBottom = borderStyle;
        this.borderLeft = borderStyle;
        this.borderRight = borderStyle;
    }

    public void setBorderTopColor(@NonNull IndexedColors indexedColor) {
        this.borderTopColor = indexedColor;
    }

    public void setBorderBottomColor(@NonNull IndexedColors indexedColor) {
        this.borderBottomColor = indexedColor;
    }

    public void setBorderLeftColor(@NonNull IndexedColors indexedColor) {
        this.borderLeftColor = indexedColor;
    }

    public void setBorderRightColor(@NonNull IndexedColors indexedColor) {
        this.borderRightColor = indexedColor;
    }

    public void setAllBorderColor(@NonNull IndexedColors indexedColor) {
        this.borderTopColor = indexedColor;
        this.borderBottomColor = indexedColor;
        this.borderLeftColor = indexedColor;
        this.borderRightColor = indexedColor;
    }
}
