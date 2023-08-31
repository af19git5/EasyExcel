package io.github.af19git5.entity;

import lombok.Data;

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * Excel合併欄位規則
 *
 * @author Jimmy Kang
 */
@Data
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

    public void setAllBorder(BorderStyle borderStyle) {
        this.borderTop = borderStyle;
        this.borderBottom = borderStyle;
        this.borderLeft = borderStyle;
        this.borderRight = borderStyle;
    }
}
