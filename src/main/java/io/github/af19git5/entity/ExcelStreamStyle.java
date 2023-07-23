package io.github.af19git5.entity;

import lombok.Data;

import org.apache.poi.ss.usermodel.*;

/**
 * Excel欄位樣式資料
 *
 * @author Jimmy Kang
 */
@Data
public class ExcelStreamStyle {

    /** 上邊線樣式 */
    private BorderStyle borderTop = BorderStyle.NONE;

    /** 下邊線樣式 */
    private BorderStyle borderBottom = BorderStyle.NONE;

    /** 左邊線樣式 */
    private BorderStyle borderLeft = BorderStyle.NONE;

    /** 右邊線樣式 */
    private BorderStyle borderRight = BorderStyle.NONE;

    /** 上邊線顏色 */
    public IndexedColors borderTopColor;

    /** 下邊線顏色 */
    public IndexedColors borderBottomColor;

    /** 左邊線顏色 */
    public IndexedColors borderLeftColor;

    /** 右邊線顏色 */
    public IndexedColors borderRightColor;

    /** 水平對齊位置 */
    public HorizontalAlignment horizontalAlignment = HorizontalAlignment.LEFT;

    /** 垂直對齊位置 */
    public VerticalAlignment verticalAlignment = VerticalAlignment.TOP;

    /** 背景顏色 */
    public IndexedColors backgroundColor;

    /** 文字字體 */
    public String fontName;

    /** 文字大小 */
    public Integer fontSize = 10;

    /** 是否為粗體 */
    public Boolean bold = false;

    /** 是否為斜體 */
    public Boolean italic = false;

    /** 是否加入刪除線 */
    public Boolean strikeout = false;

    /** 文字顏色 */
    public IndexedColors fontColor;

    public void setAllBorder(BorderStyle borderStyle) {
        this.borderTop = borderStyle;
        this.borderBottom = borderStyle;
        this.borderLeft = borderStyle;
        this.borderRight = borderStyle;
    }

    public void setAllBorderColor(IndexedColors indexedColor) {
        this.borderTopColor = indexedColor;
        this.borderBottomColor = indexedColor;
        this.borderLeftColor = indexedColor;
        this.borderRightColor = indexedColor;
    }

    public ExcelStreamStyle() {}

    public CellStyle toCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(this.borderTop);
        cellStyle.setBorderBottom(this.borderBottom);
        cellStyle.setBorderLeft(this.borderLeft);
        cellStyle.setBorderRight(this.borderRight);

        if (null != this.borderTopColor) {
            cellStyle.setTopBorderColor(borderTopColor.getIndex());
        }

        if (null != this.borderBottomColor) {
            cellStyle.setBottomBorderColor(borderBottomColor.getIndex());
        }

        if (null != this.borderLeftColor) {
            cellStyle.setLeftBorderColor(borderLeftColor.getIndex());
        }

        if (null != this.borderRightColor) {
            cellStyle.setRightBorderColor(borderRightColor.getIndex());
        }

        cellStyle.setAlignment(this.horizontalAlignment);
        cellStyle.setVerticalAlignment(this.verticalAlignment);

        if (null != this.backgroundColor) {
            cellStyle.setFillForegroundColor(backgroundColor.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        Font font = workbook.createFont();
        if (null != this.fontName && !this.fontName.isEmpty()) {
            font.setFontName(this.fontName);
        }
        font.setFontHeightInPoints(this.fontSize.shortValue());
        font.setBold(this.bold);
        font.setItalic(this.italic);
        font.setStrikeout(this.strikeout);

        if (null != this.fontColor) {
            font.setColor(fontColor.getIndex());
        }

        cellStyle.setFont(font);
        return cellStyle;
    }
}
