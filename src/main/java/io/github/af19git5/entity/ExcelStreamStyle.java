package io.github.af19git5.entity;

import io.github.af19git5.builder.StreamStyleBuilder;

import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;

import org.apache.poi.ss.usermodel.*;

/**
 * Excel欄位樣式資料
 *
 * @author Jimmy Kang
 */
@Getter
@Setter
public class ExcelStreamStyle {

    /** 是否自動換行 */
    @NonNull private Boolean isWrapText = false;

    @NonNull private Boolean isLock = false;

    /** 上邊線樣式 */
    @NonNull private BorderStyle borderTop = BorderStyle.NONE;

    /** 下邊線樣式 */
    @NonNull private BorderStyle borderBottom = BorderStyle.NONE;

    /** 左邊線樣式 */
    @NonNull private BorderStyle borderLeft = BorderStyle.NONE;

    /** 右邊線樣式 */
    @NonNull private BorderStyle borderRight = BorderStyle.NONE;

    /** 上邊線顏色 */
    private IndexedColors borderTopColor;

    /** 下邊線顏色 */
    private IndexedColors borderBottomColor;

    /** 左邊線顏色 */
    private IndexedColors borderLeftColor;

    /** 右邊線顏色 */
    private IndexedColors borderRightColor;

    /** 水平對齊位置 */
    @NonNull private HorizontalAlignment horizontalAlignment = HorizontalAlignment.LEFT;

    /** 垂直對齊位置 */
    @NonNull private VerticalAlignment verticalAlignment = VerticalAlignment.TOP;

    /** 背景顏色 */
    private IndexedColors backgroundColor;

    /** 文字字體 */
    private String fontName;

    /** 文字大小 */
    @NonNull private Integer fontSize = 10;

    /** 文字顏色 */
    private IndexedColors fontColor;

    /** 是否為粗體 */
    @NonNull private Boolean bold = false;

    /** 是否為斜體 */
    @NonNull private Boolean italic = false;

    /** 是否加入刪除線 */
    @NonNull private Boolean strikeout = false;

    public static StreamStyleBuilder init() {
        return new StreamStyleBuilder();
    }

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

    public CellStyle toCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setWrapText(this.isWrapText);
        cellStyle.setLocked(this.isLock);
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
