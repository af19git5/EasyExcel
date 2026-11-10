package io.github.af19git5.entity;

import io.github.af19git5.builder.ExcelStyleBuilder;
import io.github.af19git5.utils.ExcelColorUtils;
import io.github.af19git5.utils.HSSFColorPicker;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

/**
 * Excel欄位樣式資料
 *
 * <p>以欄位值判斷是否相等，相同樣式在同一份活頁簿只會建立一次
 *
 * @author Jimmy Kang
 */
@Getter
@Setter
@EqualsAndHashCode
public class ExcelStyle {

    /** 是否自動換行 */
    @NonNull private Boolean isWrapText = false;

    /** 是否鎖定 */
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
    private String borderTopColor;

    /** 下邊線顏色 */
    private String borderBottomColor;

    /** 左邊線顏色 */
    private String borderLeftColor;

    /** 右邊線顏色 */
    private String borderRightColor;

    /** 水平對齊位置 */
    @NonNull private HorizontalAlignment horizontalAlignment = HorizontalAlignment.LEFT;

    /** 垂直對齊位置 */
    @NonNull private VerticalAlignment verticalAlignment = VerticalAlignment.TOP;

    /** 背景顏色 */
    private String backgroundColor;

    /** 文字字體 */
    private String fontName;

    /** 文字大小 */
    @NonNull private Integer fontSize = 10;

    /** 文字顏色 */
    private String fontColor;

    /** 是否為粗體 */
    @NonNull private Boolean bold = false;

    /** 是否為斜體 */
    @NonNull private Boolean italic = false;

    /** 是否加入刪除線 */
    @NonNull private Boolean strikeout = false;

    /** 數值格式(例如: #,##0、0.00%) */
    private String dataFormat;

    public static ExcelStyleBuilder init() {
        return new ExcelStyleBuilder();
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

    public void setBackgroundColor(@NonNull String colorHex) {
        String hex = ExcelColorUtils.normalizeHex(colorHex);
        if (null != hex) {
            this.backgroundColor = hex;
        }
    }

    public void setBackgroundColor(@NonNull IndexedColors indexedColor) {
        this.backgroundColor = ExcelColorUtils.toHex(indexedColor);
    }

    public void setFontColor(@NonNull String colorHex) {
        String hex = ExcelColorUtils.normalizeHex(colorHex);
        if (null != hex) {
            this.fontColor = hex;
        }
    }

    public void setFontColor(@NonNull IndexedColors indexedColor) {
        this.fontColor = ExcelColorUtils.toHex(indexedColor);
    }

    public void setAllBorder(@NonNull BorderStyle borderStyle) {
        this.borderTop = borderStyle;
        this.borderBottom = borderStyle;
        this.borderLeft = borderStyle;
        this.borderRight = borderStyle;
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

    public ExcelStyle() {}

    public ExcelStyle(@NonNull HSSFWorkbook workbook, @NonNull HSSFCellStyle cellStyle) {
        HSSFPalette palette = workbook.getCustomPalette();
        this.isWrapText = cellStyle.getWrapText();
        this.isLock = cellStyle.getLocked();
        this.borderTop = cellStyle.getBorderTop();
        this.borderBottom = cellStyle.getBorderBottom();
        this.borderLeft = cellStyle.getBorderLeft();
        this.borderRight = cellStyle.getBorderRight();
        this.borderTopColor =
                convertHSSFColorToHex(palette.getColor(cellStyle.getTopBorderColor()));
        this.borderBottomColor =
                convertHSSFColorToHex(palette.getColor(cellStyle.getBottomBorderColor()));
        this.borderLeftColor =
                convertHSSFColorToHex(palette.getColor(cellStyle.getLeftBorderColor()));
        this.borderRightColor =
                convertHSSFColorToHex(palette.getColor(cellStyle.getRightBorderColor()));
        this.horizontalAlignment = cellStyle.getAlignment();
        this.verticalAlignment = cellStyle.getVerticalAlignment();
        if (FillPatternType.SOLID_FOREGROUND.equals(cellStyle.getFillPattern())) {
            this.backgroundColor = convertHSSFColorToHex(cellStyle.getFillForegroundColorColor());
        }
        HSSFFont font = cellStyle.getFont(workbook);
        this.fontName = font.getFontName();
        this.fontSize = (int) font.getFontHeightInPoints();
        this.bold = font.getBold();
        this.italic = font.getItalic();
        this.strikeout = font.getStrikeout();
        this.fontColor = convertHSSFColorToHex(font.getHSSFColor(workbook));
        this.dataFormat = convertDataFormat(cellStyle.getDataFormatString());
    }

    private String convertHSSFColorToHex(HSSFColor color) {
        if (null == color) return null;
        short[] rgbArray = color.getTriplet();
        if (null == rgbArray || rgbArray.length != 3) return null;
        return String.format("#%02X%02X%02X", rgbArray[0], rgbArray[1], rgbArray[2]);
    }

    public ExcelStyle(@NonNull XSSFCellStyle cellStyle) {
        this.isWrapText = cellStyle.getWrapText();
        this.isLock = cellStyle.getLocked();
        this.borderTop = cellStyle.getBorderTop();
        this.borderBottom = cellStyle.getBorderBottom();
        this.borderLeft = cellStyle.getBorderLeft();
        this.borderRight = cellStyle.getBorderRight();
        this.borderTopColor = convertXSSColorToHax(cellStyle.getTopBorderXSSFColor());
        this.borderBottomColor = convertXSSColorToHax(cellStyle.getBottomBorderXSSFColor());
        this.borderLeftColor = convertXSSColorToHax(cellStyle.getLeftBorderXSSFColor());
        this.borderRightColor = convertXSSColorToHax(cellStyle.getRightBorderXSSFColor());
        this.horizontalAlignment = cellStyle.getAlignment();
        this.verticalAlignment = cellStyle.getVerticalAlignment();
        if (FillPatternType.SOLID_FOREGROUND.equals(cellStyle.getFillPattern())) {
            this.backgroundColor = convertXSSColorToHax(cellStyle.getFillForegroundColorColor());
        }
        XSSFFont font = cellStyle.getFont();
        this.fontName = font.getFontName();
        this.fontSize = (int) font.getFontHeightInPoints();
        this.bold = font.getBold();
        this.italic = font.getItalic();
        this.strikeout = font.getStrikeout();
        this.fontColor = convertXSSColorToHax(font.getXSSFColor());
        this.dataFormat = convertDataFormat(cellStyle.getDataFormatString());
    }

    private String convertXSSColorToHax(XSSFColor color) {
        if (null == color) return null;
        if (null == color.getARGBHex()) return null;
        return "#" + color.getARGBHex().substring(2);
    }

    private String convertDataFormat(String dataFormat) {
        if (null == dataFormat || "General".equals(dataFormat)) return null;
        return dataFormat;
    }

    /**
     * 轉換為POI欄位樣式
     *
     * @param workbook 活頁簿(xls或xlsx)
     * @param colorPicker xls調色盤，xlsx時可為null
     * @return POI欄位樣式
     */
    public CellStyle toCellStyle(@NonNull Workbook workbook, HSSFColorPicker colorPicker) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setLocked(this.isLock);
        cellStyle.setWrapText(this.isWrapText);
        cellStyle.setBorderTop(this.borderTop);
        cellStyle.setBorderBottom(this.borderBottom);
        cellStyle.setBorderLeft(this.borderLeft);
        cellStyle.setBorderRight(this.borderRight);
        ExcelColorUtils.setBorderColor(cellStyle, BorderSide.TOP, this.borderTopColor, colorPicker);
        ExcelColorUtils.setBorderColor(
                cellStyle, BorderSide.BOTTOM, this.borderBottomColor, colorPicker);
        ExcelColorUtils.setBorderColor(
                cellStyle, BorderSide.LEFT, this.borderLeftColor, colorPicker);
        ExcelColorUtils.setBorderColor(
                cellStyle, BorderSide.RIGHT, this.borderRightColor, colorPicker);
        cellStyle.setAlignment(this.horizontalAlignment);
        cellStyle.setVerticalAlignment(this.verticalAlignment);
        ExcelColorUtils.setBackgroundColor(cellStyle, this.backgroundColor, colorPicker);

        if (null != this.dataFormat && !this.dataFormat.isEmpty()) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(this.dataFormat));
        }

        Font font = workbook.createFont();
        if (null != this.fontName && !this.fontName.isEmpty()) {
            font.setFontName(this.fontName);
        }
        font.setFontHeightInPoints(this.fontSize.shortValue());
        font.setBold(this.bold);
        font.setItalic(this.italic);
        font.setStrikeout(this.strikeout);
        ExcelColorUtils.setFontColor(font, this.fontColor, colorPicker);
        cellStyle.setFont(font);
        return cellStyle;
    }
}
