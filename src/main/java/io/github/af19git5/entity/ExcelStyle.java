package io.github.af19git5.entity;

import io.github.af19git5.builder.ExcelStyleBuilder;

import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

import java.awt.Color;

/**
 * Excel欄位樣式資料
 *
 * @author Jimmy Kang
 */
@Getter
@Setter
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

    public static ExcelStyleBuilder init() {
        return new ExcelStyleBuilder();
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

    public void setBackgroundColor(@NonNull String colorHex) {
        if (isValidColorHex(colorHex)) {
            this.backgroundColor = colorHex;
        }
    }

    public void setFontColor(@NonNull String colorHex) {
        if (isValidColorHex(colorHex)) {
            this.fontColor = colorHex;
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

    public void setAllBorder(@NonNull BorderStyle borderStyle) {
        this.borderTop = borderStyle;
        this.borderBottom = borderStyle;
        this.borderLeft = borderStyle;
        this.borderRight = borderStyle;
    }

    public void setAllBorderColor(@NonNull String colorHex) {
        if (isValidColorHex(colorHex)) {
            this.borderTopColor = colorHex;
            this.borderBottomColor = colorHex;
            this.borderLeftColor = colorHex;
            this.borderRightColor = colorHex;
        }
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
                convertRGBToHex(palette.getColor(cellStyle.getTopBorderColor()).getTriplet());
        this.borderBottomColor =
                convertRGBToHex(palette.getColor(cellStyle.getBottomBorderColor()).getTriplet());
        this.borderLeftColor =
                convertRGBToHex(palette.getColor(cellStyle.getLeftBorderColor()).getTriplet());
        this.borderRightColor =
                convertRGBToHex(palette.getColor(cellStyle.getRightBorderColor()).getTriplet());
        this.horizontalAlignment = cellStyle.getAlignment();
        this.verticalAlignment = cellStyle.getVerticalAlignment();
        this.backgroundColor =
                convertRGBToHex(cellStyle.getFillForegroundColorColor().getTriplet());
        HSSFFont font = cellStyle.getFont(workbook);
        this.fontName = font.getFontName();
        this.fontSize = (int) font.getFontHeightInPoints();
        this.bold = font.getBold();
        this.italic = font.getItalic();
        this.strikeout = font.getStrikeout();
        HSSFColor hssfColor = font.getHSSFColor(workbook);
        if (null != hssfColor) {
            this.fontColor = convertRGBToHex(hssfColor.getTriplet());
        }
    }

    private String convertRGBToHex(short[] rgbArray) {
        if (rgbArray.length != 3) return null;
        try {
            String hexRed = String.format("%02X", rgbArray[0]);
            String hexGreen = String.format("%02X", rgbArray[1]);
            String hexBlue = String.format("%02X", rgbArray[2]);
            return "#" + hexRed + hexGreen + hexBlue;
        } catch (NumberFormatException e) {
            return null;
        }
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
        this.backgroundColor = convertXSSColorToHax(cellStyle.getFillForegroundColorColor());
        XSSFFont font = cellStyle.getFont();
        this.fontName = font.getFontName();
        this.fontSize = (int) font.getFontHeightInPoints();
        this.bold = font.getBold();
        this.italic = font.getItalic();
        this.strikeout = font.getStrikeout();
        this.fontColor = convertXSSColorToHax(font.getXSSFColor());
    }

    private String convertXSSColorToHax(XSSFColor color) {
        if (null == color) return null;
        if (null == color.getARGBHex()) return null;
        return "#" + color.getARGBHex().substring(2);
    }

    public HSSFCellStyle toHSSCellStyle(@NonNull HSSFWorkbook workbook) {
        HSSFPalette palette = workbook.getCustomPalette();
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setLocked(this.isLock);
        cellStyle.setWrapText(this.isWrapText);
        cellStyle.setBorderTop(this.borderTop);
        cellStyle.setBorderBottom(this.borderBottom);
        cellStyle.setBorderLeft(this.borderLeft);
        cellStyle.setBorderRight(this.borderRight);

        if (null != this.borderTopColor) {
            Color rgbColor = Color.decode(this.borderTopColor);
            HSSFColor color =
                    palette.findSimilarColor(
                            (byte) rgbColor.getRed(),
                            (byte) rgbColor.getGreen(),
                            (byte) rgbColor.getBlue());
            cellStyle.setTopBorderColor(color.getIndex());
        }

        if (null != this.borderBottomColor) {
            Color rgbColor = Color.decode(this.borderBottomColor);
            HSSFColor color =
                    palette.findSimilarColor(
                            (byte) rgbColor.getRed(),
                            (byte) rgbColor.getGreen(),
                            (byte) rgbColor.getBlue());
            cellStyle.setBottomBorderColor(color.getIndex());
        }

        if (null != this.borderLeftColor) {
            Color rgbColor = Color.decode(this.borderLeftColor);
            HSSFColor color =
                    palette.findSimilarColor(
                            (byte) rgbColor.getRed(),
                            (byte) rgbColor.getGreen(),
                            (byte) rgbColor.getBlue());
            cellStyle.setLeftBorderColor(color.getIndex());
        }

        if (null != this.borderRightColor) {
            Color rgbColor = Color.decode(this.borderRightColor);
            HSSFColor color =
                    palette.findSimilarColor(
                            (byte) rgbColor.getRed(),
                            (byte) rgbColor.getGreen(),
                            (byte) rgbColor.getBlue());
            cellStyle.setRightBorderColor(color.getIndex());
        }

        cellStyle.setAlignment(this.horizontalAlignment);
        cellStyle.setVerticalAlignment(this.verticalAlignment);

        if (null != this.backgroundColor) {
            Color rgbColor = Color.decode(this.backgroundColor);
            HSSFColor color =
                    palette.findSimilarColor(
                            (byte) rgbColor.getRed(),
                            (byte) rgbColor.getGreen(),
                            (byte) rgbColor.getBlue());
            cellStyle.setFillForegroundColor(color.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        HSSFFont font = workbook.createFont();
        if (null != this.fontName && !this.fontName.isEmpty()) {
            font.setFontName(this.fontName);
        }
        font.setFontHeightInPoints(this.fontSize.shortValue());
        font.setBold(this.bold);
        font.setItalic(this.italic);
        font.setStrikeout(this.strikeout);

        if (null != this.fontColor) {
            Color rgbColor = Color.decode(this.fontColor);
            HSSFColor hssfColor =
                    palette.findSimilarColor(
                            rgbColor.getRed(), rgbColor.getGreen(), rgbColor.getBlue());
            font.setColor(hssfColor.getIndex());
        }
        cellStyle.setFont(font);
        return cellStyle;
    }

    public CellStyle toXSSCellStyle(@NonNull XSSFWorkbook workbook) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setLocked(this.isLock);
        cellStyle.setWrapText(this.isWrapText);
        cellStyle.setBorderTop(this.borderTop);
        cellStyle.setBorderBottom(this.borderBottom);
        cellStyle.setBorderLeft(this.borderLeft);
        cellStyle.setBorderRight(this.borderRight);

        if (null != this.borderTopColor) {
            XSSFColor color =
                    XSSFColor.from(CTColor.Factory.newInstance(), new DefaultIndexedColorMap());
            color.setARGBHex(this.borderTopColor.substring(1));
            cellStyle.setTopBorderColor(color);
        }

        if (null != this.borderBottomColor) {
            XSSFColor color =
                    XSSFColor.from(CTColor.Factory.newInstance(), new DefaultIndexedColorMap());
            color.setARGBHex(this.borderBottomColor.substring(1));
            cellStyle.setBottomBorderColor(color);
        }

        if (null != this.borderLeftColor) {
            XSSFColor color =
                    XSSFColor.from(CTColor.Factory.newInstance(), new DefaultIndexedColorMap());
            color.setARGBHex(this.borderLeftColor.substring(1));
            cellStyle.setLeftBorderColor(color);
        }

        if (null != this.borderRightColor) {
            XSSFColor color =
                    XSSFColor.from(CTColor.Factory.newInstance(), new DefaultIndexedColorMap());
            color.setARGBHex(this.borderRightColor.substring(1));
            cellStyle.setRightBorderColor(color);
        }

        cellStyle.setAlignment(this.horizontalAlignment);
        cellStyle.setVerticalAlignment(this.verticalAlignment);

        if (null != this.backgroundColor) {
            XSSFColor color =
                    XSSFColor.from(CTColor.Factory.newInstance(), new DefaultIndexedColorMap());
            color.setARGBHex(this.backgroundColor.substring(1));
            cellStyle.setFillForegroundColor(color);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        XSSFFont font = workbook.createFont();
        if (null != this.fontName && !this.fontName.isEmpty()) {
            font.setFontName(this.fontName);
        }
        font.setFontHeightInPoints(this.fontSize.shortValue());
        font.setBold(this.bold);
        font.setItalic(this.italic);
        font.setStrikeout(this.strikeout);

        if (null != this.fontColor) {
            XSSFColor color =
                    XSSFColor.from(CTColor.Factory.newInstance(), new DefaultIndexedColorMap());
            color.setARGBHex(this.fontColor.substring(1));
            font.setColor(color);
        }
        cellStyle.setFont(font);
        return cellStyle;
    }
}
