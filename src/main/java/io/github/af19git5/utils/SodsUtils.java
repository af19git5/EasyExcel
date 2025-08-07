package io.github.af19git5.utils;

import com.github.miachm.sods.Borders;
import com.github.miachm.sods.Style;

import io.github.af19git5.entity.ExcelMergedRegion;
import io.github.af19git5.entity.ExcelStreamStyle;
import io.github.af19git5.entity.ExcelStyle;
import io.github.af19git5.entity.StreamMergedRegion;
import io.github.af19git5.type.IndexedColorHex;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.Map;

/**
 * Excel樣式轉換Sods樣式
 *
 * @author Jimmy Kang
 */
public class SodsUtils {

    /** 邊線種類 */
    public enum BORDER_TYPE {
        TOP,
        BOTTOM,
        LEFT,
        RIGHT
    }

    /** 對應sods邊線寬度 */
    private static final Map<BorderStyle, String> borderWidths =
            Map.ofEntries(
                    Map.entry(BorderStyle.HAIR, "0.02cm"),
                    Map.entry(BorderStyle.THIN, "0.035cm"),
                    Map.entry(BorderStyle.DOTTED, "0.035cm"),
                    Map.entry(BorderStyle.DASHED, "0.035cm"),
                    Map.entry(BorderStyle.DASH_DOT, "0.035cm"),
                    Map.entry(BorderStyle.DASH_DOT_DOT, "0.035cm"),
                    Map.entry(BorderStyle.MEDIUM, "0.07cm"),
                    Map.entry(BorderStyle.MEDIUM_DASHED, "0.07cm"),
                    Map.entry(BorderStyle.MEDIUM_DASH_DOT, "0.07cm"),
                    Map.entry(BorderStyle.MEDIUM_DASH_DOT_DOT, "0.07cm"),
                    Map.entry(BorderStyle.THICK, "0.09cm"),
                    Map.entry(BorderStyle.DOUBLE, "0.07cm"));

    /** 對應sods邊線樣式 */
    private static final Map<BorderStyle, String> borderLineStyles =
            Map.ofEntries(
                    Map.entry(BorderStyle.HAIR, "solid"),
                    Map.entry(BorderStyle.THIN, "solid"),
                    Map.entry(BorderStyle.DOTTED, "dotted"),
                    Map.entry(BorderStyle.DASHED, "dashed"),
                    Map.entry(BorderStyle.DASH_DOT, "dashed"),
                    Map.entry(BorderStyle.DASH_DOT_DOT, "dashed"),
                    Map.entry(BorderStyle.MEDIUM, "solid"),
                    Map.entry(BorderStyle.MEDIUM_DASHED, "dashed"),
                    Map.entry(BorderStyle.MEDIUM_DASH_DOT, "dashed"),
                    Map.entry(BorderStyle.MEDIUM_DASH_DOT_DOT, "dashed"),
                    Map.entry(BorderStyle.THICK, "solid"),
                    Map.entry(BorderStyle.DOUBLE, "double"));

    /**
     * EasyExcel樣式轉為SODS樣式
     *
     * @param excelStyle Excel欄位樣式
     * @return SODS樣式
     */
    public static Style toSodsStyle(ExcelStyle excelStyle) {
        Style style = new Style();
        style.setWrap(excelStyle.getIsWrapText());
        Borders borders =
                new Borders(
                        isBorder(excelStyle.getBorderTop()),
                        isBorder(excelStyle.getBorderBottom()),
                        isBorder(excelStyle.getBorderLeft()),
                        isBorder(excelStyle.getBorderRight()));
        boolean isBorder =
                isBorder(excelStyle.getBorderTop())
                        && isBorder(excelStyle.getBorderBottom())
                        && isBorder(excelStyle.getBorderLeft())
                        && isBorder(excelStyle.getBorderRight());
        borders.setBorder(isBorder);
        if (isBorder) {
            borders.setBorderTopProperties(getBorderProperties(excelStyle, BORDER_TYPE.TOP));
            borders.setBorderBottomProperties(getBorderProperties(excelStyle, BORDER_TYPE.BOTTOM));
            borders.setBorderLeftProperties(getBorderProperties(excelStyle, BORDER_TYPE.LEFT));
            borders.setBorderRightProperties(getBorderProperties(excelStyle, BORDER_TYPE.RIGHT));
        }

        style.setBorders(borders);
        style.setTextAligment(toTextAlignment(excelStyle.getHorizontalAlignment()));
        style.setVerticalTextAligment(toVerticalTextAlignment(excelStyle.getVerticalAlignment()));
        if (null != excelStyle.getBackgroundColor()) {
            style.setBackgroundColor(
                    new com.github.miachm.sods.Color(excelStyle.getBackgroundColor()));
        }
        if (null != excelStyle.getFontColor()) {
            style.setFontColor(new com.github.miachm.sods.Color(excelStyle.getFontColor()));
        }
        style.setFontSize(excelStyle.getFontSize());
        style.setBold(excelStyle.getBold());
        style.setItalic(excelStyle.getItalic());

        return style;
    }

    /**
     * Excel欄位樣式轉為SODS樣式
     *
     * @param excelStyle Excel欄位樣式
     * @return SODS樣式
     */
    public static Style toSodsStyle(ExcelStreamStyle excelStyle) {
        Style style = new Style();
        style.setWrap(excelStyle.getIsWrapText());
        Borders borders =
                new Borders(
                        isBorder(excelStyle.getBorderTop()),
                        isBorder(excelStyle.getBorderBottom()),
                        isBorder(excelStyle.getBorderLeft()),
                        isBorder(excelStyle.getBorderRight()));
        boolean isBorder =
                isBorder(excelStyle.getBorderTop())
                        && isBorder(excelStyle.getBorderBottom())
                        && isBorder(excelStyle.getBorderLeft())
                        && isBorder(excelStyle.getBorderRight());
        borders.setBorder(isBorder);
        if (isBorder) {
            borders.setBorderTopProperties(getBorderProperties(excelStyle, BORDER_TYPE.TOP));
            borders.setBorderBottomProperties(getBorderProperties(excelStyle, BORDER_TYPE.BOTTOM));
            borders.setBorderLeftProperties(getBorderProperties(excelStyle, BORDER_TYPE.LEFT));
            borders.setBorderRightProperties(getBorderProperties(excelStyle, BORDER_TYPE.RIGHT));
        }

        style.setBorders(borders);
        style.setTextAligment(toTextAlignment(excelStyle.getHorizontalAlignment()));
        style.setVerticalTextAligment(toVerticalTextAlignment(excelStyle.getVerticalAlignment()));
        if (null != excelStyle.getBackgroundColor()) {
            style.setBackgroundColor(
                    new com.github.miachm.sods.Color(
                            IndexedColorHex.convertToHex(excelStyle.getBackgroundColor())));
        }
        if (null != excelStyle.getFontColor()) {
            style.setFontColor(
                    new com.github.miachm.sods.Color(
                            IndexedColorHex.convertToHex(excelStyle.getFontColor())));
        }
        style.setFontSize(excelStyle.getFontSize());
        style.setBold(excelStyle.getBold());
        style.setItalic(excelStyle.getItalic());

        return style;
    }

    /**
     * 是否有邊線
     *
     * @param borderStyle apache邊線樣式
     * @return 是否有邊線
     */
    public static boolean isBorder(BorderStyle borderStyle) {
        return !borderStyle.equals(BorderStyle.NONE);
    }

    /**
     * Excel欄位樣式轉換成sods邊線屬性
     *
     * @param excelStyle Excel欄位樣式
     * @param type 上下左右邊線
     * @return sods邊線屬性
     */
    public static String getBorderProperties(ExcelStyle excelStyle, BORDER_TYPE type) {
        BorderStyle style = null;
        String color = null;
        switch (type) {
            case TOP:
                style = excelStyle.getBorderTop();
                color = excelStyle.getBorderTopColor();
                break;
            case BOTTOM:
                style = excelStyle.getBorderBottom();
                color = excelStyle.getBorderBottomColor();
                break;
            case LEFT:
                style = excelStyle.getBorderLeft();
                color = excelStyle.getBorderLeftColor();
                break;
            case RIGHT:
                style = excelStyle.getBorderRight();
                color = excelStyle.getBorderRightColor();
                break;
        }

        // 從 Map 取粗細與樣式，沒有找到給預設值
        String width = borderWidths.getOrDefault(style, "0.035cm");
        String lineStyle = borderLineStyles.getOrDefault(style, "solid");

        return width + " " + lineStyle + " " + (null != color ? color : "#000000");
    }

    /**
     * Excel欄位樣式轉換成sods邊線屬性
     *
     * @param excelStyle Excel欄位樣式
     * @param type 上下左右邊線
     * @return sods邊線屬性
     */
    public static String getBorderProperties(ExcelStreamStyle excelStyle, BORDER_TYPE type) {
        BorderStyle style = null;
        String color = null;
        switch (type) {
            case TOP:
                style = excelStyle.getBorderTop();
                color = IndexedColorHex.convertToHex(excelStyle.getBorderTopColor());
                break;
            case BOTTOM:
                style = excelStyle.getBorderBottom();
                color = IndexedColorHex.convertToHex(excelStyle.getBorderBottomColor());
                break;
            case LEFT:
                style = excelStyle.getBorderLeft();
                color = IndexedColorHex.convertToHex(excelStyle.getBorderLeftColor());
                break;
            case RIGHT:
                style = excelStyle.getBorderRight();
                color = IndexedColorHex.convertToHex(excelStyle.getBorderRightColor());
                break;
        }

        // 從 Map 取粗細與樣式，沒有找到給預設值
        String width = borderWidths.getOrDefault(style, "0.035cm");
        String lineStyle = borderLineStyles.getOrDefault(style, "solid");

        return width + " " + lineStyle + " " + (null != color ? color : "#000000");
    }

    /**
     * ExcelMergedRegion轉換成sods邊線屬性
     *
     * @param excelMergedRegion Excel合併欄位規則
     * @param type 上下左右邊線
     * @return sods邊線屬性
     */
    public static String getBorderProperties(
            ExcelMergedRegion excelMergedRegion, BORDER_TYPE type) {
        BorderStyle style = null;
        String color = null;
        switch (type) {
            case TOP:
                style = excelMergedRegion.getBorderTop();
                color = excelMergedRegion.getBorderTopColor();
                break;
            case BOTTOM:
                style = excelMergedRegion.getBorderBottom();
                color = excelMergedRegion.getBorderBottomColor();
                break;
            case LEFT:
                style = excelMergedRegion.getBorderLeft();
                color = excelMergedRegion.getBorderLeftColor();
                break;
            case RIGHT:
                style = excelMergedRegion.getBorderRight();
                color = excelMergedRegion.getBorderRightColor();
                break;
        }

        // 從 Map 取粗細與樣式，沒有找到給預設值
        String width = borderWidths.getOrDefault(style, "0.035cm");
        String lineStyle = borderLineStyles.getOrDefault(style, "solid");

        return width + " " + lineStyle + " " + (null != color ? color : "#000000");
    }

    /**
     * StreamMergedRegion轉換成sods邊線屬性
     *
     * @param streamMergedRegion Excel合併欄位規則
     * @param type 上下左右邊線
     * @return sods邊線屬性
     */
    public static String getBorderProperties(
            StreamMergedRegion streamMergedRegion, BORDER_TYPE type) {
        BorderStyle style = null;
        String color = null;
        switch (type) {
            case TOP:
                style = streamMergedRegion.getBorderTop();
                color = IndexedColorHex.convertToHex(streamMergedRegion.getBorderTopColor());
                break;
            case BOTTOM:
                style = streamMergedRegion.getBorderBottom();
                color = IndexedColorHex.convertToHex(streamMergedRegion.getBorderBottomColor());
                break;
            case LEFT:
                style = streamMergedRegion.getBorderLeft();
                color = IndexedColorHex.convertToHex(streamMergedRegion.getBorderLeftColor());
                break;
            case RIGHT:
                style = streamMergedRegion.getBorderRight();
                color = IndexedColorHex.convertToHex(streamMergedRegion.getBorderRightColor());
                break;
        }

        // 從 Map 取粗細與樣式，沒有找到給預設值
        String width = borderWidths.getOrDefault(style, "0.035cm");
        String lineStyle = borderLineStyles.getOrDefault(style, "solid");

        return width + " " + lineStyle + " " + (null != color ? color : "#000000");
    }

    /**
     * 對應sods水平對齊
     *
     * @param align apache水平對齊
     * @return sods水平對齊
     */
    public static Style.TEXT_ALIGMENT toTextAlignment(HorizontalAlignment align) {
        if (null == align) {
            return Style.TEXT_ALIGMENT.Left;
        }

        switch (align) {
            case CENTER:
            case CENTER_SELECTION:
                return Style.TEXT_ALIGMENT.Center;
            case RIGHT:
                return Style.TEXT_ALIGMENT.Right;
            case LEFT:
            case GENERAL:
            case FILL:
            case JUSTIFY:
            case DISTRIBUTED:
            default:
                return Style.TEXT_ALIGMENT.Left;
        }
    }

    /**
     * 對應sods垂直對齊
     *
     * @param align apache垂直對齊
     * @return sods垂直對齊
     */
    public static Style.VERTICAL_TEXT_ALIGMENT toVerticalTextAlignment(VerticalAlignment align) {
        if (null == align) {
            return Style.VERTICAL_TEXT_ALIGMENT.Middle;
        }

        switch (align) {
            case TOP:
                return Style.VERTICAL_TEXT_ALIGMENT.Top;
            case BOTTOM:
                return Style.VERTICAL_TEXT_ALIGMENT.Bottom;
            case JUSTIFY:
            case DISTRIBUTED:
            default:
                return Style.VERTICAL_TEXT_ALIGMENT.Middle;
        }
    }
}
