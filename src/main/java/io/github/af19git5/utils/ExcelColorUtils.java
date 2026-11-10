package io.github.af19git5.utils;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

/**
 * Excel顏色處理工具
 *
 * @author Jimmy Kang
 */
public class ExcelColorUtils {

    private ExcelColorUtils() {}

    /**
     * 將16進位色碼正規化為#RRGGBB格式
     *
     * <p>支援#RGB、#RRGGBB、#AARRGGBB(忽略透明度)，可省略#
     *
     * @param colorHex 16進位色碼
     * @return #RRGGBB格式色碼，格式不符時回傳null
     */
    public static String normalizeHex(String colorHex) {
        if (null == colorHex) {
            return null;
        }
        String hex = colorHex.trim();
        if (hex.startsWith("#")) {
            hex = hex.substring(1);
        }
        if (!hex.matches("^(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$")) {
            return null;
        }
        if (hex.length() == 3) {
            hex =
                    new String(
                            new char[] {
                                hex.charAt(0), hex.charAt(0),
                                hex.charAt(1), hex.charAt(1),
                                hex.charAt(2), hex.charAt(2)
                            });
        } else if (hex.length() == 8) {
            hex = hex.substring(2);
        }
        return "#" + hex.toUpperCase();
    }

    /**
     * 將IndexedColors轉為#RRGGBB格式色碼
     *
     * @param indexedColor 預設索引色
     * @return #RRGGBB格式色碼，無對應RGB(例如AUTOMATIC)時回傳null
     */
    public static String toHex(IndexedColors indexedColor) {
        if (null == indexedColor) {
            return null;
        }
        byte[] rgb = DefaultIndexedColorMap.getDefaultRGB(indexedColor.getIndex());
        if (null == rgb) {
            return null;
        }
        return String.format("#%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    /**
     * 將#RRGGBB格式色碼轉為RGB陣列
     *
     * @param colorHex 16進位色碼
     * @return RGB陣列
     */
    public static byte[] toRGB(String colorHex) {
        String hex = normalizeHex(colorHex);
        if (null == hex) {
            throw new IllegalArgumentException("Invalid color hex: " + colorHex);
        }
        return new byte[] {
            (byte) Integer.parseInt(hex.substring(1, 3), 16),
            (byte) Integer.parseInt(hex.substring(3, 5), 16),
            (byte) Integer.parseInt(hex.substring(5, 7), 16)
        };
    }

    /**
     * 將16進位色碼轉為XSSFColor
     *
     * @param colorHex 16進位色碼
     * @return XSSFColor
     */
    public static XSSFColor toXSSFColor(String colorHex) {
        return new XSSFColor(toRGB(colorHex), new DefaultIndexedColorMap());
    }

    /**
     * 設定邊線顏色
     *
     * @param cellStyle 欄位樣式
     * @param side 邊線位置
     * @param colorHex 16進位色碼
     * @param colorPicker xls調色盤(xlsx時可為null)
     */
    public static void setBorderColor(
            CellStyle cellStyle, BorderSide side, String colorHex, HSSFColorPicker colorPicker) {
        if (null == colorHex) {
            return;
        }
        if (cellStyle instanceof XSSFCellStyle) {
            ((XSSFCellStyle) cellStyle).setBorderColor(side, toXSSFColor(colorHex));
            return;
        }
        short index = requirePicker(cellStyle, colorPicker).getIndex(colorHex);
        switch (side) {
            case TOP:
                cellStyle.setTopBorderColor(index);
                break;
            case BOTTOM:
                cellStyle.setBottomBorderColor(index);
                break;
            case LEFT:
                cellStyle.setLeftBorderColor(index);
                break;
            case RIGHT:
                cellStyle.setRightBorderColor(index);
                break;
            default:
                break;
        }
    }

    /**
     * 設定背景顏色(實心填滿)
     *
     * @param cellStyle 欄位樣式
     * @param colorHex 16進位色碼
     * @param colorPicker xls調色盤(xlsx時可為null)
     */
    public static void setBackgroundColor(
            CellStyle cellStyle, String colorHex, HSSFColorPicker colorPicker) {
        if (null == colorHex) {
            return;
        }
        if (cellStyle instanceof XSSFCellStyle) {
            ((XSSFCellStyle) cellStyle).setFillForegroundColor(toXSSFColor(colorHex));
        } else {
            cellStyle.setFillForegroundColor(
                    requirePicker(cellStyle, colorPicker).getIndex(colorHex));
        }
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    /**
     * 設定文字顏色
     *
     * @param font 字型
     * @param colorHex 16進位色碼
     * @param colorPicker xls調色盤(xlsx時可為null)
     */
    public static void setFontColor(Font font, String colorHex, HSSFColorPicker colorPicker) {
        if (null == colorHex) {
            return;
        }
        if (font instanceof XSSFFont) {
            ((XSSFFont) font).setColor(toXSSFColor(colorHex));
        } else if (font instanceof HSSFFont) {
            if (null == colorPicker) {
                throw new IllegalArgumentException("HSSFColorPicker is required for xls.");
            }
            font.setColor(colorPicker.getIndex(colorHex));
        }
    }

    private static HSSFColorPicker requirePicker(
            CellStyle cellStyle, HSSFColorPicker colorPicker) {
        if (cellStyle instanceof HSSFCellStyle && null == colorPicker) {
            throw new IllegalArgumentException("HSSFColorPicker is required for xls.");
        }
        return colorPicker;
    }
}
