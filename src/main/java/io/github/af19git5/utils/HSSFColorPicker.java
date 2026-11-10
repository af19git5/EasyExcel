package io.github.af19git5.utils;

import lombok.NonNull;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * xls調色盤顏色配置
 *
 * <p>xls只能使用調色盤中的56個顏色，此類別會依序處理：
 *
 * <ol>
 *   <li>調色盤已有完全相同的顏色時直接使用
 *   <li>改寫本活頁簿尚未使用的調色盤位置為指定顏色
 *   <li>調色盤位置用完時，改用最接近的顏色
 * </ol>
 *
 * @author Jimmy Kang
 */
public class HSSFColorPicker {

    /** 調色盤第一個可自訂位置 */
    private static final short FIRST_INDEX = 0x08;

    /** 調色盤最後一個可自訂位置 */
    private static final short LAST_INDEX = 0x3F;

    private final HSSFPalette palette;

    /** 色碼對應調色盤位置 */
    private final Map<String, Short> indexMap = new HashMap<>();

    /** 已使用的調色盤位置 */
    private final Set<Short> usedIndexSet = new HashSet<>();

    /** 下一個嘗試改寫的調色盤位置 */
    private short nextIndex = LAST_INDEX;

    public HSSFColorPicker(@NonNull HSSFWorkbook workbook) {
        this.palette = workbook.getCustomPalette();
        // 保留預設黑白色，避免影響未指定顏色的欄位
        usedIndexSet.add(IndexedColors.BLACK.getIndex());
        usedIndexSet.add(IndexedColors.WHITE.getIndex());
    }

    /**
     * 取得色碼對應的調色盤位置
     *
     * @param colorHex 16進位色碼
     * @return 調色盤位置
     */
    public short getIndex(@NonNull String colorHex) {
        String hex = ExcelColorUtils.normalizeHex(colorHex);
        if (null == hex) {
            throw new IllegalArgumentException("Invalid color hex: " + colorHex);
        }
        Short cachedIndex = indexMap.get(hex);
        if (null != cachedIndex) {
            return cachedIndex;
        }
        byte[] rgb = ExcelColorUtils.toRGB(hex);
        short index;
        HSSFColor color = palette.findColor(rgb[0], rgb[1], rgb[2]);
        if (null != color) {
            index = color.getIndex();
        } else {
            while (nextIndex >= FIRST_INDEX && usedIndexSet.contains(nextIndex)) {
                nextIndex--;
            }
            if (nextIndex >= FIRST_INDEX) {
                index = nextIndex--;
                palette.setColorAtIndex(index, rgb[0], rgb[1], rgb[2]);
            } else {
                index =
                        palette.findSimilarColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF)
                                .getIndex();
            }
        }
        usedIndexSet.add(index);
        indexMap.put(hex, index);
        return index;
    }
}
