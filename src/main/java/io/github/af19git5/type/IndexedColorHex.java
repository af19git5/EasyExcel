package io.github.af19git5.type;

import lombok.AllArgsConstructor;
import lombok.Getter;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.HashMap;
import java.util.Map;

/**
 * IndexedColors對應16進位色碼
 *
 * @author Cindy Hsu
 */
@AllArgsConstructor
@Getter
public enum IndexedColorHex {
    BLACK1(IndexedColors.BLACK1, "#000000"),
    WHITE1(IndexedColors.WHITE1, "#FFFFFF"),
    RED1(IndexedColors.RED1, "#FF0000"),
    BRIGHT_GREEN1(IndexedColors.BRIGHT_GREEN1, "#00FF00"),
    BLUE1(IndexedColors.BLUE1, "#0000FF"),
    YELLOW1(IndexedColors.YELLOW1, "#FFFF00"),
    PINK1(IndexedColors.PINK1, "#FF00FF"),
    TURQUOISE1(IndexedColors.TURQUOISE1, "#00FFFF"),
    BLACK(IndexedColors.BLACK, "#000000"),
    WHITE(IndexedColors.WHITE, "#FFFFFF"),
    RED(IndexedColors.RED, "#FF0000"),
    BRIGHT_GREEN(IndexedColors.BRIGHT_GREEN, "#00FF00"),
    BLUE(IndexedColors.BLUE, "#0000FF"),
    YELLOW(IndexedColors.YELLOW, "#FFFF00"),
    PINK(IndexedColors.PINK, "#FF00FF"),
    TURQUOISE(IndexedColors.TURQUOISE, "#00FFFF"),
    DARK_RED(IndexedColors.DARK_RED, "#800000"),
    GREEN(IndexedColors.GREEN, "#008000"),
    DARK_BLUE(IndexedColors.DARK_BLUE, "#000080"),
    DARK_YELLOW(IndexedColors.DARK_YELLOW, "#808000"),
    VIOLET(IndexedColors.VIOLET, "#800080"),
    TEAL(IndexedColors.TEAL, "#008080"),
    GREY_25_PERCENT(IndexedColors.GREY_25_PERCENT, "#C0C0C0"),
    GREY_50_PERCENT(IndexedColors.GREY_50_PERCENT, "#808080"),
    CORNFLOWER_BLUE(IndexedColors.CORNFLOWER_BLUE, "#9999FF"),
    MAROON(IndexedColors.MAROON, "#800000"),
    LEMON_CHIFFON(IndexedColors.LEMON_CHIFFON, "#FFFFCC"),
    LIGHT_TURQUOISE1(IndexedColors.LIGHT_TURQUOISE1, "#CCFFFF"),
    ORCHID(IndexedColors.ORCHID, "#660066"),
    CORAL(IndexedColors.CORAL, "#FF8080"),
    ROYAL_BLUE(IndexedColors.ROYAL_BLUE, "#0066CC"),
    LIGHT_CORNFLOWER_BLUE(IndexedColors.LIGHT_CORNFLOWER_BLUE, "#CCCCFF"),
    SKY_BLUE(IndexedColors.SKY_BLUE, "#00CCFF"),
    LIGHT_TURQUOISE(IndexedColors.LIGHT_TURQUOISE, "#CCFFFF"),
    LIGHT_GREEN(IndexedColors.LIGHT_GREEN, "#CCFFCC"),
    LIGHT_YELLOW(IndexedColors.LIGHT_YELLOW, "#FFFFCC"),
    PALE_BLUE(IndexedColors.PALE_BLUE, "#9999FF"),
    ROSE(IndexedColors.ROSE, "#FF99CC"),
    LAVENDER(IndexedColors.LAVENDER, "#9999FF"),
    TAN(IndexedColors.TAN, "#FFCC99"),
    LIGHT_BLUE(IndexedColors.LIGHT_BLUE, "#3366FF"),
    AQUA(IndexedColors.AQUA, "#33CCCC"),
    LIME(IndexedColors.LIME, "#99CC00"),
    GOLD(IndexedColors.GOLD, "#FFCC00"),
    LIGHT_ORANGE(IndexedColors.LIGHT_ORANGE, "#FF9900"),
    ORANGE(IndexedColors.ORANGE, "#FF6600"),
    BLUE_GREY(IndexedColors.BLUE_GREY, "#666699"),
    GREY_40_PERCENT(IndexedColors.GREY_40_PERCENT, "#969696"),
    DARK_TEAL(IndexedColors.DARK_TEAL, "#003366"),
    SEA_GREEN(IndexedColors.SEA_GREEN, "#339966"),
    DARK_GREEN(IndexedColors.DARK_GREEN, "#003300"),
    OLIVE_GREEN(IndexedColors.OLIVE_GREEN, "#333300"),
    BROWN(IndexedColors.BROWN, "#993300"),
    PLUM(IndexedColors.PLUM, "#993366"),
    INDIGO(IndexedColors.INDIGO, "#333399"),
    GREY_80_PERCENT(IndexedColors.GREY_80_PERCENT, "#333333"),
    AUTOMATIC(IndexedColors.AUTOMATIC, "#000000");

    private final IndexedColors indexedColor;

    private final String hex;

    private static final Map<IndexedColors, IndexedColorHex> lookupMap = new HashMap<>();

    static {
        for (IndexedColorHex value : values()) {
            lookupMap.put(value.indexedColor, value);
        }
    }

    public static String convertToHex(IndexedColors color) {
        return lookupMap.getOrDefault(color, AUTOMATIC).getHex();
    }
}
