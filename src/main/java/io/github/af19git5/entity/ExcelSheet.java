package io.github.af19git5.entity;

import io.github.af19git5.builder.ExcelSheetBuilder;

import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;

import java.util.*;

/**
 * Excel工作表資料
 *
 * @author Jimmy Kang
 */
@Getter
public class ExcelSheet {

    /** 工作表名稱 */
    @Setter @NonNull private String name = "";

    /** 欄位資料 */
    @Setter @NonNull private List<ExcelCell> cellList = new ArrayList<>();

    /** 合併欄位規則 */
    @Setter @NonNull private List<ExcelMergedRegion> mergedRegionList = new ArrayList<>();

    /** 隱藏列數 */
    @Setter @NonNull private Set<Integer> hiddenRowNumSet = new HashSet<>();

    /** 隱藏欄數 */
    @Setter @NonNull private Set<Integer> hiddenColumnNumSet = new HashSet<>();

    /** 是否被保護 */
    private Boolean isProtect = false;

    /** 密碼 */
    private String password = "";

    /** 凍結行數 */
    private Integer freezeColumnNum = 0;

    /** 凍結列數 */
    private Integer freezeRowNum = 0;

    /** 覆寫欄位寬度Map */
    private final Map<Integer, Integer> overrideColumnWidthMap = new HashMap<>();

    public ExcelSheet() {}

    public ExcelSheet(@NonNull String name, @NonNull List<ExcelCell> cellList) {
        this.name = name;
        this.cellList = cellList;
    }

    public static ExcelSheetBuilder init() {
        return new ExcelSheetBuilder();
    }

    public void addHiddenRowNum(int hiddenRowNum) {
        this.hiddenRowNumSet.add(hiddenRowNum);
    }

    public void addHiddenColumnNum(int hiddenColumnNum) {
        this.hiddenColumnNumSet.add(hiddenColumnNum);
    }

    public void addOverrideColumnWidth(int columnNum, int width) {
        this.overrideColumnWidthMap.put(columnNum, width);
    }

    public void protect(@NonNull String password) {
        this.isProtect = true;
        this.password = password;
    }

    public void freezePane(int columnNum, int rowNum) {
        this.freezeColumnNum = columnNum;
        this.freezeRowNum = rowNum;
    }
    /**
     * 將cell陣列資料轉為二維資料陣列(逐列讀出)
     *
     * @return 二維資料陣列
     */
    public List<List<String>> toValueList() {
        List<ExcelCell> cellList = new ArrayList<>(this.cellList);
        cellList.sort(
                Comparator.comparingInt(ExcelCell::getRow).thenComparingInt(ExcelCell::getColumn));
        Map<Integer, List<String>> dataMap = new LinkedHashMap<>();
        for (ExcelCell cell : cellList) {
            List<String> columnList =
                    dataMap.computeIfAbsent(cell.getRow(), k -> new ArrayList<>());
            columnList.add(cell.getValue());
        }
        List<List<String>> dataList = new ArrayList<>();
        dataMap.forEach((integer, columnList) -> dataList.add(columnList));
        return dataList;
    }
}
