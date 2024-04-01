package io.github.af19git5.entity;

import lombok.Getter;
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
    @Setter private String name = "";

    /** 欄位資料 */
    @Setter private List<ExcelCell> cellList = new ArrayList<>();

    /** 合併欄位規則 */
    @Setter private List<ExcelMergedRegion> mergedRegionList = new ArrayList<>();

    /** 隱藏列數 */
    @Setter private Set<Integer> hiddenRowNumList = new HashSet<>();

    /** 隱藏欄數 */
    @Setter private Set<Integer> hiddenColumnNumList = new HashSet<>();

    /** 覆寫欄位寬度Map */
    private final Map<Integer, Integer> overrideColumnWidthMap = new HashMap<>();

    public ExcelSheet() {}

    public ExcelSheet(String name, List<ExcelCell> cellList) {
        this.name = name;
        this.cellList = cellList;
    }

    public void addHiddenRowNum(int hiddenRowNum) {
        this.hiddenRowNumList.add(hiddenRowNum);
    }

    public void addHiddenColumnNum(int hiddenColumnNum) {
        this.hiddenColumnNumList.add(hiddenColumnNum);
    }

    public void addOverrideColumnWidth(int columnNum, int width) {
        this.overrideColumnWidthMap.put(columnNum, width);
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
