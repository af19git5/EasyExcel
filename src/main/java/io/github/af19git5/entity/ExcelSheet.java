package io.github.af19git5.entity;

import lombok.Data;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

/**
 * Excel工作表資料
 *
 * @author Jimmy Kang
 */
@Data
public class ExcelSheet {

    /** 工作表名稱 */
    private String name = "";

    /** 欄位資料 */
    private List<ExcelCell> cellList = new ArrayList<>();

    /** 合併欄位規則 */
    private List<CellRangeAddress> cellRangeAddressList = new ArrayList<>();

    public ExcelSheet() {}

    public ExcelSheet(String name, List<ExcelCell> cellList) {
        this.name = name;
        this.cellList = cellList;
    }

    /**
     * 將cell陣列資料轉為二維資料陣列(逐列讀出)
     *
     * @return 二維資料陣列
     */
    public List<List<String>> toValueList() {
        List<ExcelCell> cellList = new ArrayList<>(this.cellList);
        cellList.sort(new CellComparator());
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

    private static class CellComparator implements Comparator<ExcelCell> {
        @Override
        public int compare(ExcelCell cell1, ExcelCell cell2) {
            int result = Integer.compare(cell1.getRow(), cell2.getRow());
            if (result == 0) {
                result = Integer.compare(cell1.getColumn(), cell2.getColumn());
            }
            return result;
        }
    }
}
