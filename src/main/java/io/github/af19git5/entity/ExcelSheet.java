package io.github.af19git5.entity;

import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel頁籤資料
 *
 * @author Jimmy Kang
 */
@Data
public class ExcelSheet {

    /** 頁籤名稱 */
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
}
