package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelCell;
import io.github.af19git5.entity.ExcelMergedRegion;
import io.github.af19git5.entity.ExcelSheet;

import lombok.NonNull;

import java.util.List;
import java.util.Set;

/**
 * Excel工作表資料建構器
 *
 * @author Jimmy Kang
 */
public class ExcelSheetBuilder {

    private final ExcelSheet sheet;

    public ExcelSheetBuilder() {
        this.sheet = new ExcelSheet();
    }

    public ExcelSheetBuilder name(@NonNull String name) {
        sheet.setName(name);
        return this;
    }

    public ExcelSheetBuilder cells(@NonNull ExcelCell... cells) {
        sheet.getCellList().addAll(List.of(cells));
        return this;
    }

    public ExcelSheetBuilder cellList(@NonNull List<ExcelCell> cellList) {
        sheet.setCellList(cellList);
        return this;
    }

    public ExcelSheetBuilder mergedRegions(@NonNull ExcelMergedRegion... mergedRegions) {
        sheet.getMergedRegionList().addAll(List.of(mergedRegions));
        return this;
    }

    public ExcelSheetBuilder mergedRegionList(@NonNull List<ExcelMergedRegion> mergedRegionList) {
        sheet.setMergedRegionList(mergedRegionList);
        return this;
    }

    public ExcelSheetBuilder hiddenRowNums(@NonNull Integer... rowNums) {
        sheet.getHiddenRowNumSet().addAll(List.of(rowNums));
        return this;
    }

    public ExcelSheetBuilder hiddenRowNumSet(@NonNull Set<Integer> hiddenRowNumSet) {
        sheet.setHiddenRowNumSet(hiddenRowNumSet);
        return this;
    }

    public ExcelSheetBuilder hiddenColumnNums(@NonNull Integer... columnNums) {
        sheet.getHiddenColumnNumSet().addAll(List.of(columnNums));
        return this;
    }

    public ExcelSheetBuilder hiddenColumnNumSet(@NonNull Set<Integer> hiddenColumnNumSet) {
        sheet.setHiddenColumnNumSet(hiddenColumnNumSet);
        return this;
    }

    public ExcelSheetBuilder overrideColumnWidth(int columnNum, int width) {
        sheet.addOverrideColumnWidth(columnNum, width);
        return this;
    }

    public ExcelSheetBuilder protect(@NonNull String password) {
        sheet.protect(password);
        return this;
    }

    public ExcelSheetBuilder freezePane(int columnNum, int rowNum) {
        sheet.freezePane(columnNum, rowNum);
        return this;
    }

    public ExcelSheet build() {
        return sheet;
    }
}
