package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelCell;
import io.github.af19git5.entity.ExcelStyle;

import lombok.NonNull;

import org.apache.poi.ss.usermodel.CellType;

/**
 * Excel欄位資料建構器
 *
 * @author Jimmy Kang
 */
public class ExcelCellBuilder {

    private final ExcelCell cell;

    public ExcelCellBuilder(@NonNull Integer row, @NonNull Integer column, String value) {
        cell = new ExcelCell(row, column, value);
    }

    public ExcelCellBuilder cellType(@NonNull CellType cellType) {
        cell.setCellType(cellType);
        return this;
    }

    /**
     * 設定數值(自動設為NUMERIC)
     *
     * @param value 數值，為null時寫出空白欄位
     * @return 原方法
     */
    public ExcelCellBuilder numberValue(Number value) {
        cell.setValue(ExcelCell.toNumberText(value));
        cell.setCellType(CellType.NUMERIC);
        return this;
    }

    public ExcelCellBuilder style(@NonNull ExcelStyle style) {
        cell.setStyle(style);
        return this;
    }

    public ExcelCell build() {
        return cell;
    }
}
