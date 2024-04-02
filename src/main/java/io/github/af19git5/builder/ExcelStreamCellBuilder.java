package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelStreamCell;
import io.github.af19git5.entity.ExcelStreamStyle;

import lombok.NonNull;

import org.apache.poi.ss.usermodel.CellType;

/**
 * Excel欄位資料建構器
 *
 * @author Jimmy Kang
 */
public class ExcelStreamCellBuilder {

    private final ExcelStreamCell cell;

    public ExcelStreamCellBuilder(@NonNull Integer row, @NonNull Integer column, String value) {
        cell = new ExcelStreamCell(row, column, value);
    }

    public ExcelStreamCellBuilder cellType(@NonNull CellType cellType) {
        cell.setCellType(cellType);
        return this;
    }

    public ExcelStreamCellBuilder style(@NonNull ExcelStreamStyle style) {
        cell.setStyle(style);
        return this;
    }

    public ExcelStreamCell build() {
        return cell;
    }
}
