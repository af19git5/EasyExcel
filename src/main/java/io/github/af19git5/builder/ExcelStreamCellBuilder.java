package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelStreamStyle;
import io.github.af19git5.entity.StreamCell;

import lombok.NonNull;

import org.apache.poi.ss.usermodel.CellType;

/**
 * Excel欄位資料建構器(資料流輸出)
 *
 * @author Jimmy Kang
 */
public class ExcelStreamCellBuilder {

    private final StreamCell cell;

    public ExcelStreamCellBuilder(@NonNull Integer row, @NonNull Integer column, String value) {
        cell = new StreamCell(row, column, value);
    }

    public ExcelStreamCellBuilder cellType(@NonNull CellType cellType) {
        cell.setCellType(cellType);
        return this;
    }

    public ExcelStreamCellBuilder style(@NonNull ExcelStreamStyle style) {
        cell.setStyle(style);
        return this;
    }

    public StreamCell build() {
        return cell;
    }
}
