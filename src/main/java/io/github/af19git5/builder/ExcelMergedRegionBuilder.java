package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelMergedRegion;

import lombok.NonNull;

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * Excel合併欄位規則建構器
 *
 * @author Jimmy Kang
 */
public class ExcelMergedRegionBuilder {

    private final ExcelMergedRegion mergedRegion;

    public ExcelMergedRegionBuilder(int firstRow, int lastRow, int firstColumn, int lastColumn) {
        mergedRegion = new ExcelMergedRegion(firstRow, lastRow, firstColumn, lastColumn);
    }

    public ExcelMergedRegionBuilder border(@NonNull BorderStyle borderStyle) {
        mergedRegion.setAllBorder(borderStyle);
        return this;
    }

    public ExcelMergedRegionBuilder border(
            @NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        mergedRegion.setAllBorder(borderStyle);
        mergedRegion.setAllBorderColor(colorHex);
        return this;
    }

    public ExcelMergedRegionBuilder borderTop(@NonNull BorderStyle borderStyle) {
        mergedRegion.setBorderTop(borderStyle);
        return this;
    }

    public ExcelMergedRegionBuilder borderTop(
            @NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        mergedRegion.setBorderTop(borderStyle);
        mergedRegion.setBorderTopColor(colorHex);
        return this;
    }

    public ExcelMergedRegionBuilder borderBottom(@NonNull BorderStyle borderStyle) {
        mergedRegion.setBorderBottom(borderStyle);
        return this;
    }

    public ExcelMergedRegionBuilder borderBottom(
            @NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        mergedRegion.setBorderBottom(borderStyle);
        mergedRegion.setBorderBottomColor(colorHex);
        return this;
    }

    public ExcelMergedRegionBuilder borderLeft(@NonNull BorderStyle borderStyle) {
        mergedRegion.setBorderLeft(borderStyle);
        return this;
    }

    public ExcelMergedRegionBuilder borderLeft(
            @NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        mergedRegion.setBorderLeft(borderStyle);
        mergedRegion.setBorderLeftColor(colorHex);
        return this;
    }

    public ExcelMergedRegionBuilder borderRight(@NonNull BorderStyle borderStyle) {
        mergedRegion.setBorderRight(borderStyle);
        return this;
    }

    public ExcelMergedRegionBuilder borderRight(
            @NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        mergedRegion.setBorderRight(borderStyle);
        mergedRegion.setBorderRightColor(colorHex);
        return this;
    }

    public ExcelMergedRegion build() {
        return mergedRegion;
    }
}
