package io.github.af19git5.builder;

import io.github.af19git5.entity.StreamMergedRegion;

import lombok.NonNull;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Excel合併欄位規則建構器(資料流輸出)
 *
 * @author Jimmy Kang
 */
public class StreamMergedRegionBuilder {

    private final StreamMergedRegion mergedRegion;

    public StreamMergedRegionBuilder(
            int firstRow, int lastRow, int firstColumn, int lastColumn) {
        mergedRegion = new StreamMergedRegion(firstRow, lastRow, firstColumn, lastColumn);
    }

    public StreamMergedRegionBuilder border(@NonNull BorderStyle borderStyle) {
        mergedRegion.setAllBorder(borderStyle);
        return this;
    }

    public StreamMergedRegionBuilder border(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        mergedRegion.setAllBorder(borderStyle);
        mergedRegion.setAllBorderColor(indexedColor);
        return this;
    }

    public StreamMergedRegionBuilder borderTop(@NonNull BorderStyle borderStyle) {
        mergedRegion.setBorderTop(borderStyle);
        return this;
    }

    public StreamMergedRegionBuilder borderTop(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        mergedRegion.setBorderTop(borderStyle);
        mergedRegion.setBorderTopColor(indexedColor);
        return this;
    }

    public StreamMergedRegionBuilder borderBottom(@NonNull BorderStyle borderStyle) {
        mergedRegion.setBorderBottom(borderStyle);
        return this;
    }

    public StreamMergedRegionBuilder borderBottom(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        mergedRegion.setBorderBottom(borderStyle);
        mergedRegion.setBorderBottomColor(indexedColor);
        return this;
    }

    public StreamMergedRegionBuilder borderLeft(@NonNull BorderStyle borderStyle) {
        mergedRegion.setBorderLeft(borderStyle);
        return this;
    }

    public StreamMergedRegionBuilder borderLeft(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        mergedRegion.setBorderLeft(borderStyle);
        mergedRegion.setBorderLeftColor(indexedColor);
        return this;
    }

    public StreamMergedRegionBuilder borderRight(@NonNull BorderStyle borderStyle) {
        mergedRegion.setBorderRight(borderStyle);
        return this;
    }

    public StreamMergedRegionBuilder borderRight(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        mergedRegion.setBorderRight(borderStyle);
        mergedRegion.setBorderRightColor(indexedColor);
        return this;
    }

    public StreamMergedRegion build() {
        return mergedRegion;
    }
}
