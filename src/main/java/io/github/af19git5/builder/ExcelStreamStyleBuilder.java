package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelStreamStyle;

import lombok.NonNull;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * Excel樣式建構器(資料流輸出)
 *
 * @author Jimmy Kang
 */
public class ExcelStreamStyleBuilder {

    private final ExcelStreamStyle style;

    public ExcelStreamStyleBuilder() {
        style = new ExcelStreamStyle();
    }

    public ExcelStreamStyleBuilder wrapText() {
        style.setIsWrapText(true);
        return this;
    }

    public ExcelStreamStyleBuilder lock() {
        style.setIsLock(true);
        return this;
    }

    public ExcelStreamStyleBuilder border(@NonNull BorderStyle borderStyle) {
        style.setAllBorder(borderStyle);
        return this;
    }

    public ExcelStreamStyleBuilder border(@NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setAllBorder(borderStyle);
        style.setAllBorderColor(color);
        return this;
    }

    public ExcelStreamStyleBuilder borderTop(@NonNull BorderStyle borderStyle) {
        style.setBorderTop(borderStyle);
        return this;
    }

    public ExcelStreamStyleBuilder borderTop(@NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setBorderTop(borderStyle);
        style.setBorderTopColor(color);
        return this;
    }

    public ExcelStreamStyleBuilder borderBottom(@NonNull BorderStyle borderStyle) {
        style.setBorderBottom(borderStyle);
        return this;
    }

    public ExcelStreamStyleBuilder borderBottom(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setBorderBottom(borderStyle);
        style.setBorderBottomColor(color);
        return this;
    }

    public ExcelStreamStyleBuilder borderLeft(@NonNull BorderStyle borderStyle) {
        style.setBorderLeft(borderStyle);
        return this;
    }

    public ExcelStreamStyleBuilder borderLeft(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setBorderLeft(borderStyle);
        style.setBorderLeftColor(color);
        return this;
    }

    public ExcelStreamStyleBuilder borderRight(@NonNull BorderStyle borderStyle) {
        style.setBorderRight(borderStyle);
        return this;
    }

    public ExcelStreamStyleBuilder borderRight(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setBorderRight(borderStyle);
        style.setBorderRightColor(color);
        return this;
    }

    public ExcelStreamStyleBuilder horizontalAlignment(@NonNull HorizontalAlignment alignment) {
        style.setHorizontalAlignment(alignment);
        return this;
    }

    public ExcelStreamStyleBuilder verticalAlignment(@NonNull VerticalAlignment alignment) {
        style.setVerticalAlignment(alignment);
        return this;
    }

    public ExcelStreamStyleBuilder backgroundColor(@NonNull IndexedColors color) {
        style.setBackgroundColor(color);
        return this;
    }

    public ExcelStreamStyleBuilder fontName(@NonNull String fontName) {
        style.setFontName(fontName);
        return this;
    }

    public ExcelStreamStyleBuilder fontSize(int fontSize) {
        style.setFontSize(fontSize);
        return this;
    }

    public ExcelStreamStyleBuilder fontColor(@NonNull IndexedColors color) {
        style.setFontColor(color);
        return this;
    }

    public ExcelStreamStyleBuilder bold() {
        style.setBold(true);
        return this;
    }

    public ExcelStreamStyleBuilder italic() {
        style.setItalic(true);
        return this;
    }

    public ExcelStreamStyleBuilder strikeout() {
        style.setStrikeout(true);
        return this;
    }

    public ExcelStreamStyle build() {
        return style;
    }
}
