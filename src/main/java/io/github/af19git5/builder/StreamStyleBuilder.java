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
public class StreamStyleBuilder {

    private final ExcelStreamStyle style;

    public StreamStyleBuilder() {
        style = new ExcelStreamStyle();
    }

    public StreamStyleBuilder wrapText() {
        style.setIsWrapText(true);
        return this;
    }

    public StreamStyleBuilder lock() {
        style.setIsLock(true);
        return this;
    }

    public StreamStyleBuilder border(@NonNull BorderStyle borderStyle) {
        style.setAllBorder(borderStyle);
        return this;
    }

    public StreamStyleBuilder border(@NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setAllBorder(borderStyle);
        style.setAllBorderColor(color);
        return this;
    }

    public StreamStyleBuilder borderTop(@NonNull BorderStyle borderStyle) {
        style.setBorderTop(borderStyle);
        return this;
    }

    public StreamStyleBuilder borderTop(@NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setBorderTop(borderStyle);
        style.setBorderTopColor(color);
        return this;
    }

    public StreamStyleBuilder borderBottom(@NonNull BorderStyle borderStyle) {
        style.setBorderBottom(borderStyle);
        return this;
    }

    public StreamStyleBuilder borderBottom(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setBorderBottom(borderStyle);
        style.setBorderBottomColor(color);
        return this;
    }

    public StreamStyleBuilder borderLeft(@NonNull BorderStyle borderStyle) {
        style.setBorderLeft(borderStyle);
        return this;
    }

    public StreamStyleBuilder borderLeft(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setBorderLeft(borderStyle);
        style.setBorderLeftColor(color);
        return this;
    }

    public StreamStyleBuilder borderRight(@NonNull BorderStyle borderStyle) {
        style.setBorderRight(borderStyle);
        return this;
    }

    public StreamStyleBuilder borderRight(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors color) {
        style.setBorderRight(borderStyle);
        style.setBorderRightColor(color);
        return this;
    }

    public StreamStyleBuilder horizontalAlignment(@NonNull HorizontalAlignment alignment) {
        style.setHorizontalAlignment(alignment);
        return this;
    }

    public StreamStyleBuilder verticalAlignment(@NonNull VerticalAlignment alignment) {
        style.setVerticalAlignment(alignment);
        return this;
    }

    public StreamStyleBuilder backgroundColor(@NonNull IndexedColors color) {
        style.setBackgroundColor(color);
        return this;
    }

    public StreamStyleBuilder fontName(@NonNull String fontName) {
        style.setFontName(fontName);
        return this;
    }

    public StreamStyleBuilder fontSize(int fontSize) {
        style.setFontSize(fontSize);
        return this;
    }

    public StreamStyleBuilder fontColor(@NonNull IndexedColors color) {
        style.setFontColor(color);
        return this;
    }

    public StreamStyleBuilder bold() {
        style.setBold(true);
        return this;
    }

    public StreamStyleBuilder italic() {
        style.setItalic(true);
        return this;
    }

    public StreamStyleBuilder strikeout() {
        style.setStrikeout(true);
        return this;
    }

    public ExcelStreamStyle build() {
        return style;
    }
}
