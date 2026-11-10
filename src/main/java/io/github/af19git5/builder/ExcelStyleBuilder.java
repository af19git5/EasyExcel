package io.github.af19git5.builder;

import io.github.af19git5.entity.ExcelStyle;

import lombok.NonNull;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * Excel樣式建構器
 *
 * @author Jimmy Kang
 */
public class ExcelStyleBuilder {

    private final ExcelStyle style;

    public ExcelStyleBuilder() {
        style = new ExcelStyle();
    }

    public ExcelStyleBuilder wrapText() {
        style.setIsWrapText(true);
        return this;
    }

    public ExcelStyleBuilder lock() {
        style.setIsLock(true);
        return this;
    }

    public ExcelStyleBuilder border(@NonNull BorderStyle borderStyle) {
        style.setAllBorder(borderStyle);
        return this;
    }

    public ExcelStyleBuilder border(@NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        style.setAllBorder(borderStyle);
        style.setAllBorderColor(colorHex);
        return this;
    }

    public ExcelStyleBuilder border(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        style.setAllBorder(borderStyle);
        style.setAllBorderColor(indexedColor);
        return this;
    }

    public ExcelStyleBuilder borderTop(@NonNull BorderStyle borderStyle) {
        style.setBorderTop(borderStyle);
        return this;
    }

    public ExcelStyleBuilder borderTop(@NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        style.setBorderTop(borderStyle);
        style.setBorderTopColor(colorHex);
        return this;
    }

    public ExcelStyleBuilder borderTop(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        style.setBorderTop(borderStyle);
        style.setBorderTopColor(indexedColor);
        return this;
    }

    public ExcelStyleBuilder borderBottom(@NonNull BorderStyle borderStyle) {
        style.setBorderBottom(borderStyle);
        return this;
    }

    public ExcelStyleBuilder borderBottom(
            @NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        style.setBorderBottom(borderStyle);
        style.setBorderBottomColor(colorHex);
        return this;
    }

    public ExcelStyleBuilder borderBottom(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        style.setBorderBottom(borderStyle);
        style.setBorderBottomColor(indexedColor);
        return this;
    }

    public ExcelStyleBuilder borderLeft(@NonNull BorderStyle borderStyle) {
        style.setBorderLeft(borderStyle);
        return this;
    }

    public ExcelStyleBuilder borderLeft(
            @NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        style.setBorderLeft(borderStyle);
        style.setBorderLeftColor(colorHex);
        return this;
    }

    public ExcelStyleBuilder borderLeft(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        style.setBorderLeft(borderStyle);
        style.setBorderLeftColor(indexedColor);
        return this;
    }

    public ExcelStyleBuilder borderRight(@NonNull BorderStyle borderStyle) {
        style.setBorderRight(borderStyle);
        return this;
    }

    public ExcelStyleBuilder borderRight(
            @NonNull BorderStyle borderStyle, @NonNull String colorHex) {
        style.setBorderRight(borderStyle);
        style.setBorderRightColor(colorHex);
        return this;
    }

    public ExcelStyleBuilder borderRight(
            @NonNull BorderStyle borderStyle, @NonNull IndexedColors indexedColor) {
        style.setBorderRight(borderStyle);
        style.setBorderRightColor(indexedColor);
        return this;
    }

    public ExcelStyleBuilder horizontalAlignment(@NonNull HorizontalAlignment alignment) {
        style.setHorizontalAlignment(alignment);
        return this;
    }

    public ExcelStyleBuilder verticalAlignment(@NonNull VerticalAlignment alignment) {
        style.setVerticalAlignment(alignment);
        return this;
    }

    public ExcelStyleBuilder backgroundColor(@NonNull String colorHex) {
        style.setBackgroundColor(colorHex);
        return this;
    }

    public ExcelStyleBuilder backgroundColor(@NonNull IndexedColors indexedColor) {
        style.setBackgroundColor(indexedColor);
        return this;
    }

    public ExcelStyleBuilder fontName(@NonNull String fontName) {
        style.setFontName(fontName);
        return this;
    }

    public ExcelStyleBuilder fontSize(int fontSize) {
        style.setFontSize(fontSize);
        return this;
    }

    public ExcelStyleBuilder fontColor(@NonNull String colorHex) {
        style.setFontColor(colorHex);
        return this;
    }

    public ExcelStyleBuilder fontColor(@NonNull IndexedColors indexedColor) {
        style.setFontColor(indexedColor);
        return this;
    }

    public ExcelStyleBuilder bold() {
        style.setBold(true);
        return this;
    }

    public ExcelStyleBuilder italic() {
        style.setItalic(true);
        return this;
    }

    public ExcelStyleBuilder strikeout() {
        style.setStrikeout(true);
        return this;
    }

    /**
     * 數值格式
     *
     * @param dataFormat Excel數值格式(例如: #,##0、0.00%)
     * @return 原方法
     */
    public ExcelStyleBuilder dataFormat(@NonNull String dataFormat) {
        style.setDataFormat(dataFormat);
        return this;
    }

    public ExcelStyle build() {
        return style;
    }
}
