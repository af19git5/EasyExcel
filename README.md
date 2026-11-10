# Easy Excel

[繁體中文文檔](README_ZH.md)

## About

Based on the apache-poi framework, Easy Excel provides easy reading and writing of xls and xlsx files.

## How to Use in Your Project

Requires Java 17 or later. Add this in your pom.xml dependencies.

```xml
<dependency>
  <groupId>io.github.af19git5</groupId>
  <artifactId>easy-excel</artifactId>
  <version>2.1.0</version>
</dependency>
```

## How to Use

### Basic Classes

* **EasyExcel** -> Entry point for Easy Excel.
* **ExcelWriteBuilder** -> Write builder for xls and xlsx (xlsx is written as a stream).
* **ExcelSheet** -> Excel worksheet information, records the worksheet name and all column data.
* **ExcelCell** -> Column data, records row and column values, and corresponding column information.
* **ExcelStyle** -> Column style data.
* **ExcelMergedRegion** -> Merged region rule.

### Reading Example

```java
File file = new File("Your excel file path.");
List<ExcelSheet> sheetList = EasyExcel.read(file);
```

### Writing Example

`EasyExcel.write()` outputs xlsx; use `EasyExcel.write(ExcelType.XLS)` for xls.
Data that has not been flushed is flushed automatically when outputting, so calling `flush` is optional.
When dealing with large amounts of data, call `flush` in batches to write rows out and avoid OOM issues (xlsx only).

```java
try (ExcelWriteBuilder writeBuilder = EasyExcel.write()) {
    ExcelStyle titleStyle =
            ExcelStyle.init()
                    .border(BorderStyle.THIN, "#000000")
                    .backgroundColor("#F2F2F2")
                    .bold()
                    .wrapText()
                    .build();
    ExcelStyle amountStyle =
            ExcelStyle.init()
                    .border(BorderStyle.THIN, IndexedColors.BLACK)
                    .dataFormat("#,##0")
                    .build();

    writeBuilder
            .createSheet("detail", "Detail")
            .createSheet("grant", "Grant")
            .rowHeight("detail", 0, 81.6f)
            .freezePane("detail", 7, 1)
            .zoom("detail", 55)
            .cells(
                    "detail",
                    new ExcelCell(0, 0, "Title", titleStyle),
                    ExcelCell.number(1, 0, new BigDecimal("1234567"), amountStyle),
                    // Leading "=" is allowed; quote sheet names for cross-sheet references
                    new ExcelCell(2, 0, "='Grant'!A1+A2", CellType.FORMULA, amountStyle))
            .cells("grant", ExcelCell.number(0, 0, 100))
            .output("Your output path.");
}
```

An `ExcelSheet` (for example, one returned by `EasyExcel.read`) can be written directly:

```java
try (ExcelWriteBuilder writeBuilder = EasyExcel.write(ExcelType.XLS)) {
    for (ExcelSheet sheet : EasyExcel.read(file)) {
        writeBuilder.sheet(sheet); // sheet code is the sheet name
    }
    byte[] bytes = writeBuilder.outputBytes();
}
```

### Notes

* `NUMERIC` cells accept numbers or `yyyy-MM-dd HH:mm:ss` date times. Other values throw `ExcelException` with the cell position. Empty values are written as blank cells.
* Formulas are syntax-checked. Invalid formulas throw `ExcelException`.
* Formula recalculation on open is enabled automatically when any formula is written (`forceFormulaRecalculation(false)` disables it).
* Colors accept `#RGB`, `#RRGGBB`, `#AARRGGBB` (alpha is ignored) or `IndexedColors`. Theme colors with tint are not supported; convert them to RGB first.
* Rows that have been flushed (xlsx) cannot be changed. Set row heights / hidden rows before flushing those rows.
* `autoSizeColumns` must be called before the first `flush` (xlsx). Column widths set by `overrideColumnWidth` take precedence.

### xls / xlsx Support

| Feature | xlsx | xls |
|---|---|---|
| Max rows / columns | 1,048,576 / 16,384 | 65,536 / 256 |
| Memory | Streamed; `flush` writes rows out | Whole workbook kept in memory; `flush` does not free memory |
| Cell value / style / merged region / hidden rows & columns / protect sheet | ✅ | ✅ |
| Number format (`dataFormat`) | ✅ | ✅ |
| Row height / default row height | ✅ | ✅ |
| Freeze pane / zoom | ✅ | ✅ |
| Auto column width | ✅ (tracked from the first flush) | ✅ |
| RGB color | ✅ Exact | ✅ Exact for up to 56 colors per workbook (the xls palette size); additional colors use the nearest palette color |
| Formula | ✅ Functions added in Excel 2010+ need the `_xlfn.` prefix (e.g. `_xlfn.TEXTJOIN`) | ✅ Excel 97-2003 functions only |
| Cross-sheet formula | ✅ The referenced sheet must exist when the file is opened | ✅ The referenced sheet must be created before output, otherwise `ExcelException` |
| Recalculation | Excel recalculates on open (`forceFormulaRecalculation`) | Results are calculated when writing; unsupported functions are calculated by Excel on open |

## Migrating from 1.x to 2.1.0

* Java 17 or later is required.
* `ExcelStreamWriteBuilder` / `ExcelStreamCell` / `ExcelStreamStyle` / `ExcelStreamMergedRegion` are replaced by `ExcelWriteBuilder` / `ExcelCell` / `ExcelStyle` / `ExcelMergedRegion`. Style builders still accept `IndexedColors`.
* `EasyExcel.writeStream()` is replaced by `EasyExcel.write()`, and `outputXlsx(...)` by `output(...)`.
* The previous non-stream `EasyExcel.write().addSheet(sheet).outputXls/outputXlsx(...)` becomes `EasyExcel.write(ExcelType.XLS or XLSX).sheet(sheet).output(...)`. The output format is decided when the builder is created.
* Columns are no longer auto-sized by default; call `autoSizeColumns`.
* Invalid numeric values and invalid formulas now throw `ExcelException` instead of being written as `0` or as broken formulas.

## License

```
Copyright 2023 Jimmy Kang

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

  http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
