# Easy Excel

[繁體中文文檔](README_ZH.md)

## About

Based on the apache-poi framework, Easy Excel provides easy reading and writing of xls and xlsx files.

## How to Use in Your Project

Add this in your pom.xml dependencies.

```xml
<dependency>
  <groupId>io.github.af19git5</groupId>
  <artifactId>easy-excel</artifactId>
  <version>1.0.13</version>
</dependency>
```

## How to Use

### Basic Classes

* **EasyExcel** -> Entry point for Easy Excel.
* **ExcelSheet** -> Excel worksheet information, records the worksheet name and all column data.
* **ExcelCell** -> Column data, records row and column values, and corresponding column information.
* **ExcelStyle** -> Column style data.
* **ExcelStreamCell** -> Column data, used for data stream output processing.
* **ExcelStreamStyle** -> Column style data, used for data stream output processing.

### Reading Example

```java
File file = new File("Your excel file path.");
List<ExcelSheet> sheetList = EasyExcel.read(file);
```

### Writing Example

```java
ExcelStyle style =
        ExcelStyle.init()
                .border(BorderStyle.THIN, "#FFF0AC")
                .fontColor("#000079")
                .build();

ExcelSheet sheet =
        ExcelSheet.init()
                .name("Shee1")
                .cells(
                        new ExcelCell("Test Data 1", 0, 0, style),
                        new ExcelCell("Test Data 2", 1, 0, style)
                )
                .build();

ExcelWriteBuilder excelWriteBuilder = EasyExcel.write().addSheet(sheet);

// Output as xls
excelWriteBuilder.outputXls("Your output path.");

// Output as xlsx
excelWriteBuilder.outputXlsx("Your output path.");
```

## Stream Output Example

The `flush` method is called in the example to save the batch of added column data. When dealing with large amounts of data, flushing in batches can help avoid OOM issues.

```java
try (ExcelStreamWriteBuilder writeBuilder = EasyExcel.writeStream()) {
    ExcelStreamStyle style =
            ExcelStreamStyle.init()
                    .border(BorderStyle.THIN, IndexedColors.BLACK)
                    .backgroundColor(IndexedColors.LIGHT_YELLOW)
                    .fontColor(IndexedColors.DARK_BLUE)
                    .build();

    writeBuilder
            .createSheet("sheet", "Test Sheet")
            .cells("sheet", new ExcelStreamCell("Test Data 1", 0, 0, style))
            .flush("sheet")
            .cells("sheet", new ExcelStreamCell("Test Data 2", 1, 0, style))
            .flush("sheet")
            .outputXlsx("Your output path.");
}
```

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
