# Easy Excel

## 關於

基於apache-poi框架，可以簡易讀取及寫出xls及xlsx檔案。

## 如何使用在自己專案中

將以下代碼加入至專案中的pom.xml dependencies.

```xml
<dependency>
  <groupId>io.github.af19git5</groupId>
  <artifactId>easy-excel</artifactId>
  <version>1.0.10</version>
</dependency>
```

## 如何使用

### 基礎類別

* **EasyExcel** -> Easy Excel功能入口。
* **ExcelSheet** -> Excel工作表資訊，紀錄工作表名稱及所有欄位資料。
* **ExcelCell** -> 欄位資料，紀錄行列數值，及對應欄位資訊。
* **ExcelStyle** -> 欄位樣式資料。
* **ExcelStreamCell** -> 欄位資料，資料流輸出處理用。
* **ExcelStreamStyle** -> 欄位樣式資料，資料流輸出處理用。

### 讀取範例

```java
File file = new File("Your excel file path.");
List<ExcelSheet> sheetList = EasyExcel.read(file);
```

### 輸出範例

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

// 輸出xls
excelWriteBuilder.outputXls("Your output path.");

// 輸出xlsx
excelWriteBuilder.outputXlsx("Your output path.");
```

### 資料流輸出範例

範例中呼叫`flush`方法為儲存該批次下有加入的欄位資料，處理大量資料時可以分批次進行flush避免oom問題。
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
