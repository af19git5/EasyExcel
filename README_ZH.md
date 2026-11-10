# Easy Excel

## 關於

基於apache-poi框架，可以簡易讀取及寫出xls及xlsx檔案。

## 如何使用在自己專案中

需要Java 17以上。將以下代碼加入至專案中的pom.xml dependencies.

```xml
<dependency>
  <groupId>io.github.af19git5</groupId>
  <artifactId>easy-excel</artifactId>
  <version>2.1.0</version>
</dependency>
```

## 如何使用

### 基礎類別

* **EasyExcel** -> Easy Excel功能入口。
* **ExcelWriteBuilder** -> xls、xlsx寫出建構器(xlsx以資料流方式輸出)。
* **ExcelSheet** -> Excel工作表資訊，紀錄工作表名稱及所有欄位資料。
* **ExcelCell** -> 欄位資料，紀錄行列數值，及對應欄位資訊。
* **ExcelStyle** -> 欄位樣式資料。
* **ExcelMergedRegion** -> 欄位合併規則。

### 讀取範例

```java
File file = new File("Your excel file path.");
List<ExcelSheet> sheetList = EasyExcel.read(file);
```

### 輸出範例

`EasyExcel.write()`輸出xlsx，`EasyExcel.write(ExcelType.XLS)`輸出xls。
輸出時會自動flush尚未寫出的資料，不呼叫`flush`也能正確輸出。
處理大量資料時可以分批次呼叫`flush`將資料列寫出，避免oom問題(僅xlsx)。

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
            .createSheet("detail", "經費支用明細")
            .createSheet("grant", "核撥明細")
            .rowHeight("detail", 0, 81.6f)
            .freezePane("detail", 7, 1)
            .zoom("detail", 55)
            .cells(
                    "detail",
                    new ExcelCell(0, 0, "標題", titleStyle),
                    ExcelCell.number(1, 0, new BigDecimal("1234567"), amountStyle),
                    // 公式開頭可帶「=」，跨工作表參照的名稱請用單引號包住
                    new ExcelCell(2, 0, "='核撥明細'!A1+A2", CellType.FORMULA, amountStyle))
            .cells("grant", ExcelCell.number(0, 0, 100))
            .output("Your output path.");
}
```

`ExcelSheet`(例如`EasyExcel.read`讀出的資料)可以直接寫出：

```java
try (ExcelWriteBuilder writeBuilder = EasyExcel.write(ExcelType.XLS)) {
    for (ExcelSheet sheet : EasyExcel.read(file)) {
        writeBuilder.sheet(sheet); // 工作表代碼為工作表名稱
    }
    byte[] bytes = writeBuilder.outputBytes();
}
```

### 注意事項

* `NUMERIC`欄位接受數字或`yyyy-MM-dd HH:mm:ss`格式的日期時間，其他值會丟出`ExcelException`並附上欄位位置。空值寫成空白欄位。
* 公式會檢查語法，語法錯誤時丟出`ExcelException`。
* 寫過公式時會自動設定開檔時重新計算公式，可用`forceFormulaRecalculation(false)`關閉。
* 顏色可使用`#RGB`、`#RRGGBB`、`#AARRGGBB`(忽略透明度)或`IndexedColors`。不支援主題色加tint，請先換算成RGB。
* xlsx已flush的列無法再修改，列高、隱藏列請在該列flush前設定。
* xlsx的`autoSizeColumns`請在第一次`flush`前設定，`overrideColumnWidth`設定的寬度優先。

### xls / xlsx 支援度

| 功能 | xlsx | xls |
|---|---|---|
| 列數 / 欄數上限 | 1,048,576 / 16,384 | 65,536 / 256 |
| 記憶體 | 資料流輸出，`flush`會將資料列寫出 | 整份活頁簿保留在記憶體中，`flush`不會釋放記憶體 |
| 欄位值、樣式、合併欄位、隱藏行列、保護工作表 | ✅ | ✅ |
| 數值格式(`dataFormat`) | ✅ | ✅ |
| 列高 / 預設列高 | ✅ | ✅ |
| 凍結窗格 / 顯示比例 | ✅ | ✅ |
| 自動欄寬 | ✅ (從第一次flush開始計算) | ✅ |
| RGB顏色 | ✅ 精確 | ✅ 每份活頁簿最多56色精確(xls調色盤上限)，超過時使用最接近的顏色 |
| 公式 | ✅ Excel 2010以後新增的函式需加`_xlfn.`前綴(例如`_xlfn.TEXTJOIN`) | ✅ 僅支援Excel 97-2003的函式 |
| 跨工作表公式 | ✅ 開檔時參照的工作表需存在 | ✅ 參照的工作表需在輸出前建立，否則丟出`ExcelException` |
| 重新計算 | 開檔時由Excel重新計算(`forceFormulaRecalculation`) | 寫出時先計算結果，不支援的函式由Excel開檔時計算 |

## 1.x 升級至 2.1.0

* 需要Java 17以上。
* `ExcelStreamWriteBuilder`、`ExcelStreamCell`、`ExcelStreamStyle`、`ExcelStreamMergedRegion`改為`ExcelWriteBuilder`、`ExcelCell`、`ExcelStyle`、`ExcelMergedRegion`。樣式建構器仍可使用`IndexedColors`。
* `EasyExcel.writeStream()`改為`EasyExcel.write()`，`outputXlsx(...)`改為`output(...)`。
* 原本非資料流的`EasyExcel.write().addSheet(sheet).outputXls/outputXlsx(...)`改為`EasyExcel.write(ExcelType.XLS或XLSX).sheet(sheet).output(...)`，輸出格式在建立建構器時決定。
* 不再預設自動調整所有欄寬，需要時請呼叫`autoSizeColumns`。
* 數值格式錯誤或公式錯誤改為丟出`ExcelException`，不再寫成`0`或錯誤的公式。

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
