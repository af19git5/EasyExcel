import static org.junit.jupiter.api.Assertions.*;

import io.github.af19git5.EasyExcel;
import io.github.af19git5.builder.ExcelWriteBuilder;
import io.github.af19git5.entity.*;
import io.github.af19git5.exception.ExcelException;
import io.github.af19git5.type.ExcelType;

import lombok.extern.java.Log;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.Date;
import java.util.List;

/**
 * 單元測試
 *
 * @author Jimmy Kang
 */
@Log
public class EasyExcelTests {

    private static final String TEST_OUTPUT_PATH = "test-output/";

    private static final String ACCOUNTING_FORMAT = "_-* #,##0_-;\\-* #,##0_-;_-* \"-\"??_-;_-@_-";

    @BeforeAll
    public static void createOutputFolder() {
        new File(TEST_OUTPUT_PATH).mkdirs();
    }

    private static File getTestFile(String name) throws IOException, URISyntaxException {
        URL url = EasyExcelTests.class.getResource(name);
        if (null == url) {
            throw new IOException("查無測試檔案");
        }
        return new File(url.toURI());
    }

    private static String extension(ExcelType type) {
        return ExcelType.XLS.equals(type) ? ".xls" : ".xlsx";
    }

    /** 取得輸出檔案，檔名為案例名稱加上副檔名，例如 numeric-values.xlsx */
    private static File outputFile(String caseName, ExcelType type) {
        return new File(TEST_OUTPUT_PATH + caseName + extension(type));
    }

    /** 以唯讀模式讀回輸出檔案 */
    private static Workbook readBack(File file) throws IOException {
        return WorkbookFactory.create(file, null, true);
    }

    /** 測試讀取 */
    @Test
    public void testRead() throws ExcelException, URISyntaxException, IOException {
        List<ExcelSheet> xlsSheetList = EasyExcel.read(getTestFile("test.xls"));
        log.info("xls 讀取結果：" + xlsSheetList.get(0).toValueList());
        List<ExcelSheet> xlsxSheetList = EasyExcel.read(getTestFile("test.xlsx"));
        log.info("xlsx 讀取結果：" + xlsxSheetList.get(0).toValueList());
    }

    /** 測試讀取不會改寫原檔 */
    @Test
    public void testReadDoesNotModifyFile() throws ExcelException, URISyntaxException, IOException {
        for (String fileName : new String[] {"test.xls", "test.xlsx"}) {
            File copy = new File(TEST_OUTPUT_PATH + "read-does-not-modify-source-" + fileName);
            Files.copy(
                    getTestFile(fileName).toPath(),
                    copy.toPath(),
                    StandardCopyOption.REPLACE_EXISTING);
            byte[] before = Files.readAllBytes(copy.toPath());
            EasyExcel.read(copy);
            assertArrayEquals(before, Files.readAllBytes(copy.toPath()), fileName);
        }
    }

    /** 測試讀取後再寫出 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testReadThenWrite(ExcelType type)
            throws ExcelException, URISyntaxException, IOException {
        for (String fileName : new String[] {"test.xls", "test.xlsx"}) {
            List<ExcelSheet> sheetList = EasyExcel.read(getTestFile(fileName));
            File file = outputFile("read-then-write-from-" + fileName.replace('.', '-'), type);
            try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
                for (ExcelSheet sheet : sheetList) {
                    writeBuilder.sheet(sheet);
                }
                writeBuilder.output(file);
            }
            try (Workbook workbook = readBack(file)) {
                assertEquals(sheetList.size(), workbook.getNumberOfSheets());
                Sheet sheet = workbook.getSheetAt(0);
                assertEquals(sheetList.get(0).getName(), sheet.getSheetName());
                assertEquals(
                        sheetList.get(0).getMergedRegionList().size(), sheet.getNumMergedRegions());
            }
        }
    }

    /** 測試寫出(不呼叫flush，輸出時自動寫出) */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testWrite(ExcelType type) throws ExcelException, IOException {
        File file = outputFile("write-basic-without-flush", type);
        ExcelStyle style =
                ExcelStyle.init()
                        .border(BorderStyle.THIN, "#D3526F")
                        .backgroundColor("#FFF0AC")
                        .fontColor("#000079")
                        .build();
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "工作表1")
                    .cells(
                            "sheet",
                            new ExcelCell("測試資料1", 0, 0, style),
                            new ExcelCell("測試資料2", 2, 0, style),
                            new ExcelCell("測試資料3", 3, 0, style))
                    .mergedRegions(
                            "sheet",
                            ExcelMergedRegion.init(0, 1, 0, 0).border(BorderStyle.THIN).build())
                    .freezePane("sheet", 1, 0)
                    .autoSizeColumns("sheet", 0)
                    .output(file);
        }
        try (Workbook workbook = readBack(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("測試資料3", sheet.getRow(3).getCell(0).getStringCellValue());
            assertEquals(1, sheet.getNumMergedRegions());
        }
    }

    /** 測試數值欄位 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testNumeric(ExcelType type) throws ExcelException, IOException {
        File file = outputFile("numeric-values", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "數值")
                    .cells(
                            "sheet",
                            new ExcelCell(0, 0, "1234.5", CellType.NUMERIC),
                            ExcelCell.number(0, 1, new BigDecimal("12345678901.25")),
                            ExcelCell.number(0, 2, 42L),
                            ExcelCell.init(0, 3, null).numberValue(3.5d).build(),
                            ExcelCell.number(0, 4, null),
                            new ExcelCell(0, 5, "2026-01-02 03:04:05", CellType.NUMERIC))
                    .output(file);
        }
        try (Workbook workbook = readBack(file)) {
            Row row = workbook.getSheetAt(0).getRow(0);
            assertEquals(CellType.NUMERIC, row.getCell(0).getCellType());
            assertEquals(1234.5, row.getCell(0).getNumericCellValue());
            assertEquals(12345678901.25, row.getCell(1).getNumericCellValue());
            assertEquals(42, row.getCell(2).getNumericCellValue());
            assertEquals(3.5, row.getCell(3).getNumericCellValue());
            assertEquals(CellType.BLANK, row.getCell(4).getCellType());
            assertEquals(CellType.NUMERIC, row.getCell(5).getCellType());
        }
    }

    /** 測試數值格式錯誤 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testInvalidNumeric(ExcelType type) throws ExcelException, IOException {
        File file = outputFile("invalid-numeric-should-not-exist", type);
        Files.deleteIfExists(file.toPath());
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "數值")
                    .cells("sheet", new ExcelCell(3, 2, "abc", CellType.NUMERIC));
            ExcelException e = assertThrows(ExcelException.class, () -> writeBuilder.output(file));
            assertTrue(e.getMessage().contains("row: 3, column: 2"), e.getMessage());
        }
        // 資料錯誤時不應留下檔案
        assertFalse(file.exists());
    }

    /** 測試數值格式樣式 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testDataFormat(ExcelType type) throws ExcelException, IOException {
        File file = outputFile("data-format-accounting-percent", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "格式")
                    .cells(
                            "sheet",
                            ExcelCell.number(
                                    0,
                                    0,
                                    0,
                                    ExcelStyle.init().dataFormat(ACCOUNTING_FORMAT).build()),
                            ExcelCell.number(
                                    0, 1, 0.1234, ExcelStyle.init().dataFormat("0.000%").build()),
                            ExcelCell.number(
                                    0,
                                    2,
                                    1000,
                                    ExcelStyle.init()
                                            .dataFormat("#,##0_);[Red]\\(#,##0\\)")
                                            .build()))
                    .output(file);
        }
        try (Workbook workbook = readBack(file)) {
            Row row = workbook.getSheetAt(0).getRow(0);
            assertEquals(ACCOUNTING_FORMAT, row.getCell(0).getCellStyle().getDataFormatString());
            assertEquals("0.000%", row.getCell(1).getCellStyle().getDataFormatString());
            assertEquals(
                    "#,##0_);[Red]\\(#,##0\\)",
                    row.getCell(2).getCellStyle().getDataFormatString());
        }
    }

    /** 測試相同樣式只建立一次 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testStyleReuse(ExcelType type) throws ExcelException, IOException {
        File file = outputFile("style-reuse-100-rows", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder.createSheet("sheet", "樣式");
            for (int rowNum = 0; rowNum < 100; rowNum++) {
                writeBuilder.cells(
                        "sheet",
                        new ExcelCell(
                                rowNum,
                                0,
                                "test",
                                ExcelStyle.init().bold().backgroundColor("#F2F2F2").build()));
            }
            writeBuilder.output(file);
        }
        try (Workbook workbook = readBack(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(
                    sheet.getRow(0).getCell(0).getCellStyle().getIndex(),
                    sheet.getRow(99).getCell(0).getCellStyle().getIndex());
        }
    }

    /** 測試跨工作表公式 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testCrossSheetFormula(ExcelType type) throws ExcelException, IOException {
        File file = outputFile("cross-sheet-formula", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            // 公式寫入時參照的工作表尚未建立
            writeBuilder
                    .createSheet("detail", "經費支用明細")
                    .cells(
                            "detail",
                            ExcelCell.number(0, 0, 100),
                            new ExcelCell(1, 0, "='核撥明細'!B1+A1", CellType.FORMULA))
                    .createSheet("grant", "核撥明細")
                    .cells(
                            "grant",
                            ExcelCell.number(0, 1, 50),
                            new ExcelCell(1, 1, "SUM('經費支用明細'!A1:A2)", CellType.FORMULA));
            writeBuilder.output(file);
        }
        try (Workbook workbook = readBack(file)) {
            Cell detailFormula = workbook.getSheet("經費支用明細").getRow(1).getCell(0);
            Cell grantFormula = workbook.getSheet("核撥明細").getRow(1).getCell(1);
            assertEquals(CellType.FORMULA, detailFormula.getCellType());
            assertFalse(detailFormula.getCellFormula().contains("#REF!"));
            assertTrue(detailFormula.getCellFormula().contains("核撥明細"));
            assertTrue(grantFormula.getCellFormula().contains("經費支用明細"));
            if (ExcelType.XLSX.equals(type)) {
                assertTrue(workbook.getForceFormulaRecalculation());
            } else {
                // xls輸出時已計算公式結果
                assertEquals(150, detailFormula.getNumericCellValue());
                assertEquals(250, grantFormula.getNumericCellValue());
            }
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            assertEquals(150, evaluator.evaluate(detailFormula).getNumberValue());
            assertEquals(250, evaluator.evaluate(grantFormula).getNumberValue());
        }
    }

    /** 測試公式錯誤 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testInvalidFormula(ExcelType type) throws ExcelException, IOException {
        File syntaxErrorFile = outputFile("invalid-formula-syntax-should-not-exist", type);
        Files.deleteIfExists(syntaxErrorFile.toPath());
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "公式")
                    .cells("sheet", new ExcelCell(0, 0, "SUM(A1:", CellType.FORMULA));
            assertThrows(ExcelException.class, () -> writeBuilder.output(syntaxErrorFile));
        }
        assertFalse(syntaxErrorFile.exists());

        File missingSheetFile = outputFile("invalid-formula-missing-sheet", type);
        Files.deleteIfExists(missingSheetFile.toPath());
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "公式")
                    .cells("sheet", new ExcelCell(0, 0, "'不存在'!A1", CellType.FORMULA));
            if (ExcelType.XLS.equals(type)) {
                assertThrows(ExcelException.class, () -> writeBuilder.output(missingSheetFile));
                assertFalse(missingSheetFile.exists());
            } else {
                // xlsx不檢查工作表是否存在，由Excel開檔時處理(開檔會顯示#REF!)
                writeBuilder.output(missingSheetFile);
                assertTrue(missingSheetFile.exists());
            }
        }
    }

    /** 測試指定不強制重新計算 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testDisableForceFormulaRecalculation(ExcelType type)
            throws ExcelException, IOException {
        File file = outputFile("formula-without-force-recalculation", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "公式")
                    .cells("sheet", new ExcelCell(0, 0, "1+1", CellType.FORMULA))
                    .forceFormulaRecalculation(false)
                    .output(file);
        }
        try (Workbook workbook = readBack(file)) {
            assertFalse(workbook.getForceFormulaRecalculation());
            if (ExcelType.XLS.equals(type)) {
                // xls不支援重算旗標，輸出時已計算公式結果
                assertEquals(2, workbook.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue());
            }
        }
    }

    /** 測試列高、隱藏列、凍結窗格、顯示比例 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testSheetSetting(ExcelType type) throws ExcelException, IOException {
        File file = outputFile("sheet-setting-row-height-hidden-freeze-zoom", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "設定")
                    .defaultRowHeight("sheet", 20f)
                    .rowHeight("sheet", 0, 81.6f)
                    .rowHeight("sheet", 5, 30f)
                    .rowHeight("sheet", 300, 16.5f)
                    .hiddenRowNums("sheet", 6)
                    .freezePane("sheet", 7, 2)
                    .zoom("sheet", 55);
            writeBuilder.cells("sheet", new ExcelCell(0, 0, "表頭"));
            writeBuilder.flush("sheet");
            // 第5、6列沒有欄位資料，也要套用列高與隱藏
            writeBuilder.cells("sheet", new ExcelCell(10, 0, "資料"));
            writeBuilder.flush("sheet");
            writeBuilder.output(file);
        }
        try (Workbook workbook = readBack(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(81.6f, sheet.getRow(0).getHeightInPoints(), 0.05f);
            assertEquals(30f, sheet.getRow(5).getHeightInPoints(), 0.05f);
            assertEquals(16.5f, sheet.getRow(300).getHeightInPoints(), 0.05f);
            assertTrue(sheet.getRow(6).getZeroHeight());
            assertEquals(20f, sheet.getDefaultRowHeightInPoints(), 0.05f);
            PaneInformation pane = sheet.getPaneInformation();
            assertNotNull(pane);
            assertEquals(7, pane.getVerticalSplitPosition());
            assertEquals(2, pane.getHorizontalSplitPosition());
            if (sheet instanceof XSSFSheet) {
                assertEquals(
                        55,
                        ((XSSFSheet) sheet)
                                .getCTWorksheet()
                                .getSheetViews()
                                .getSheetViewArray(0)
                                .getZoomScale());
            }
        }
    }

    /** 測試修改已寫出的列(僅xlsx會分批寫出，xls全部保留在記憶體中可隨時修改) */
    @Test
    public void testChangeFlushedRow() throws ExcelException {
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write()) {
            writeBuilder
                    .createSheet("sheet", "設定")
                    .cells("sheet", new ExcelCell(5, 0, "資料"))
                    .flush("sheet");
            assertThrows(
                    IllegalStateException.class, () -> writeBuilder.rowHeight("sheet", 5, 30f));
            writeBuilder.cells("sheet", new ExcelCell(3, 0, "資料"));
            assertThrows(ExcelException.class, () -> writeBuilder.flush("sheet"));
        }
    }

    /** 測試RGB顏色 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testRgbColor(ExcelType type) throws ExcelException, IOException {
        File file = outputFile("rgb-color-background-font", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "顏色")
                    .cells(
                            "sheet",
                            new ExcelCell(
                                    0,
                                    0,
                                    "灰底",
                                    ExcelStyle.init().backgroundColor("#F2F2F2").build()),
                            new ExcelCell(
                                    0, 1, "黃底", ExcelStyle.init().backgroundColor("#FF0").build()),
                            new ExcelCell(
                                    0,
                                    2,
                                    "ARGB",
                                    ExcelStyle.init()
                                            .backgroundColor("FFA1B2C3")
                                            .fontColor(IndexedColors.DARK_BLUE)
                                            .build()));
            // 超過xls調色盤數量時改用近似色
            for (int i = 0; i < 80; i++) {
                writeBuilder.cells(
                        "sheet",
                        new ExcelCell(
                                1,
                                i,
                                "c" + i,
                                ExcelStyle.init()
                                        .backgroundColor(
                                                String.format("#%02X%02X%02X", i * 3, 255 - i, i))
                                        .build()));
            }
            writeBuilder.output(file);
        }
        try (Workbook workbook = readBack(file)) {
            Row row = workbook.getSheetAt(0).getRow(0);
            assertEquals("#F2F2F2", backgroundHex(workbook, row.getCell(0)));
            assertEquals("#FFFF00", backgroundHex(workbook, row.getCell(1)));
            assertEquals("#A1B2C3", backgroundHex(workbook, row.getCell(2)));
        }
    }

    private static String backgroundHex(Workbook workbook, Cell cell) {
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle instanceof XSSFCellStyle) {
            return "#"
                    + ((XSSFCellStyle) cellStyle)
                            .getFillForegroundXSSFColor()
                            .getARGBHex()
                            .substring(2);
        }
        HSSFColor color =
                ((HSSFWorkbook) workbook)
                        .getCustomPalette()
                        .getColor(cellStyle.getFillForegroundColor());
        short[] rgb = color.getTriplet();
        return String.format("#%02X%02X%02X", rgb[0], rgb[1], rgb[2]);
    }

    /** 測試合併欄位邊線 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testMergedRegionBorder(ExcelType type) throws ExcelException, IOException {
        File file = outputFile("merged-region-border", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder
                    .createSheet("sheet", "合併")
                    .cells("sheet", new ExcelCell(0, 0, "合併標題", ExcelStyle.init().bold().build()))
                    .mergedRegions(
                            "sheet",
                            ExcelMergedRegion.init(0, 1, 0, 2)
                                    .border(BorderStyle.THIN, "#FF0000")
                                    .build());
            writeBuilder.flush("sheet");
            writeBuilder.cells("sheet", new ExcelCell(1, 1, "合併內"));
            writeBuilder.output(file);
        }
        try (Workbook workbook = readBack(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(1, sheet.getNumMergedRegions());
            CellStyle topLeft = sheet.getRow(0).getCell(0).getCellStyle();
            assertEquals(BorderStyle.THIN, topLeft.getBorderTop());
            assertTrue(workbook.getFontAt(topLeft.getFontIndex()).getBold());
            assertEquals(
                    BorderStyle.THIN, sheet.getRow(1).getCell(2).getCellStyle().getBorderBottom());
        }
    }

    /** 測試重複輸出，且輸出後無法再修改 */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testOutputTwice(ExcelType type) throws ExcelException, IOException {
        File first = outputFile("output-twice-first", type);
        File second = outputFile("output-twice-second", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder.createSheet("sheet", "輸出").cells("sheet", new ExcelCell(0, 0, "a"));
            writeBuilder.output(first);
            writeBuilder.output(second);
            assertThrows(
                    IllegalStateException.class,
                    () -> writeBuilder.cells("sheet", new ExcelCell(1, 0, "b")));
        }
        for (File file : new File[] {first, second}) {
            try (Workbook workbook = readBack(file)) {
                assertEquals("a", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            }
        }
    }

    /** 測試資料流寫出(大量資料) */
    @ParameterizedTest
    @EnumSource(ExcelType.class)
    public void testWriteLargeData(ExcelType type) throws ExcelException, IOException {
        Date startTime = new Date();
        ExcelStyle style =
                ExcelStyle.init()
                        .border(BorderStyle.THIN, IndexedColors.BLACK)
                        .backgroundColor(IndexedColors.LIGHT_YELLOW)
                        .fontColor(IndexedColors.DARK_BLUE)
                        .build();
        int rowCount = 20000;
        File file = outputFile("large-data-20000-rows-30-columns", type);
        try (ExcelWriteBuilder writeBuilder = EasyExcel.write(type)) {
            writeBuilder.createSheet("sheet", "測試工作表");
            for (int rowNum = 0; rowNum < rowCount; rowNum++) {
                for (int colNum = 0; colNum < 30; colNum++) {
                    writeBuilder.cells(
                            "sheet", ExcelCell.init(rowNum, colNum, "test").style(style).build());
                }
                if (rowNum % 2000 == 1999) {
                    writeBuilder.flush("sheet");
                }
            }
            writeBuilder.output(file);
        }
        try (Workbook workbook = readBack(file)) {
            assertEquals(rowCount - 1, workbook.getSheetAt(0).getLastRowNum());
        }
        Date endTime = new Date();
        log.info(type + " 測試資料流寫出(大量資料)，耗時" + (endTime.getTime() - startTime.getTime()) + "ms");
    }
}
