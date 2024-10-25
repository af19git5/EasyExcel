import io.github.af19git5.EasyExcel;
import io.github.af19git5.builder.ExcelStreamWriteBuilder;
import io.github.af19git5.builder.ExcelWriteBuilder;
import io.github.af19git5.entity.*;
import io.github.af19git5.exception.ExcelException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 單元測試
 *
 * @author Jimmy Kang
 */
public class EasyExcelTests {

    private static final String TEST_OUTPUT_PATH = "test-output/";

    /** 測試讀取 */
    @Test
    public void test() throws ExcelException, IOException, URISyntaxException {
        URL testXlsUrl = EasyExcelTests.class.getResource("test.xls");
        URL testXlsxUrl = EasyExcelTests.class.getResource("test.xlsx");
        if (null == testXlsUrl || null == testXlsxUrl) {
            throw new IOException("查無測試檔案");
        }
        File testXlsFile = new File(testXlsUrl.toURI());
        File testXlsxFile = new File(testXlsxUrl.toURI());
        List<ExcelSheet> xlsSheetList = EasyExcel.read(testXlsFile);
        System.out.println(xlsSheetList.get(0).toValueList());
        List<ExcelSheet> xlsxSheetList = EasyExcel.read(testXlsxFile);
        System.out.println(xlsxSheetList.get(0).toValueList());
    }


    /** 測試讀取 */
    @Test
    public void testRead() throws ExcelException, URISyntaxException, IOException {
        URL testXlsUrl = EasyExcelTests.class.getResource("test.xls");
        URL testXlsxUrl = EasyExcelTests.class.getResource("test.xlsx");
        if (null == testXlsUrl || null == testXlsxUrl) {
            throw new IOException("查無測試檔案");
        }
        File testXlsFile = new File(testXlsUrl.toURI());
        File testXlsxFile = new File(testXlsxUrl.toURI());
        List<ExcelSheet> xlsSheetList = EasyExcel.read(testXlsFile);
        System.out.println(xlsSheetList);
        List<ExcelSheet> xlsxSheetList = EasyExcel.read(testXlsxFile);
        System.out.println(xlsxSheetList);
    }

    /** 測試寫出 */
    @Test
    public void testWrite() throws ExcelException {
        ExcelStyle style =
                ExcelStyle.init()
                        .border(BorderStyle.THIN, "#D3526F")
                        .backgroundColor("#FFF0AC")
                        .fontColor("#000079")
                        .build();

        ExcelSheet sheet =
                ExcelSheet.init()
                        .name("工作表1")
                        .cells(
                                new ExcelCell("測試資料1", 0, 0, style),
                                new ExcelCell("測試資料2", 2, 0, style),
                                new ExcelCell("測試資料3", 3, 0, style))
                        .mergedRegions(
                                ExcelMergedRegion.init(0, 1, 0, 0).border(BorderStyle.THIN).build())
                        .freezePane(1, 0)
                        .build();
        ExcelWriteBuilder excelWriteBuilder = EasyExcel.write().addSheet(sheet);

        // 測試輸出xls
        excelWriteBuilder.outputXls(TEST_OUTPUT_PATH + "test.xls");

        // 測試輸出xlsx
        excelWriteBuilder.outputXlsx(TEST_OUTPUT_PATH + "test.xlsx");
    }

    /** 測試寫出(大量資料) */
    @Test
    public void testWriteLargeData() throws ExcelException {
        Date startTime = new Date();
        ExcelStyle style =
                ExcelStyle.init()
                        .border(BorderStyle.THIN, "#D3526F")
                        .backgroundColor("#FFF0AC")
                        .fontColor("#000079")
                        .build();
        ExcelSheet sheet =
                ExcelSheet.init()
                        .name("工作表1")
                        .cellList(new ArrayList<>())
                        .mergedRegions(
                                ExcelMergedRegion.init(0, 1, 0, 0)
                                        .border(BorderStyle.THIN, "#000000")
                                        .build())
                        .build();
        for (int rowNum = 0; rowNum < 100; rowNum++) {
            for (int colNum = 0; colNum < 20; colNum++) {
                sheet.getCellList()
                        .add(ExcelCell.init(rowNum, colNum, "test").style(style).build());
            }
        }
        ExcelWriteBuilder excelWriteBuilder = EasyExcel.write().addSheet(sheet);

        // 測試輸出xls
        excelWriteBuilder.outputXls(TEST_OUTPUT_PATH + "test-large.xls");

        // 測試輸出xlsx
        excelWriteBuilder.outputXlsx(TEST_OUTPUT_PATH + "test-large.xlsx");

        Date endTime = new Date();
        System.out.println("測試寫出(大量資料)，耗時" + (endTime.getTime() - startTime.getTime()) + "ms");
    }

    /** 測試資料流寫出 */
    @Test
    public void testStreamWrite() throws ExcelException {
        try (ExcelStreamWriteBuilder writeBuilder = EasyExcel.writeStream()) {
            ExcelStreamStyle style =
                    ExcelStreamStyle.init()
                            .border(BorderStyle.THIN, IndexedColors.BLACK)
                            .backgroundColor(IndexedColors.LIGHT_YELLOW)
                            .fontColor(IndexedColors.DARK_BLUE)
                            .build();
            writeBuilder
                    .createSheet("sheet", "測試工作表")
                    .cells(
                            "sheet",
                            new ExcelStreamCell("測試資料1", 0, 0, style),
                            new ExcelStreamCell("測試資料2", 2, 0, style),
                            new ExcelStreamCell("測試資料3", 3, 0, style))
                    .mergedRegions(
                            "sheet",
                            ExcelStreamMergedRegion.init(0, 1, 0, 0)
                                    .border(BorderStyle.THIN)
                                    .build())
                    .overrideColumnWidth("sheet", 0, 50 * 256)
                    .flush("sheet")
                    .outputXlsx(TEST_OUTPUT_PATH + "test-stream.xlsx");
        }
    }

    /** 測試資料流寫出(大量資料) */
    @Test
    public void testStreamWriteLargeData() throws ExcelException {
        Date startTime = new Date();
        try (ExcelStreamWriteBuilder writeBuilder = EasyExcel.writeStream()) {
            ExcelStreamStyle style =
                    ExcelStreamStyle.init()
                            .border(BorderStyle.THIN, IndexedColors.BLACK)
                            .backgroundColor(IndexedColors.LIGHT_YELLOW)
                            .fontColor(IndexedColors.DARK_BLUE)
                            .build();
            writeBuilder.createSheet("sheet", "測試工作表");
            for (int rowNum = 0; rowNum < 1000; rowNum++) {
                for (int colNum = 0; colNum < 20; colNum++) {
                    writeBuilder.cells(
                            "sheet",
                            ExcelStreamCell.init(rowNum, colNum, "test").style(style).build());
                }
            }
            writeBuilder.outputXlsx(TEST_OUTPUT_PATH + "test-stream-large.xlsx");
        }
        Date endTime = new Date();
        System.out.println("測試資料流寫出(大量資料)，耗時" + (endTime.getTime() - startTime.getTime()) + "ms");
    }
}
