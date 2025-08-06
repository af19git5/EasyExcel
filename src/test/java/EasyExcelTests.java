import com.github.miachm.sods.*;

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

        // 測試輸出ods
        excelWriteBuilder.outputOds(TEST_OUTPUT_PATH + "test.ods");
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

    @Test
    public void testLargeOds() {
        try {
            int rows = 100000;
            int columns = 5;
            int savePoint = 5000;
            SpreadSheet spread = new SpreadSheet();
            Sheet sheet = new Sheet("A", rows, columns);
            spread.appendSheet(sheet);
            for (int i = 0; i < rows; i++) {
                for (int j = 0; j < columns; j++) {
                    sheet.getRange(i, j).setValues(j);
                }
                if (i % savePoint == 0 || i == rows - 1) {
                    spread.save(new File("Out.ods"));
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testOds() throws IOException {
        // SpreadSheet對應ExcelWriteBuilder，表示整個ods檔案
        SpreadSheet spreadSheet = new SpreadSheet();

        // Sheet對應ExcelSheet，表示單一工作表
        Sheet sheet = new Sheet("測試工作表1");
        // 把建好的工作表sheet加進ods檔案
        spreadSheet.appendSheet(sheet);

        // 標題row
        String[] titles = {"測試標題1", "測試標題2", "測試標題3", "測試標題4"};
        sheet.appendColumns(3);
        for (int i = 0; i < titles.length; i++) {
            String title = titles[i];
            // Range是指該工作表中的某一區域，對應EasyExcel要一格一格看
            Range range = sheet.getRange(0, i);
            range.setValue(title);
            // Style對應ExcelStyle
            Style style = new Style();
            style.setBold(true);
            style.setBackgroundColor(new Color(255, 255, 0));
            style.setFontColor(new Color("#FF0000"));
            Borders borders = new Borders();
            borders.setBorder(true);
            borders.setBorderProperties("0.07cm solid #0000FF");
            style.setBorders(borders);
            range.setStyle(style);
        }

        // 資料rows
        String[][] datas1 = {
            {"測試資料1-1", "測試資料1-2", "測試資料1-3", "測試資料1-4"},
            {"測試資料2-1", "測試資料2-2", "測試資料2-3", "測試資料2-4"}
        };
        sheet.appendRows(datas1.length);
        for (int i = 1; i <= datas1.length; i++) {
            String[] dataRow = datas1[i - 1];
            for (int j = 0; j < dataRow.length; j++) {
                Range range = sheet.getRange(i, j);
                range.setValue(dataRow[j]);
                Style style = new Style();
                style.setBackgroundColor(new Color(255, 192, 203));
                style.setFontColor(new Color("#292421"));
                Borders borders = new Borders(true);
                style.setBorders(borders);
                range.setStyle(style);
            }

            Range mergeRange = sheet.getRange(i, 0, 1, 4);
            mergeRange.merge();
        }

        // ods檔案輸出
        spreadSheet.save(new File(TEST_OUTPUT_PATH + "ods檔案1.ods"));
    }
}
