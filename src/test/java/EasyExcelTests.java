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

    /** 測試資料流寫出 */
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
                    .addCell("sheet", new ExcelStreamCell("測試資料1", 0, 0, style))
                    .flush("sheet")
                    .addCell("sheet", new ExcelStreamCell("測試資料2", 1, 0, style))
                    .flush("sheet")
                    .outputXlsx(TEST_OUTPUT_PATH + "test-stream.xlsx");
        }
    }
}
