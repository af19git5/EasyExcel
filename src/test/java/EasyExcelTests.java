import io.github.af19git5.EasyExcel;
import io.github.af19git5.entity.ExcelSheet;
import io.github.af19git5.exception.ExcelException;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.List;

/**
 * 單元測試
 *
 * @author Jimmy Kang
 */
public class EasyExcelTests {

    @Test
    public void testRead() throws ExcelException {
        List<ExcelSheet> excelSheetList = EasyExcel.read("C:\\Users\\User\\Downloads\\test-output.xls");
        System.out.println(excelSheetList);
        EasyExcel.write()
                .addSheet(excelSheetList)
                .outputXls(new File("C:\\Users\\User\\Downloads\\test2-output.xls"));
        EasyExcel.write()
                .addSheet(excelSheetList)
                .outputXlsx(new File("C:\\Users\\User\\Downloads\\test2-output.xlsx"));
    }
}
