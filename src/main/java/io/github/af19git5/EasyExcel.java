package io.github.af19git5;

import io.github.af19git5.builder.ExcelStreamWriteBuilder;
import io.github.af19git5.builder.ExcelWriteBuilder;
import io.github.af19git5.entity.ExcelSheet;
import io.github.af19git5.exception.ExcelException;
import io.github.af19git5.service.ReadExcelService;

import java.io.File;
import java.io.InputStream;
import java.util.List;

/**
 * EasyExcel可用功能及方法
 *
 * @author Jimmy Kang
 */
public class EasyExcel {

    /**
     * 讀取excel資料
     *
     * @param excelFilePath excel檔案路徑
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public static List<ExcelSheet> read(String excelFilePath) throws ExcelException {
        return new ReadExcelService().read(excelFilePath);
    }

    /**
     * 讀取excel資料
     *
     * @param excelFilePath excel檔案路徑
     * @param password 密碼
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public static List<ExcelSheet> read(String excelFilePath, String password)
            throws ExcelException {
        return new ReadExcelService().read(excelFilePath, password);
    }

    /**
     * 讀取excel資料
     *
     * @param excelFile excel檔案
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public static List<ExcelSheet> read(File excelFile) throws ExcelException {
        return new ReadExcelService().read(excelFile);
    }

    /**
     * 讀取excel資料
     *
     * @param excelFile excel檔案
     * @param password 密碼
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public static List<ExcelSheet> read(File excelFile, String password) throws ExcelException {
        return new ReadExcelService().read(excelFile, password);
    }

    /**
     * 讀取excel資料
     *
     * @param inputStream InputStream
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public static List<ExcelSheet> read(InputStream inputStream) throws ExcelException {
        return new ReadExcelService().read(inputStream);
    }

    /**
     * 讀取excel資料
     *
     * @param inputStream InputStream
     * @param password 密碼
     * @return excel資料
     * @throws ExcelException Excel處理錯誤
     */
    public static List<ExcelSheet> read(InputStream inputStream, String password)
            throws ExcelException {
        return new ReadExcelService().read(inputStream, password);
    }

    /**
     * 寫出excel資料
     *
     * @return excel寫出檢購器
     */
    public static ExcelWriteBuilder write() {
        return new ExcelWriteBuilder();
    }

    /**
     * 寫出excel資料(資料流輸出, 可以用在大檔匯出)
     *
     * @return excel寫出檢購器
     */
    public static ExcelStreamWriteBuilder writeStream() {
        return new ExcelStreamWriteBuilder();
    }
}
