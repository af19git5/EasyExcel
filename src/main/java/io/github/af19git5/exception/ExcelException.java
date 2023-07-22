package io.github.af19git5.exception;

/**
 * Excel處理錯誤
 *
 * @author Jimmy Kang
 */
public class ExcelException extends Exception {

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(Exception e) {
        super(e);
    }
}
