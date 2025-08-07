package io.github.af19git5.builder;

import io.github.af19git5.entity.StreamCell;
import io.github.af19git5.entity.StreamMergedRegion;
import io.github.af19git5.exception.ExcelException;

import java.io.File;
import java.io.IOException;
import java.security.NoSuchAlgorithmException;
import java.util.List;
import java.util.Set;

public interface StreamWriteBuilder extends AutoCloseable {
    StreamWriteBuilder createSheet(String sheetCode, String name);

    StreamWriteBuilder mergedRegions(String sheetCode, StreamMergedRegion... mergedRegions);

    StreamWriteBuilder mergedRegions(String sheetCode, List<StreamMergedRegion> mergedRegions);

    StreamWriteBuilder hiddenRowNums(String sheetCode, Integer... rowNums);

    StreamWriteBuilder hiddenRowNumSet(String sheetCode, Set<Integer> hiddenRowNumSet);

    StreamWriteBuilder hiddenColumnNums(String sheetCode, Integer... columnNums);

    StreamWriteBuilder hiddenColumnNumSet(String sheetCode, Set<Integer> hiddenColumnNumSet);

    StreamWriteBuilder protectSheet(String sheetCode, String password)
            throws NoSuchAlgorithmException;

    StreamWriteBuilder overrideColumnWidth(String sheetCode, int columnNum, int width);

    StreamWriteBuilder cells(String sheetCode, StreamCell... cells);

    StreamWriteBuilder cellList(String sheetCode, List<StreamCell> cellList);

    StreamWriteBuilder flush(String sheetCode) throws ExcelException, IOException;

    void output(String filePath) throws ExcelException;

    void output(File file) throws ExcelException;

    @Override
    void close() throws ExcelException;
}
