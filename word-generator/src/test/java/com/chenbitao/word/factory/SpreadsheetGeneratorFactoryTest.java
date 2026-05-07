package com.chenbitao.word.factory;

import com.chenbitao.word.excel.XlsWorkbookGenerator;
import com.chenbitao.word.excel.XlsxWorkbookGenerator;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

public class SpreadsheetGeneratorFactoryTest {

    @Test
    public void getReturnsXlsGeneratorByTypeIgnoringCase() {
        assertTrue(SpreadsheetGeneratorFactory.get("XLS") instanceof XlsWorkbookGenerator);
    }

    @Test
    public void getReturnsXlsxGeneratorByTypeIgnoringCase() {
        assertTrue(SpreadsheetGeneratorFactory.get("XLSX") instanceof XlsxWorkbookGenerator);
    }

    @Test(expected = IllegalArgumentException.class)
    public void getThrowsExceptionForUnsupportedType() {
        SpreadsheetGeneratorFactory.get("csv");
    }
}
