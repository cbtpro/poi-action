package com.chenbitao.word.factory;

import com.chenbitao.word.outlook.OutlookMessageReader;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

public class OutlookMessageReaderFactoryTest {

    @Test
    public void getReturnsMsgReaderByTypeIgnoringCase() {
        assertTrue(OutlookMessageReaderFactory.get("MSG") instanceof OutlookMessageReader);
    }

    @Test(expected = IllegalArgumentException.class)
    public void getThrowsExceptionForUnsupportedType() {
        OutlookMessageReaderFactory.get("eml");
    }
}
