package com.chenbitao.word.factory;

import com.chenbitao.word.presentation.HslfPresentationGenerator;
import com.chenbitao.word.presentation.XslfPresentationGenerator;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

public class PresentationGeneratorFactoryTest {

    @Test
    public void getReturnsPptGeneratorByTypeIgnoringCase() {
        assertTrue(PresentationGeneratorFactory.get("PPT") instanceof HslfPresentationGenerator);
    }

    @Test
    public void getReturnsPptxGeneratorByTypeIgnoringCase() {
        assertTrue(PresentationGeneratorFactory.get("PPTX") instanceof XslfPresentationGenerator);
    }

    @Test(expected = IllegalArgumentException.class)
    public void getThrowsExceptionForUnsupportedType() {
        PresentationGeneratorFactory.get("pdf");
    }
}
