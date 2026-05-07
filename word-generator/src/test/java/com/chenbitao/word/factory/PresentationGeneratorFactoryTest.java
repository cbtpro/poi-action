package com.chenbitao.word.factory;

import com.chenbitao.word.presentation.HslfPresentationGenerator;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

public class PresentationGeneratorFactoryTest {

    @Test
    public void getReturnsPptGeneratorByTypeIgnoringCase() {
        assertTrue(PresentationGeneratorFactory.get("PPT") instanceof HslfPresentationGenerator);
    }

    @Test(expected = IllegalArgumentException.class)
    public void getThrowsExceptionForUnsupportedType() {
        PresentationGeneratorFactory.get("pptx");
    }
}
