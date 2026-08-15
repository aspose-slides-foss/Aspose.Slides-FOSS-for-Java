package org.aspose.slides.foss.conformance;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Reads a produced file back with Apache POI — a reader written by other people, which knows
 * nothing about this library's object model.
 *
 * <p>A deck POI cannot open is a deck no server-side pipeline can open, so POI failing to load
 * a file is a product defect and not a test problem. POI is a test-scope dependency only.</p>
 */
public final class ThirdPartyReadBack {

    private ThirdPartyReadBack() {
    }

    /**
     * Opens a produced file with POI and asserts it loads at all.
     *
     * @param pptx the file
     * @throws IOException if POI refuses the package
     */
    public static void assertOpens(Path pptx) throws IOException {
        try (XMLSlideShow ignored = read(pptx)) {
            // loading is the assertion
        }
    }

    /**
     * @param pptx the file
     * @return the number of slides POI reports, which it takes from {@code p:sldIdLst}
     * @throws IOException if POI refuses the package
     */
    public static int slideCount(Path pptx) throws IOException {
        try (XMLSlideShow show = read(pptx)) {
            return show.getSlides().size();
        }
    }

    /**
     * @param pptx       the file
     * @param slideIndex zero-based slide index
     * @return the number of shapes POI reports on that slide
     * @throws IOException if POI refuses the package
     */
    public static int shapeCount(Path pptx, int slideIndex) throws IOException {
        try (XMLSlideShow show = read(pptx)) {
            return show.getSlides().get(slideIndex).getShapes().size();
        }
    }

    /**
     * @param pptx       the file
     * @param slideIndex zero-based slide index
     * @return the text of every text-bearing shape on that slide, in shape order
     * @throws IOException if POI refuses the package
     */
    public static List<String> textOf(Path pptx, int slideIndex) throws IOException {
        try (XMLSlideShow show = read(pptx)) {
            List<String> texts = new ArrayList<>();
            XSLFSlide slide = show.getSlides().get(slideIndex);
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFTextShape textShape) {
                    texts.add(textShape.getText());
                }
            }
            return texts;
        }
    }

    /**
     * Asserts what an independent reader sees in a produced file.
     *
     * @param pptx           the file
     * @param slides         the expected slide count
     * @param shapesOnFirst  the expected shape count on the first slide
     * @param expectedText   text that must appear somewhere on the first slide
     * @throws IOException if POI refuses the package
     */
    public static void assertReadsAs(Path pptx, int slides, int shapesOnFirst, String expectedText)
            throws IOException {
        assertThat(slideCount(pptx)).as("slides POI reads in %s", pptx).isEqualTo(slides);
        assertThat(shapeCount(pptx, 0)).as("shapes POI reads on slide 1 of %s", pptx)
                .isEqualTo(shapesOnFirst);
        assertThat(String.join("\n", textOf(pptx, 0)))
                .as("text POI reads on slide 1 of %s", pptx)
                .contains(expectedText);
    }

    /**
     * Asserts that an independent reader finds a picture on a slide and can get at its bytes.
     *
     * <p>Loading the file is not enough: a picture whose relationship does not resolve still
     * loads, and only asking for the image data shows that there is nothing behind it.</p>
     *
     * @param pptx       the file
     * @param slideIndex zero-based slide index
     * @throws IOException if POI refuses the package
     */
    public static void assertPicturesResolve(Path pptx, int slideIndex) throws IOException {
        try (XMLSlideShow show = read(pptx)) {
            List<XSLFPictureShape> pictures = new ArrayList<>();
            for (XSLFShape shape : show.getSlides().get(slideIndex).getShapes()) {
                if (shape instanceof XSLFPictureShape picture) {
                    pictures.add(picture);
                }
            }
            assertThat(pictures)
                    .as("pictures an independent reader finds on slide %d of %s",
                            slideIndex + 1, pptx)
                    .isNotEmpty();
            for (XSLFPictureShape picture : pictures) {
                assertThat(picture.getPictureData())
                        .as("image data behind picture '%s' in %s", picture.getShapeName(), pptx)
                        .isNotNull();
                assertThat(picture.getPictureData().getData())
                        .as("image bytes behind picture '%s' in %s", picture.getShapeName(), pptx)
                        .isNotEmpty();
            }
        }
    }

    private static XMLSlideShow read(Path pptx) throws IOException {
        try (InputStream in = Files.newInputStream(pptx)) {
            return new XMLSlideShow(in);
        }
    }
}
