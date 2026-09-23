package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.ICommentAuthor;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.drawing.PointF;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.w3c.dom.Element;

import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.within;

/**
 * A comment must be placed where the caller put it, in the unit PowerPoint reads.
 *
 * <p>The schema types {@code p:cm/p:pos} as a point in EMU, but PowerPoint writes and reads it
 * in its own unit, one eighth of a point (576 to the inch): a comment PowerPoint places at 10 pt
 * from the top-left corner is written {@code <p:pos x="80" y="80"/>}. Positions were written in
 * EMU, 1587.5 times too large, which puts every comment far outside the slide.</p>
 *
 * <p>A comment's position is given in centimetres from the top-left corner of the slide, which
 * is what the library has always meant by it; only the value written to the file changes.</p>
 */
class CommentPositionConformanceTest {

    /** PowerPoint's comment-position units per centimetre: 576 per inch. */
    static final double UNITS_PER_CM = 576 / 2.54;

    /** EMU per comment-position unit: 12 700 EMU per point, eight units per point. */
    static final double EMU_PER_UNIT = 12_700 / 8.0;

    @TempDir
    Path tempDir;

    /** One inch across and two down must be written as PowerPoint writes 72 pt and 144 pt. */
    @Test
    void aCommentMustBeWrittenInTheUnitPowerPointReads() throws Exception {
        Path out = writeComment(2.54f, 5.08f);
        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            PackageAssertions.assertElementWithAttributes(pkg, commentsPartOf(pkg),
                    "//p:cm/p:pos", java.util.Map.of("x", "576", "y", "1152"));
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /** Every position on the slide must be written as a position on the slide. */
    @ParameterizedTest(name = "({0} cm, {1} cm)")
    @CsvSource({"0,0", "1,1", "2,3", "12.7,9.525", "25.4,19.05", "25,19"})
    void aCommentPlacedOnTheSlideMustLieOnTheSlide(float xCm, float yCm) throws Exception {
        Path out = writeComment(xCm, yCm);
        try (PptxPackage pkg = PptxPackage.open(out)) {
            Element sldSz = PackageAssertions.selectNodes(pkg, "ppt/presentation.xml",
                    "/p:presentation/p:sldSz").get(0);
            double width = Long.parseLong(sldSz.getAttribute("cx")) / EMU_PER_UNIT;
            double height = Long.parseLong(sldSz.getAttribute("cy")) / EMU_PER_UNIT;
            Element pos = PackageAssertions.selectNodes(pkg, commentsPartOf(pkg), "//p:cm/p:pos")
                    .get(0);
            long x = Long.parseLong(pos.getAttribute("x"));
            long y = Long.parseLong(pos.getAttribute("y"));
            assertThat(x).as("p:pos/@x for %s cm on a slide %s units wide", xCm, width)
                    .isBetween(0L, Math.round(width))
                    .isEqualTo(Math.round(xCm * UNITS_PER_CM));
            assertThat(y).as("p:pos/@y for %s cm on a slide %s units high", yCm, height)
                    .isBetween(0L, Math.round(height))
                    .isEqualTo(Math.round(yCm * UNITS_PER_CM));
        }
    }

    /**
     * A comment PowerPoint placed must read back at the position PowerPoint placed it.
     *
     * <p>A reader test: the expected value is PowerPoint's own {@code x="80" y="80"} for a
     * comment at 10 pt, which is 10 / 72 × 2.54 cm.</p>
     */
    @Test
    void aCommentPlacedByPowerPointMustReadBackWherePowerPointPutIt() throws Exception {
        Path written = writeComment(1f, 1f);
        String part;
        try (PptxPackage pkg = PptxPackage.open(written)) {
            part = commentsPartOf(pkg);
        }
        Path deck = ZipSurgery.copyWithReplacement(written, tempDir.resolve("powerpoint-pos.pptx"),
                part, xml -> xml.replaceAll("<p:pos [^>]*/>", "<p:pos x=\"80\" y=\"80\"/>"));

        try (var pres = new Presentation(deck.toString())) {
            PointF position = pres.getSlides().get(0).getSlideComments(null)[0].getPosition();
            double expectedCm = 10 / 72.0 * 2.54;
            assertThat(position.getX()).as("x of a comment PowerPoint placed at 10 pt, in cm")
                    .isCloseTo((float) expectedCm, within(0.001f));
            assertThat(position.getY()).as("y of a comment PowerPoint placed at 10 pt, in cm")
                    .isCloseTo((float) expectedCm, within(0.001f));
        }
    }

    private Path writeComment(float xCm, float yCm) throws Exception {
        Path out = tempDir.resolve("comment-" + xCm + "-" + yCm + ".pptx");
        try (var pres = new Presentation()) {
            ICommentAuthor author = pres.getCommentAuthors().addAuthor("Reviewer", "RV");
            author.getComments().addComment("Here", pres.getSlides().get(0),
                    new PointF(xCm, yCm), LocalDateTime.of(2026, 1, 2, 3, 4, 5));
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        return out;
    }

    private static String commentsPartOf(PptxPackage pkg) throws java.io.IOException {
        List<String> parts = PackageAssertions.relatedParts(pkg,
                PackageAssertions.slidePartNames(pkg).get(0),
                PartNumberingConformanceTest.REL_COMMENTS);
        assertThat(parts).as("comments parts of slide 1 in %s", pkg.path()).hasSize(1);
        return parts.get(0);
    }
}
