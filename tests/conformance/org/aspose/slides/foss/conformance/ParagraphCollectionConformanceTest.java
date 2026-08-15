package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.IAutoShape;
import org.aspose.slides.foss.Paragraph;
import org.aspose.slides.foss.Portion;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.ShapeType;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Element;

import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A paragraph added to a text frame must be in the file that text frame is saved to.
 *
 * <p>Setting the whole text of a shape at once works; building text up out of paragraphs and
 * portions — the only way to format part of a line differently from the rest — does not reach
 * the XML at all. The collection reports the new paragraph back, so the caller has no reason to
 * suspect anything, and the file has one paragraph where the code wrote three.</p>
 */
class ParagraphCollectionConformanceTest {

    @TempDir
    Path tempDir;

    /** A paragraph added through the collection must be written to the file. */
    @Test
    void aParagraphAddedThroughTheCollectionMustReachTheFile() throws Exception {
        Path out = tempDir.resolve("paragraphs.pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 200);
            shape.addTextFrame("First paragraph");
            Paragraph second = new Paragraph();
            second.setText("Second paragraph");
            shape.getTextFrame().getParagraphs().add(second);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<Element> paragraphs = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:txBody/a:p");
            assertThat(paragraphs)
                    .as("a:p elements in slide 1 of %s%n%s",
                            out, pkg.text("ppt/slides/slide1.xml"))
                    .hasSize(2);
        }
        assertThat(String.join("\n", ThirdPartyReadBack.textOf(out, 0)))
                .as("text an independent reader finds in %s", out)
                .contains("First paragraph")
                .contains("Second paragraph");
    }

    /**
     * A text frame emptied through the collection must still be a valid text body.
     *
     * <p>{@code CT_TextBody} is {@code bodyPr, lstStyle?, p+}: the paragraph is required, so a
     * body written with none is invalid however it came to be empty. PowerPoint keeps one
     * empty paragraph in a cleared placeholder for the same reason.</p>
     */
    @Test
    void aTextFrameEmptiedThroughTheCollectionMustStillCarryAParagraph() throws Exception {
        Path out = tempDir.resolve("cleared.pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 200);
            shape.addTextFrame("Text that is then removed");
            shape.getTextFrame().getParagraphs().clear();
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        assertTextBodiesCarryAParagraph(out);
    }

    /** Removing the last paragraph one at a time must leave a valid text body too. */
    @Test
    void removingTheLastParagraphMustLeaveTheTextBodyValid() throws Exception {
        Path out = tempDir.resolve("removed.pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 200);
            shape.addTextFrame("Only paragraph");
            var paragraphs = shape.getTextFrame().getParagraphs();
            while (paragraphs.size() > 0) {
                paragraphs.removeAt(0);
            }
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        assertTextBodiesCarryAParagraph(out);
    }

    /** Setting the text to {@code null} must empty the frame without invalidating it. */
    @Test
    void settingTheTextToNullMustLeaveTheTextBodyValid() throws Exception {
        Path out = tempDir.resolve("nulled.pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 200);
            shape.addTextFrame("Text that is then removed");
            shape.getTextFrame().setText(null);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        assertTextBodiesCarryAParagraph(out);
    }

    /**
     * Replacing the text must not leave the paragraph that emptying the frame put there.
     *
     * <p>Calibration for the three tests above: an implementation that keeps a placeholder
     * paragraph unconditionally passes them and silently prefixes every line of text with a
     * blank one.</p>
     */
    @Test
    void replacingTheTextMustNotLeaveABlankParagraphBehind() throws Exception {
        Path out = tempDir.resolve("replaced.pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 200);
            shape.addTextFrame("First text");
            shape.getTextFrame().setText("Replacement line one\nReplacement line two");
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            assertThat(PackageAssertions.selectNodes(pkg, "ppt/slides/slide1.xml",
                    "//p:sp/p:txBody/a:p"))
                    .as("a:p elements in the shape's text body of %s%n%s",
                            out, pkg.text("ppt/slides/slide1.xml"))
                    .hasSize(2);
        }
    }

    /** Asserts that no text body in slide 1 is written without a paragraph. */
    private void assertTextBodiesCarryAParagraph(Path out) throws Exception {
        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<Element> bodies = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:txBody");
            assertThat(bodies)
                    .as("text bodies in slide 1 of %s", out)
                    .isNotEmpty();
            for (Element body : bodies) {
                assertThat(PackageAssertions.childNames(body))
                        .as("children of a <p:txBody> in slide 1 of %s%n%s",
                                out, pkg.text("ppt/slides/slide1.xml"))
                        .contains("a:p");
            }
        }
    }

    /** A portion added to a paragraph must be written as its own run. */
    @Test
    void aPortionAddedToAParagraphMustReachTheFileAsItsOwnRun() throws Exception {
        Path out = tempDir.resolve("portions.pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 200);
            shape.addTextFrame("Plain ");
            shape.getTextFrame().getParagraphs().get(0).getPortions()
                    .add(new Portion("and emphasised"));
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<Element> runs = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:txBody/a:p[1]/a:r");
            assertThat(runs)
                    .as("a:r elements in the first paragraph of %s%n%s",
                            out, pkg.text("ppt/slides/slide1.xml"))
                    .hasSize(2);
        }
    }
}
