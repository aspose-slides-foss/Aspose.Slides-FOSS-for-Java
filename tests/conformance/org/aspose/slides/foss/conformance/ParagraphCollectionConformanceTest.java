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
