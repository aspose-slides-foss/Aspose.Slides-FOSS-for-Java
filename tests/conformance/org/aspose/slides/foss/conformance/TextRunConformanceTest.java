package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.FillType;
import org.aspose.slides.foss.IAutoShape;
import org.aspose.slides.foss.IPortionFormat;
import org.aspose.slides.foss.NullableBool;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.ShapeType;
import org.aspose.slides.foss.drawing.Color;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Element;

import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Character formatting must be written where the schema puts it: before the text.
 *
 * <p>{@code CT_RegularTextRun} is a sequence, {@code rPr?} then {@code t}. A run whose
 * {@code <a:rPr>} follows its {@code <a:t>} is out of sequence, and readers drop the properties
 * rather than reordering them — so the text appears and the bold, the size and the colour the
 * caller asked for are silently gone.</p>
 */
class TextRunConformanceTest {

    @TempDir
    Path tempDir;

    /** The run properties element must precede the text element. */
    @Test
    void characterFormattingMustBeWrittenBeforeTheTextItApplies() throws Exception {
        Path out = writeFormattedRun("run-order.pptx");

        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<Element> runs = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:txBody/a:p/a:r");
            assertThat(runs).as("runs in slide 1 of %s", out).hasSize(1);
            assertThat(PackageAssertions.childNames(runs.get(0)))
                    .as("children of <a:r> in %s", out)
                    .containsExactly("a:rPr", "a:t");
        }
    }

    /** The same rule stated through the schema-order helper, for the whole run. */
    @Test
    void runChildrenMustFollowTheSchemaSequence() throws Exception {
        Path out = writeFormattedRun("run-sequence.pptx");

        try (PptxPackage pkg = PptxPackage.open(out)) {
            Element run = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:txBody/a:p/a:r").get(0);
            PackageAssertions.assertChildrenInSchemaOrder(run, "a:rPr", "a:t");
        }
    }

    /** An independent reader must see the formatting that was asked for. */
    @Test
    void anIndependentReaderMustSeeTheFormattingThatWasAskedFor() throws Exception {
        Path out = writeFormattedRun("run-readback.pptx");

        try (PptxPackage pkg = PptxPackage.open(out)) {
            Element run = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:txBody/a:p/a:r").get(0);
            List<Element> properties = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:txBody/a:p/a:r/a:rPr");
            assertThat(properties).as("a:rPr in %s", out).hasSize(1);
            assertThat(PackageAssertions.childNames(run).indexOf("a:rPr"))
                    .as("position of a:rPr among the children of a:r in %s", out)
                    .isZero();
        }
        assertThat(ThirdPartyReadBack.textOf(out, 0))
                .as("text an independent reader finds in %s", out)
                .anySatisfy(text -> assertThat(text).contains("Bold 24pt blue"));
    }

    private Path writeFormattedRun(String fileName) throws Exception {
        Path out = tempDir.resolve(fileName);
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 150);
            shape.addTextFrame("Bold 24pt blue");
            IPortionFormat format = shape.getTextFrame().getParagraphs().get(0)
                    .getPortions().get(0).getPortionFormat();
            format.setFontHeight(24);
            format.setFontBold(NullableBool.TRUE);
            format.setFontItalic(NullableBool.TRUE);
            format.getFillFormat().setFillType(FillType.SOLID);
            format.getFillFormat().getSolidFillColor().setColor(Color.fromArgb(255, 0, 70, 127));
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        return out;
    }
}
