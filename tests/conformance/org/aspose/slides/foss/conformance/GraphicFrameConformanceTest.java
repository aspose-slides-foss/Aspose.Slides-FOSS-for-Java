package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.ITable;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Element;

import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A graphic frame must use the element name DrawingML defines.
 *
 * <p>{@code CT_NonVisualGraphicFrameProperties} allows exactly one child,
 * {@code <a:graphicFrameLocks>}. {@code <a:graphicFrameLocking>} is not an element in the
 * standard — it reads like one because {@code CT_GraphicalObjectFrameLocking} is the name of
 * the type. PowerPoint tolerates it by discarding it, which is why every table this library has
 * ever written has been invalid without anyone noticing.</p>
 */
class GraphicFrameConformanceTest {

    @TempDir
    Path tempDir;

    /** A table's non-visual properties must carry a:graphicFrameLocks. */
    @Test
    void aTableMustLockGroupingWithTheElementTheStandardDefines() throws Exception {
        Path out = writeDeckWithTable();

        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<Element> properties = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml",
                    "//p:graphicFrame/p:nvGraphicFramePr/p:cNvGraphicFramePr");
            assertThat(properties)
                    .as("p:cNvGraphicFramePr elements in slide 1 of %s", out)
                    .hasSize(1);
            PackageAssertions.assertChildrenInSchemaOrder(
                    properties.get(0), "a:graphicFrameLocks", "a:extLst");
        }
    }

    /** The misspelt element must not appear in any part of the package. */
    @Test
    void theMisspeltGraphicFrameLockingElementMustNotAppearAnywhere() throws Exception {
        Path out = writeDeckWithTable();

        try (PptxPackage pkg = PptxPackage.open(out)) {
            for (String partName : pkg.xmlPartNames()) {
                assertThat(pkg.text(partName))
                        .as("part %s of %s", partName, out)
                        .doesNotContain("graphicFrameLocking");
            }
        }
    }

    private Path writeDeckWithTable() throws Exception {
        Path out = tempDir.resolve("table.pptx");
        try (var pres = new Presentation()) {
            ITable table = pres.getSlides().get(0).getShapes()
                    .addTable(50, 50, new double[]{120, 120, 120}, new double[]{40, 40});
            table.getRows().get(0).get(0).getTextFrame().setText("Name");
            table.getRows().get(0).get(1).getTextFrame().setText("Value");
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        return out;
    }
}
