package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.IAutoShape;
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
 * The extended properties a file advertises must describe the file.
 *
 * <p>{@code docProps/app.xml} is what a document management system, a search indexer and the
 * Explorer properties pane read without opening the presentation. Copying it out of the
 * template means every file this library writes reports zero slides, no matter how many it
 * has — cheap to get right, and misleading in exactly the places nobody checks.</p>
 */
class DocumentPropertiesConformanceTest {

    @TempDir
    Path tempDir;

    /** The slide count in docProps/app.xml must match the slides in the package. */
    @Test
    void theAdvertisedSlideCountMustMatchTheSlidesInTheFile() throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "three.pptx", "First", "Second", "Third");
        Path out = tempDir.resolve("resaved.pptx");
        try (var pres = new Presentation(source.toString())) {
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            int slides = PackageAssertions.registeredSlideCount(pkg);
            assertThat(slides)
                    .as("the fixture must have several slides for this test to mean anything")
                    .isEqualTo(3);

            List<Element> declared = PackageAssertions.selectNodes(
                    pkg, "docProps/app.xml", "//ep:Slides");
            assertThat(declared).as("ep:Slides element in docProps/app.xml of %s", out)
                    .hasSize(1);
            assertThat(declared.get(0).getTextContent())
                    .as("slide count advertised by docProps/app.xml of %s, which has %d slides",
                            out, slides)
                    .isEqualTo(String.valueOf(slides));
        }
    }

    /**
     * Words separated by any whitespace must be counted as separate words.
     *
     * <p>A tab or a newline between two words is as much a separator as a space. Counting on
     * spaces alone under-reports every document that uses one, and the under-report is what a
     * search indexer stores.</p>
     */
    @Test
    void theAdvertisedWordCountMustCountWordsSeparatedByAnyWhitespace() throws Exception {
        Path out = tempDir.resolve("words.pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 200);
            shape.addTextFrame("alpha\tbeta gamma");
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<Element> words = PackageAssertions.selectNodes(pkg, "docProps/app.xml",
                    "//ep:Words");
            assertThat(words)
                    .as("ep:Words element in docProps/app.xml of %s", out)
                    .hasSize(1);
            assertThat(words.get(0).getTextContent())
                    .as("word count advertised by docProps/app.xml of %s, whose one text run "
                            + "is 'alpha<tab>beta gamma'%n%s", out, pkg.text("docProps/app.xml"))
                    .isEqualTo("3");
        }
    }
}
