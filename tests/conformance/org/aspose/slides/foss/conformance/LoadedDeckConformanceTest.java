package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.Presentation;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Opening a deck must expose the masters and layouts the deck actually contains.
 *
 * <p>This is the one place where the library's own report is the thing under test rather than
 * the file it wrote — a reader defect is only visible by comparing what the reader says with
 * what the file says. So the expected numbers are taken from the package itself, by counting
 * the layout parts in the ZIP, and never from the library.</p>
 *
 * <p>What breaks when this is wrong: {@code addEmptySlide(pres.getLayoutSlides().get(0))} — the
 * form every documented example uses — hands the new slide a layout that is not in the
 * document, so placeholder inheritance and theme resolution have nothing to resolve
 * against.</p>
 */
class LoadedDeckConformanceTest {

    @TempDir
    Path tempDir;

    /** The layouts reported must be the layouts the file contains. */
    @Test
    void openingADeckMustExposeTheLayoutsItContains() throws Exception {
        Path deck = Fixtures.authoredDeck(tempDir, "real.pptx", "First", "Second");

        long layoutPartsInFile;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            layoutPartsInFile = pkg.entryNames().stream()
                    .filter(name -> name.matches("ppt/slideLayouts/slideLayout\\d+\\.xml"))
                    .count();
        }
        assertThat(layoutPartsInFile)
                .as("the fixture must have several layouts for this test to mean anything")
                .isEqualTo(Fixtures.LAYOUTS_IN_FIXTURE);

        try (var pres = new Presentation(deck.toString())) {
            assertThat(pres.getLayoutSlides().size())
                    .as("layouts reported for a deck whose package contains %d of them",
                            layoutPartsInFile)
                    .isEqualTo((int) layoutPartsInFile);
        }
    }

    /** The master reported must be the master the file contains. */
    @Test
    void openingADeckMustExposeTheMasterItContains() throws Exception {
        Path deck = Fixtures.authoredDeck(tempDir, "real.pptx", "First");

        long masterPartsInFile;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            masterPartsInFile = pkg.entryNames().stream()
                    .filter(name -> name.matches("ppt/slideMasters/slideMaster\\d+\\.xml"))
                    .count();
        }

        long layoutPartsInFile;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            layoutPartsInFile = pkg.entryNames().stream()
                    .filter(name -> name.matches("ppt/slideLayouts/slideLayout\\d+\\.xml"))
                    .count();
        }

        try (var pres = new Presentation(deck.toString())) {
            assertThat(pres.getMasters().size())
                    .as("masters reported for a deck whose package contains %d of them",
                            masterPartsInFile)
                    .isEqualTo((int) masterPartsInFile);
            assertThat(pres.getMasters().get(0).getLayoutSlides().size())
                    .as("layouts on the master of a deck whose package contains %d of them",
                            layoutPartsInFile)
                    .isEqualTo((int) layoutPartsInFile);
        }
    }

    /** A slide must report the layout it is actually related to. */
    @Test
    void aSlideMustReportTheLayoutItIsRelatedTo() throws Exception {
        Path deck = Fixtures.authoredDeck(tempDir, "real.pptx", "First");

        String layoutNameInFile;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            String target = PackageAssertions.selectNodes(pkg,
                            "ppt/slides/_rels/slide1.xml.rels",
                            "//rel:Relationship[contains(@Type,'slideLayout')]").get(0)
                    .getAttribute("Target");
            String layoutPart = PptxPackage.resolveTarget("ppt/slides/slide1.xml", target);
            layoutNameInFile = PackageAssertions
                    .selectNodes(pkg, layoutPart, "//p:cSld").get(0).getAttribute("name");
        }
        assertThat(layoutNameInFile)
                .as("the fixture's layout must be named for this test to mean anything")
                .isNotEmpty();

        try (var pres = new Presentation(deck.toString())) {
            assertThat(pres.getSlides().get(0).getLayoutSlide().getName())
                    .as("layout name reported for a slide related to a layout named '%s'",
                            layoutNameInFile)
                    .isEqualTo(layoutNameInFile);
        }
    }
}
