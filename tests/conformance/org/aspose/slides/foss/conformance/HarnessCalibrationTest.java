package org.aspose.slides.foss.conformance;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Calibrates the harness itself against a known-good package and against known damage.
 *
 * <p>An assertion that never fires proves nothing. Before any of these rules is used to
 * report a defect, it has to be shown to pass on a package that is correct and to fail on a
 * package that is not — otherwise a rule that silently does nothing reads as a clean bill of
 * health. These tests are expected to pass; if one of them fails, the harness is broken, not
 * the library.</p>
 */
class HarnessCalibrationTest {

    @TempDir
    Path tempDir;

    /** Every package rule passes on a deck written by an independent implementation. */
    @Test
    void packageRulesAcceptAKnownGoodDeck() throws IOException {
        Path good = Fixtures.authoredDeck(tempDir, "good.pptx", "One", "Two");
        try (PptxPackage pkg = PptxPackage.open(good)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            PackageAssertions.assertRegisteredSlideCount(pkg, 2);
        }
        ThirdPartyReadBack.assertReadsAs(good, 2, 1, "One");
    }

    /** The slide count is read from p:sldIdLst, not from the number of slide parts. */
    @Test
    void slideCountIgnoresSlidePartsThatAreNotRegistered() throws IOException {
        Path good = Fixtures.authoredDeck(tempDir, "good.pptx", "One");
        Path tampered = ZipSurgery.copyWithExtraPart(good, tempDir.resolve("orphan.pptx"),
                "ppt/slides/slide99.xml", "<?xml version=\"1.0\"?><root/>".getBytes());
        try (PptxPackage pkg = PptxPackage.open(tampered)) {
            assertThat(PackageAssertions.registeredSlideCount(pkg))
                    .as("an unregistered slide part must not be counted as a slide")
                    .isEqualTo(1);
        }
    }

    /** A dangling r:embed is detected. */
    @Test
    void dangingRelationshipReferenceIsDetected() throws IOException {
        Path good = Fixtures.authoredDeck(tempDir, "good.pptx", "One");
        Path tampered = ZipSurgery.copyWithReplacement(good, tempDir.resolve("dangling.pptx"),
                "ppt/slides/slide1.xml",
                xml -> xml.replace("</p:spTree>",
                        "<p:pic xmlns:p=\"" + PptxPackage.NS_P + "\">"
                                + "<p:blipFill xmlns:a=\"" + PptxPackage.NS_A + "\">"
                                + "<a:blip xmlns:r=\"" + PptxPackage.NS_REL
                                + "\" r:embed=\"rIdNoSuchThing\"/>"
                                + "</p:blipFill></p:pic></p:spTree>"));
        try (PptxPackage pkg = PptxPackage.open(tampered)) {
            assertThat(catchAssertionError(
                    () -> PackageAssertions.assertEveryRelationshipReferenceResolves(pkg)))
                    .as("the rule must reject an r:embed with no matching Relationship")
                    .isNotNull()
                    .hasMessageContaining("rIdNoSuchThing");
        }
    }

    /** A part with no content type is detected. */
    @Test
    void partWithoutAContentTypeIsDetected() throws IOException {
        Path good = Fixtures.authoredDeck(tempDir, "good.pptx", "One");
        Path tampered = ZipSurgery.copyWithExtraPart(good, tempDir.resolve("untyped.pptx"),
                "ppt/media/image1.unknownext", new byte[]{1, 2, 3});
        try (PptxPackage pkg = PptxPackage.open(tampered)) {
            assertThat(catchAssertionError(
                    () -> PackageAssertions.assertEveryPartHasAContentType(pkg)))
                    .as("the rule must reject a part with no Override and no Default")
                    .isNotNull()
                    .hasMessageContaining("image1.unknownext");
        }
    }

    /** A content-type override naming a deleted part is detected. */
    @Test
    void overrideNamingAMissingPartIsDetected() throws IOException {
        Path good = Fixtures.authoredDeck(tempDir, "good.pptx", "One", "Two");
        Path tampered = ZipSurgery.copyWithoutPart(good, tempDir.resolve("halfdeleted.pptx"),
                "ppt/slides/slide2.xml");
        try (PptxPackage pkg = PptxPackage.open(tampered)) {
            assertThat(catchAssertionError(
                    () -> PackageAssertions.assertNoContentTypeOverrideNamesAMissingPart(pkg)))
                    .as("the rule must reject an Override for a part that was removed")
                    .isNotNull()
                    .hasMessageContaining("/ppt/slides/slide2.xml");
        }
    }

    /** Child order violations are detected. */
    @Test
    void childOrderViolationIsDetected() throws IOException {
        Path good = Fixtures.authoredDeck(tempDir, "good.pptx", "One");
        try (PptxPackage pkg = PptxPackage.open(good)) {
            var spTree = PackageAssertions.selectNodes(pkg, "ppt/slides/slide1.xml",
                    "//p:cSld/p:spTree").get(0);
            assertThat(catchAssertionError(() -> PackageAssertions.assertChildrenInSchemaOrder(
                    spTree, "p:sp", "p:nvGrpSpPr", "p:grpSpPr")))
                    .as("the rule must reject children that appear out of sequence")
                    .isNotNull();
        }
    }

    /** Files.size is stable, so the byte-length comparisons the tests make are meaningful. */
    @Test
    void writtenFixtureIsANonEmptyZip() throws IOException {
        Path good = Fixtures.authoredDeck(tempDir, "good.pptx", "One");
        assertThat(Files.size(good)).isGreaterThan(0);
        try (PptxPackage pkg = PptxPackage.open(good)) {
            assertThat(pkg.entryNames()).contains("[Content_Types].xml", "ppt/presentation.xml");
        }
    }

    private static AssertionError catchAssertionError(ThrowingRunnable runnable) {
        try {
            runnable.run();
            return null;
        } catch (AssertionError e) {
            return e;
        } catch (Exception e) {
            throw new IllegalStateException(e);
        }
    }

    @FunctionalInterface
    private interface ThrowingRunnable {
        void run() throws Exception;
    }
}
