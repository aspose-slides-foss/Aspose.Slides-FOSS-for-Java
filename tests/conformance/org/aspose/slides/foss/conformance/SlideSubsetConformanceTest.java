package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Element;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Asking for some of the slides must not silently produce all of them.
 *
 * <p>Ignoring an argument the caller passed is worse than refusing it: the caller who asked for
 * one slide and got the whole deck has no signal that anything went wrong, and may well be
 * sending the rest of the deck to someone who was not supposed to see it.</p>
 */
class SlideSubsetConformanceTest {

    @TempDir
    Path tempDir;

    /** Saving a one-slide subset of a three-slide deck must write one slide. */
    @Test
    void savingASubsetMustWriteOnlyTheRequestedSlides() throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "three.pptx", "First", "Second", "Third");
        Path out = tempDir.resolve("subset.pptx");
        try (var pres = new Presentation(source.toString())) {
            pres.save(out.toString(), new int[]{1}, SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertRegisteredSlideCount(pkg, 1);
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            // The count alone would be satisfied by keeping the wrong slide.
            assertThat(slideTexts(pkg))
                    .as("the slides kept in %s, which was asked for slide 1 of "
                            + "[First, Second, Third]", out)
                    .containsExactly("Second");
        }
    }

    /** The same rule applies to the stream overload. */
    @Test
    void savingASubsetToAStreamMustWriteOnlyTheRequestedSlides() throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "three.pptx", "First", "Second", "Third");
        Path out = tempDir.resolve("subset-stream.pptx");
        try (var pres = new Presentation(source.toString())) {
            var buffer = new ByteArrayOutputStream();
            pres.save(buffer, new int[]{0, 2}, SaveFormat.PPTX);
            java.nio.file.Files.write(out, buffer.toByteArray());
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertRegisteredSlideCount(pkg, 2);
            // Kept in document order, whatever order they were named in.
            assertThat(slideTexts(pkg))
                    .as("the slides kept in %s, which was asked for slides 0 and 2 of "
                            + "[First, Second, Third]", out)
                    .containsExactly("First", "Third");
        }
    }

    /**
     * Returns the text of each slide in the package, in the order {@code p:sldIdLst} gives.
     *
     * <p>Read from the slide parts themselves, so it says which slides the file has rather
     * than how many.</p>
     */
    private static List<String> slideTexts(PptxPackage pkg) throws IOException {
        List<String> texts = new ArrayList<>();
        for (Element sldId : PackageAssertions.selectNodes(
                pkg, "ppt/presentation.xml", "//p:sldIdLst/p:sldId")) {
            String relId = sldId.getAttributeNS(PptxPackage.NS_REL, "id");
            List<Element> rel = PackageAssertions.selectNodes(
                    pkg, "ppt/_rels/presentation.xml.rels",
                    "//rel:Relationship[@Id='" + relId + "']");
            String part = PptxPackage.resolveTarget(
                    "ppt/presentation.xml", rel.get(0).getAttribute("Target"));
            var text = new StringBuilder();
            for (Element t : PackageAssertions.selectNodes(pkg, part, "//a:t")) {
                text.append(t.getTextContent());
            }
            texts.add(text.toString());
        }
        return texts;
    }
}
