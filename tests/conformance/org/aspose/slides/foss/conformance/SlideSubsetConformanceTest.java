package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayOutputStream;
import java.nio.file.Path;

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
        }
    }
}
