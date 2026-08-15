package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Opening a file and saving it again without changing anything must not change the file.
 *
 * <p>Re-indenting XML that is already indented adds whitespace text nodes on every pass, so a
 * file that is opened and saved in a loop — the shape of every batch job — grows without bound
 * while its content stays the same. It also makes byte length useless as a "did anything
 * actually change" signal, which is the cheapest such signal there is.</p>
 */
class ResaveConformanceTest {

    @TempDir
    Path tempDir;

    /** A save with no modification in between must not change the size of the file. */
    @Test
    void savingAgainWithoutChangesMustNotChangeTheFileSize() throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "source.pptx", "First", "Second");

        Path first = tempDir.resolve("first.pptx");
        try (var pres = new Presentation(source.toString())) {
            pres.save(first.toString(), SaveFormat.PPTX);
        }

        Path second = tempDir.resolve("second.pptx");
        try (var pres = new Presentation(first.toString())) {
            pres.save(second.toString(), SaveFormat.PPTX);
        }

        assertThat(Files.size(second))
                .as("size after re-saving %s unchanged (first save was %d bytes)",
                        first, Files.size(first))
                .isEqualTo(Files.size(first));
    }

    /** Neither may the slide XML itself grow. */
    @Test
    void savingAgainWithoutChangesMustNotChangeTheSlideXml() throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "source.pptx", "First");

        Path first = tempDir.resolve("first.pptx");
        try (var pres = new Presentation(source.toString())) {
            pres.save(first.toString(), SaveFormat.PPTX);
        }

        Path second = tempDir.resolve("second.pptx");
        try (var pres = new Presentation(first.toString())) {
            pres.save(second.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage before = PptxPackage.open(first);
             PptxPackage after = PptxPackage.open(second)) {
            assertThat(after.text("ppt/slides/slide1.xml"))
                    .as("slide 1 of %s after a no-op re-save", second)
                    .isEqualTo(before.text("ppt/slides/slide1.xml"));
        }
    }
}
