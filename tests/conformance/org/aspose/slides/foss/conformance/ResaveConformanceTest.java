package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;

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

    /**
     * Opening a deck this library did not write and saving it must carry every part across
     * and invent none.
     *
     * <p>A part that is silently dropped is invisible in everything but the file: the deck
     * still opens, and the chart, the embedded workbook or the theme that went missing is
     * only noticed by whoever needed it. Asserted by name and not by count, so a part
     * dropped and another added cannot cancel out.</p>
     *
     * <p>The other direction is asserted too. A part appearing that the source did not have
     * means the writer substituted something of its own - a synthetic master, a default
     * theme - for what was in the file, and the deck that comes back is then not the deck
     * that went in even though nothing was lost.</p>
     */
    @Test
    void openingAndSavingADeckMustLoseNoPartOfIt() throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "authored.pptx", "First", "Second", "Third");

        Path out = tempDir.resolve("round-tripped.pptx");
        try (var pres = new Presentation(source.toString())) {
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage before = PptxPackage.open(source);
             PptxPackage after = PptxPackage.open(out)) {
            var lost = new ArrayList<>(before.entryNames());
            lost.removeAll(after.entryNames());
            assertThat(lost)
                    .as("parts of %s that did not survive being saved to %s; the file had %d "
                            + "parts and now has %d", source, out,
                            before.entryNames().size(), after.entryNames().size())
                    .isEmpty();

            var added = new ArrayList<>(after.entryNames());
            added.removeAll(before.entryNames());
            assertThat(added)
                    .as("parts in %s that %s did not have; the file had %d parts and now has %d",
                            out, source,
                            before.entryNames().size(), after.entryNames().size())
                    .isEmpty();
        }
    }
}
