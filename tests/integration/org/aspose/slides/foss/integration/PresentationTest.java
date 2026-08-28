package org.aspose.slides.foss.integration;
import org.aspose.slides.foss.*;

import org.aspose.slides.foss.conformance.Fixtures;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Integration tests for Presentation create / load / save / properties.
 */
class PresentationTest implements AutoCloseable {

    @TempDir
    Path tempDir;

    @Override
    public void close() {
        // TempDir handles cleanup
    }

    /**
     * Saves a Presentation to a
     * temporary file, disposes the original, and reopens from that file.
     */
    private Presentation saveAndReopen(Presentation pres) throws IOException {
        String path = tempDir.resolve("roundtrip.pptx").toString();
        pres.save(path, SaveFormat.PPTX);
        pres.dispose();
        return new Presentation(path);
    }

    /** A brand-new presentation has exactly 1 slide. */
    @Test
    void testCreateEmpty() {
        try (var pres = new Presentation()) {
            assertThat(pres.getSlides().size()).isEqualTo(1);
        }
    }

    /** Round-trip: create -> save -> reload preserves slide count. */
    @Test
    void testSaveAndReload() throws IOException {
        try (var pres = new Presentation()) {
            try (var pres2 = saveAndReopen(pres)) {
                assertThat(pres2.getSlides().size()).isEqualTo(1);
            }
        }
    }

    /** Saving to a ByteArrayOutputStream produces a non-empty buffer. */
    @Test
    void testSaveToStream() throws IOException {
        try (var pres = new Presentation()) {
            ByteArrayOutputStream buf = new ByteArrayOutputStream();
            pres.save(buf, SaveFormat.PPTX);
            assertThat(buf.size()).isGreaterThan(0);
        }
    }

    /** Presentation can be used as a try-with-resources (context manager). */
    @Test
    void testContextManager() {
        try (Presentation pres = new Presentation()) {
            assertThat(pres.getSlides().size()).isGreaterThanOrEqualTo(1);
        }
    }

    /** first_slide_number persists across save/reload. */
    @Test
    void testFirstSlideNumber() throws IOException {
        try (var pres = new Presentation()) {
            pres.setFirstSlideNumber(5);
            assertThat(pres.getFirstSlideNumber()).isEqualTo(5);

            try (var pres2 = saveAndReopen(pres)) {
                assertThat(pres2.getFirstSlideNumber()).isEqualTo(5);
            }
        }
    }

    /**
     * Load a .pptx this library did not write and verify what it reports about it.
     *
     * <p>This used to point at a file under {@code test_data} that is not in the
     * repository, so {@code assumeThat} skipped it on every run and it never
     * checked anything. It now builds its input with a third-party writer, and
     * asserts the masters and layouts of the deck rather than only that opening
     * it did not throw.</p>
     */
    @Test
    void testLoadExisting() throws Exception {
        Path path = Fixtures.authoredDeck(tempDir, "authored.pptx", "First", "Second");
        try (var pres = new Presentation(path.toString())) {
            assertThat(pres.getSlides().size()).isEqualTo(2);
            assertThat(pres.getMasters().size()).isEqualTo(1);
            assertThat(pres.getLayoutSlides().size()).isEqualTo(Fixtures.LAYOUTS_IN_FIXTURE);
            assertThat(pres.getSlides().get(0).getLayoutSlide()).isNotNull();
        }
    }

    /** Calling dispose() twice must not raise. */
    @Test
    void testDisposeIsIdempotent() {
        Presentation pres = new Presentation();
        pres.dispose();
        pres.dispose(); // second call should be harmless
    }

    /** Adding a slide increases slide count to 2. */
    @Test
    void testSlideCountAfterAdd() throws IOException {
        try (var pres = new Presentation()) {
            ILayoutSlide layout = pres.getLayoutSlides().get(0);
            pres.getSlides().addEmptySlide(layout);
            assertThat(pres.getSlides().size()).isEqualTo(2);
        }
    }
}
