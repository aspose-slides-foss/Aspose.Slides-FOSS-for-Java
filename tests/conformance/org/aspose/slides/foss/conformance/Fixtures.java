package org.aspose.slides.foss.conformance;

import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Decks written by something other than this library, for tests that need a realistic input.
 *
 * <p>Several defects only appear on a file the library did not itself produce, and a test that
 * feeds the library its own output cannot reach them. These fixtures are built with Apache POI
 * from the PowerPoint-authored template POI ships, so they carry a real slide master, the full
 * set of eleven slide layouts and a theme — the structure every real-world deck has and the
 * blank deck this library creates does not.</p>
 */
public final class Fixtures {

    /** Number of slide layouts in the template the fixtures are built from. */
    public static final int LAYOUTS_IN_FIXTURE = 11;

    private Fixtures() {
    }

    /**
     * Resolves a file in {@code tests/test_data}.
     *
     * <p>Anchored on the {@code basedir} property the build sets rather than on the working
     * directory, so that a test finds its fixture whatever launched it.</p>
     *
     * @param fileName the file name within {@code tests/test_data}
     * @return the path to it
     */
    public static Path testData(String fileName) {
        return Path.of(System.getProperty("basedir", "."))
                .resolve(Path.of("tests", "test_data", fileName));
    }

    /**
     * Writes a multi-slide deck with a real master, eleven layouts and a title on every slide.
     *
     * @param directory where to write it
     * @param fileName  the file name, including the extension
     * @param titles    one title per slide; at least one
     * @return the path written
     * @throws IOException if the file cannot be written
     */
    public static Path authoredDeck(Path directory, String fileName, String... titles)
            throws IOException {
        Path target = directory.resolve(fileName);
        try (XMLSlideShow show = new XMLSlideShow()) {
            XSLFSlideMaster master = show.getSlideMasters().get(0);
            XSLFSlideLayout layout = master.getLayout(SlideLayout.TITLE_ONLY);
            for (String title : titles) {
                XSLFSlide slide = show.createSlide(layout);
                XSLFTextShape placeholder = slide.getPlaceholder(0);
                if (placeholder != null) {
                    placeholder.setText(title);
                }
            }
            try (OutputStream out = Files.newOutputStream(target)) {
                show.write(out);
            }
        }
        return target;
    }

    /**
     * Writes a multi-slide deck like {@link #authoredDeck}, with speaker notes on some slides.
     *
     * <p>The notes are written by Apache POI, so the deck carries a notes master and notes slides
     * wired the way another producer wires them. POI names each notes slide after the slide part
     * it annotates. POI also relates every notes slide from the presentation part, which
     * PowerPoint does not do; those relationships are removed, so the deck has the shape
     * PowerPoint writes.</p>
     *
     * @param directory where to write it
     * @param fileName  the file name, including the extension
     * @param titles    one title per slide; at least one
     * @param notes     the notes text per slide, {@code null} for a slide without notes; the same
     *                  length as {@code titles}
     * @return the path written
     * @throws IOException if the file cannot be written
     */
    public static Path authoredDeckWithNotes(Path directory, String fileName, String[] titles,
                                             String[] notes) throws IOException {
        Path written = directory.resolve("poi-" + fileName);
        try (XMLSlideShow show = new XMLSlideShow()) {
            XSLFSlideMaster master = show.getSlideMasters().get(0);
            XSLFSlideLayout layout = master.getLayout(SlideLayout.TITLE_ONLY);
            for (int i = 0; i < titles.length; i++) {
                XSLFSlide slide = show.createSlide(layout);
                XSLFTextShape placeholder = slide.getPlaceholder(0);
                if (placeholder != null) {
                    placeholder.setText(titles[i]);
                }
                if (notes[i] != null) {
                    XSLFNotes notesSlide = show.getNotesSlide(slide);
                    for (XSLFTextShape shape : notesSlide.getPlaceholders()) {
                        if (shape.getTextType() == Placeholder.BODY) {
                            shape.setText(notes[i]);
                        }
                    }
                }
            }
            try (OutputStream out = Files.newOutputStream(written)) {
                show.write(out);
            }
        }
        return ZipSurgery.copyWithReplacement(written, directory.resolve(fileName),
                "ppt/_rels/presentation.xml.rels",
                rels -> rels.replaceAll("<Relationship [^>]*relationships/notesSlide\"[^>]*/>", ""));
    }
}
