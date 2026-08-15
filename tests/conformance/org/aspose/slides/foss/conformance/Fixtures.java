package org.aspose.slides.foss.conformance;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
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
}
