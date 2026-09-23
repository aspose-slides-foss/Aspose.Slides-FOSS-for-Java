package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.ICommentAuthor;
import org.aspose.slides.foss.INotesSlide;
import org.aspose.slides.foss.ISlide;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.drawing.PointF;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Element;

import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Notes and comments belong to the slide that is related to them, not to the slide whose
 * number they share.
 *
 * <p>Part names carry no meaning in OPC: {@code notesSlide1.xml} is the notes of whichever
 * slide has a relationship to it. Producers number these parts independently of the slides —
 * a deck whose first slide has no notes commonly has its second slide's notes in
 * {@code notesSlide1.xml} — and a library that pairs them by number reads another slide's notes,
 * overwrites them, or deletes them and leaves the owning slide with a relationship to a part
 * that is gone, which is a file PowerPoint refuses to open.</p>
 *
 * <p>The decks here are valid packages whose part numbers differ from their slide numbers. The
 * notes deck is written by Apache POI and renumbered; the comments deck is renumbered the same
 * way.</p>
 */
class PartNumberingConformanceTest {

    static final String REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    static final String REL_NOTES_SLIDE = REL_BASE + "notesSlide";
    static final String REL_SLIDE = REL_BASE + "slide";
    static final String REL_COMMENTS = REL_BASE + "comments";

    @TempDir
    Path tempDir;

    // ---------------------------------------------------------------- notes

    /** Removing notes from a slide that has none must not delete another slide's notes. */
    @Test
    void removingNotesFromASlideWithoutNotesMustNotDeleteAnotherSlidesNotes() throws Exception {
        Path deck = notesDeckNumberedApartFromItsSlides();
        Path out = tempDir.resolve("removed-from-first.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.getSlides().get(0).getNotesSlideManager().removeNotesSlide();
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            assertThat(notesTextOf(pkg, slides.get(0))).as("notes of slide 1 in %s", out).isNull();
            assertThat(notesTextOf(pkg, slides.get(1))).as("notes of slide 2 in %s", out)
                    .isEqualTo("Note B");
            assertThat(notesTextOf(pkg, slides.get(2))).as("notes of slide 3 in %s", out)
                    .isEqualTo("Note C");
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /** Removing a slide's notes must remove the notes that slide is related to, and no other. */
    @Test
    void removingNotesMustRemoveTheNotesTheSlideIsRelatedTo() throws Exception {
        Path deck = notesDeckNumberedApartFromItsSlides();
        Path out = tempDir.resolve("removed-from-second.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.getSlides().get(1).getNotesSlideManager().removeNotesSlide();
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            assertThat(notesTextOf(pkg, slides.get(1))).as("notes of slide 2 in %s", out).isNull();
            assertThat(notesTextOf(pkg, slides.get(2))).as("notes of slide 3 in %s", out)
                    .isEqualTo("Note C");
            assertThat(pkg.entryNames().stream()
                    .filter(name -> name.matches("ppt/notesSlides/notesSlide\\d+\\.xml")))
                    .as("notes parts left in %s", out)
                    .containsExactlyElementsOf(PackageAssertions.relatedParts(pkg, slides.get(2),
                            REL_NOTES_SLIDE));
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /** Adding notes to a slide that has none must not take over another slide's notes. */
    @Test
    void addingNotesToASlideWithoutNotesMustNotTakeOverAnotherSlidesNotes() throws Exception {
        Path deck = notesDeckNumberedApartFromItsSlides();
        Path out = tempDir.resolve("added-to-first.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.getSlides().get(0).getNotesSlideManager().addNotesSlide()
                    .getNotesTextFrame().setText("Note A");
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            String[] expected = {"Note A", "Note B", "Note C"};
            for (int i = 0; i < slides.size(); i++) {
                String slide = slides.get(i);
                assertThat(notesTextOf(pkg, slide)).as("notes of slide %d in %s", i + 1, out)
                        .isEqualTo(expected[i]);
                String notes = PackageAssertions.relatedParts(pkg, slide, REL_NOTES_SLIDE).get(0);
                assertThat(PackageAssertions.relatedParts(pkg, notes, REL_SLIDE))
                        .as("slide that %s points back to in %s", notes, out)
                        .containsExactly(slide);
            }
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /**
     * The notes the library reports for a slide must be the notes the package relates to it.
     *
     * <p>A reader test: the expected text is read out of the package first.</p>
     */
    @Test
    void theNotesReportedForASlideMustBeTheNotesItIsRelatedTo() throws Exception {
        Path deck = notesDeckNumberedApartFromItsSlides();
        List<String> expected = new java.util.ArrayList<>();
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            for (String slide : PackageAssertions.slidePartNames(pkg)) {
                expected.add(notesTextOf(pkg, slide));
            }
        }
        assertThat(expected).as("notes in the fixture").containsExactly(null, "Note B", "Note C");

        try (var pres = new Presentation(deck.toString())) {
            for (int i = 0; i < expected.size(); i++) {
                INotesSlide notes = pres.getSlides().get(i).getNotesSlideManager().getNotesSlide();
                String reported = notes == null ? null : notes.getNotesTextFrame().getText();
                assertThat(reported)
                        .as("notes reported for slide %d, which the package relates to notes %s",
                                i + 1, expected.get(i))
                        .isEqualTo(expected.get(i));
            }
        }
    }

    // ---------------------------------------------------------------- comments

    /** Removing a comment must remove the comments part the slide is related to. */
    @Test
    void removingACommentMustRemoveTheCommentsPartTheSlideIsRelatedTo() throws Exception {
        Path deck = commentsDeckNumberedApartFromItsSlides();
        Path out = tempDir.resolve("comment-removed.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.getCommentAuthors().get(0).getComments().removeAt(0);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            for (String slide : PackageAssertions.slidePartNames(pkg)) {
                assertThat(PackageAssertions.relatedParts(pkg, slide, REL_COMMENTS))
                        .as("comments relationships of %s in %s", slide, out)
                        .isEmpty();
            }
            assertThat(pkg.entryNames())
                    .as("comments parts left in %s", out)
                    .noneMatch(name -> name.startsWith("ppt/comments/"));
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /** Saving a deck unchanged must leave each comment on the slide it was on. */
    @Test
    void resavingMustKeepEachCommentOnItsOwnSlide() throws Exception {
        Path deck = commentsDeckNumberedApartFromItsSlides();
        Path out = tempDir.resolve("comments-resaved.pptx");
        try (var pres = new Presentation(deck.toString())) {
            assertThat(pres.getSlides().get(0).getSlideComments(null))
                    .as("comments the library reports on slide 1, which has none in the file")
                    .isEmpty();
            assertThat(pres.getSlides().get(1).getSlideComments(null))
                    .as("comments the library reports on slide 2, which has one in the file")
                    .hasSize(1);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            assertThat(PackageAssertions.relatedParts(pkg, slides.get(0), REL_COMMENTS))
                    .as("comments relationships of slide 1 in %s", out)
                    .isEmpty();
            List<String> second = PackageAssertions.relatedParts(pkg, slides.get(1), REL_COMMENTS);
            assertThat(second).as("comments relationships of slide 2 in %s", out).hasSize(1);
            assertThat(commentTexts(pkg, second.get(0)))
                    .as("comments on slide 2 in %s", out)
                    .containsExactly("On slide two");
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /**
     * Cloning a commented slide past another slide must not delete the clone's comments part
     * while the clone still points at it.
     */
    @Test
    void cloningACommentedSlideMustLeaveEveryCommentsRelationshipResolvable() throws Exception {
        Path commented = tempDir.resolve("commented.pptx");
        try (var pres = new Presentation()) {
            ISlide first = pres.getSlides().get(0);
            pres.getSlides().addEmptySlide(pres.getLayoutSlides().get(0));
            ICommentAuthor author = pres.getCommentAuthors().addAuthor("Reviewer", "RV");
            author.getComments().addComment("On slide one", first, new PointF(1f, 1f),
                    LocalDateTime.of(2026, 1, 2, 3, 4, 5));
            pres.save(commented.toString(), SaveFormat.PPTX);
        }

        Path out = tempDir.resolve("cloned-commented.pptx");
        try (var pres = new Presentation(commented.toString())) {
            pres.getSlides().addClone(pres.getSlides().get(0));
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            assertThat(slides).as("slides in %s", out).hasSize(3);
            List<String> first = PackageAssertions.relatedParts(pkg, slides.get(0), REL_COMMENTS);
            assertThat(first).as("comments relationships of slide 1 in %s", out).hasSize(1);
            assertThat(commentTexts(pkg, first.get(0))).containsExactly("On slide one");
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    // ---------------------------------------------------------------- fixtures

    /**
     * Three slides; slide 1 has no notes, slide 2's notes are {@code notesSlide1.xml} and slide
     * 3's are {@code notesSlide2.xml}.
     */
    private Path notesDeckNumberedApartFromItsSlides() throws IOException {
        Path written = Fixtures.authoredDeckWithNotes(tempDir, "notes-by-slide-number.pptx",
                new String[]{"Slide A", "Slide B", "Slide C"},
                new String[]{null, "Note B", "Note C"});
        String second;
        String third;
        try (PptxPackage pkg = PptxPackage.open(written)) {
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            second = PackageAssertions.relatedParts(pkg, slides.get(1), REL_NOTES_SLIDE).get(0);
            third = PackageAssertions.relatedParts(pkg, slides.get(2), REL_NOTES_SLIDE).get(0);
        }
        Path step = ZipSurgery.copyWithRenamedPart(written, tempDir.resolve("notes-step.pptx"),
                second, "ppt/notesSlides/notesSlide1.xml");
        Path deck = ZipSurgery.copyWithRenamedPart(step, tempDir.resolve("notes-renumbered.pptx"),
                third, "ppt/notesSlides/notesSlide2.xml");

        try (PptxPackage pkg = PptxPackage.open(deck)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            assertThat(slides).as("slides of the fixture").hasSize(3);
            assertThat(PackageAssertions.relatedParts(pkg, slides.get(0), REL_NOTES_SLIDE))
                    .as("notes of slide 1 in the fixture").isEmpty();
            assertThat(PackageAssertions.relatedParts(pkg, slides.get(1), REL_NOTES_SLIDE))
                    .as("notes of slide 2 in the fixture")
                    .containsExactly("ppt/notesSlides/notesSlide1.xml");
            assertThat(PackageAssertions.relatedParts(pkg, slides.get(2), REL_NOTES_SLIDE))
                    .as("notes of slide 3 in the fixture")
                    .containsExactly("ppt/notesSlides/notesSlide2.xml");
        }
        return deck;
    }

    /** Two slides; the only comment is on slide 2, in {@code comment1.xml}. */
    private Path commentsDeckNumberedApartFromItsSlides() throws IOException {
        Path written = tempDir.resolve("comments-by-slide-number.pptx");
        try (var pres = new Presentation()) {
            ISlide second = pres.getSlides().addEmptySlide(pres.getLayoutSlides().get(0));
            ICommentAuthor author = pres.getCommentAuthors().addAuthor("Reviewer", "RV");
            author.getComments().addComment("On slide two", second, new PointF(1f, 1f),
                    LocalDateTime.of(2026, 1, 2, 3, 4, 5));
            pres.save(written.toString(), SaveFormat.PPTX);
        }
        String commentsPart;
        try (PptxPackage pkg = PptxPackage.open(written)) {
            commentsPart = PackageAssertions.relatedParts(pkg,
                    PackageAssertions.slidePartNames(pkg).get(1), REL_COMMENTS).get(0);
        }
        Path deck = ZipSurgery.copyWithRenamedPart(written,
                tempDir.resolve("comments-renumbered.pptx"), commentsPart,
                "ppt/comments/comment1.xml");

        try (PptxPackage pkg = PptxPackage.open(deck)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            assertThat(PackageAssertions.relatedParts(pkg, slides.get(0), REL_COMMENTS))
                    .as("comments of slide 1 in the fixture").isEmpty();
            assertThat(PackageAssertions.relatedParts(pkg, slides.get(1), REL_COMMENTS))
                    .as("comments of slide 2 in the fixture")
                    .containsExactly("ppt/comments/comment1.xml");
        }
        return deck;
    }

    // ---------------------------------------------------------------- package reads

    /** The notes body text of the notes a slide is related to, or {@code null} if it has none. */
    private static String notesTextOf(PptxPackage pkg, String slidePart) throws IOException {
        List<String> notes = PackageAssertions.relatedParts(pkg, slidePart, REL_NOTES_SLIDE);
        if (notes.isEmpty()) {
            return null;
        }
        assertThat(notes).as("notesSlide relationships of %s in %s", slidePart, pkg.path())
                .hasSize(1);
        return NotesMasterConformanceTest.notesBodyText(pkg, notes.get(0));
    }

    private static List<String> commentTexts(PptxPackage pkg, String commentsPart)
            throws IOException {
        return PackageAssertions.selectNodes(pkg, commentsPart, "//p:cm/p:text").stream()
                .map(Element::getTextContent)
                .toList();
    }
}
