package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.ICommentAuthor;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.drawing.PointF;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Deleting a part means deleting everything that points at it.
 *
 * <p>A part removal is four operations at once: the bytes, the part's own {@code .rels}, the
 * {@code <Relationship>} in the owner's {@code .rels}, and the {@code <Override>} in
 * {@code [Content_Types].xml}. Doing only the first leaves a package that names a part it does
 * not contain — which strict readers reject outright and PowerPoint refuses, so the user loses
 * the whole document by deleting one comment.</p>
 */
class PartRemovalConformanceTest {

    @TempDir
    Path tempDir;

    /** Removing speaker notes must not leave the notes relationship or override behind. */
    @Test
    void removingSpeakerNotesMustNotLeaveTheNotesPartDeclared() throws Exception {
        Path withNotes = tempDir.resolve("with-notes.pptx");
        try (var pres = new Presentation()) {
            pres.getSlides().get(0).getNotesSlideManager().addNotesSlide()
                    .getNotesTextFrame().setText("Speaker notes go here.");
            pres.save(withNotes.toString(), SaveFormat.PPTX);
        }

        Path removed = tempDir.resolve("notes-removed.pptx");
        try (var pres = new Presentation(withNotes.toString())) {
            pres.getSlides().get(0).getNotesSlideManager().removeNotesSlide();
            pres.save(removed.toString(), SaveFormat.PPTX);
        }

        assertNothingRefersToAMissingPart(removed);
        ThirdPartyReadBack.assertOpens(removed);
    }

    /** Removing the last comment must not leave the comments relationship or override behind. */
    @Test
    void removingTheLastCommentMustNotLeaveTheCommentsPartDeclared() throws Exception {
        Path withComment = tempDir.resolve("with-comment.pptx");
        try (var pres = new Presentation()) {
            ICommentAuthor author = pres.getCommentAuthors().addAuthor("Jane Smith", "JS");
            author.getComments().addComment("Review this slide", pres.getSlides().get(0),
                    new PointF(2.0f, 2.0f), LocalDateTime.now());
            pres.save(withComment.toString(), SaveFormat.PPTX);
        }

        Path removed = tempDir.resolve("comment-removed.pptx");
        try (var pres = new Presentation(withComment.toString())) {
            pres.getCommentAuthors().get(0).getComments().removeAt(0);
            pres.save(removed.toString(), SaveFormat.PPTX);
        }

        assertNothingRefersToAMissingPart(removed);
        ThirdPartyReadBack.assertOpens(removed);
    }

    private static void assertNothingRefersToAMissingPart(Path pptx) throws IOException {
        try (PptxPackage pkg = PptxPackage.open(pptx)) {
            PackageAssertions.assertNoContentTypeOverrideNamesAMissingPart(pkg);
            PackageAssertions.assertEveryInternalRelationshipTargetExists(pkg);
            PackageAssertions.assertEveryRelationshipReferenceResolves(pkg);
            assertThat(pkg.entryNames())
                    .as("parts left in %s", pptx)
                    .isNotEmpty();
        }
    }
}
