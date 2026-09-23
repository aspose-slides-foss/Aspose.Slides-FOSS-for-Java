package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.ICommentAuthor;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.drawing.PointF;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

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

    /**
     * Removing a comment author must remove the author from the file, not only their comments.
     *
     * <p>Each way the API offers to remove the last author is swept. The author used to stay in
     * {@code ppt/commentAuthors.xml} whenever no author was left, because a list with nothing in
     * it was never written back.</p>
     */
    @ParameterizedTest(name = "{0}")
    @ValueSource(strings = {"removeAt", "remove", "clear"})
    void removingTheLastCommentAuthorMustRemoveThemFromTheFile(String how) throws Exception {
        Path withComment = tempDir.resolve("with-author.pptx");
        try (var pres = new Presentation()) {
            ICommentAuthor author = pres.getCommentAuthors().addAuthor("Jane Smith", "JS");
            author.getComments().addComment("Review this slide", pres.getSlides().get(0),
                    new PointF(2.0f, 2.0f), LocalDateTime.now());
            pres.save(withComment.toString(), SaveFormat.PPTX);
        }

        Path removed = tempDir.resolve("author-removed-" + how + ".pptx");
        try (var pres = new Presentation(withComment.toString())) {
            switch (how) {
                case "removeAt" -> pres.getCommentAuthors().removeAt(0);
                case "remove" -> pres.getCommentAuthors().remove(pres.getCommentAuthors().get(0));
                default -> pres.getCommentAuthors().clear();
            }
            pres.save(removed.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(removed)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            if (pkg.hasPart("ppt/commentAuthors.xml")) {
                assertThat(PackageAssertions.selectNodes(pkg, "ppt/commentAuthors.xml",
                        "//p:cmAuthor"))
                        .as("authors left in ppt/commentAuthors.xml of %s%n%s", removed,
                                pkg.text("ppt/commentAuthors.xml"))
                        .isEmpty();
            }
            assertThat(pkg.entryNames())
                    .as("comments parts left in %s", removed)
                    .noneMatch(name -> name.startsWith("ppt/comments/"));
        }
        ThirdPartyReadBack.assertOpens(removed);
    }

    /** Removing one of two authors must leave exactly the other one in the file. */
    @Test
    void removingOneOfTwoCommentAuthorsMustLeaveTheOther() throws Exception {
        Path withComments = tempDir.resolve("two-authors.pptx");
        try (var pres = new Presentation()) {
            ICommentAuthor jane = pres.getCommentAuthors().addAuthor("Jane Smith", "JS");
            ICommentAuthor john = pres.getCommentAuthors().addAuthor("John Doe", "JD");
            jane.getComments().addComment("From Jane", pres.getSlides().get(0),
                    new PointF(1.0f, 1.0f), LocalDateTime.now());
            john.getComments().addComment("From John", pres.getSlides().get(0),
                    new PointF(3.0f, 3.0f), LocalDateTime.now());
            pres.save(withComments.toString(), SaveFormat.PPTX);
        }

        Path removed = tempDir.resolve("one-author-removed.pptx");
        try (var pres = new Presentation(withComments.toString())) {
            pres.getCommentAuthors().removeAt(0);
            pres.save(removed.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(removed)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            assertThat(PackageAssertions.selectNodes(pkg, "ppt/commentAuthors.xml", "//p:cmAuthor")
                    .stream().map(e -> e.getAttribute("name")).toList())
                    .as("authors in ppt/commentAuthors.xml of %s", removed)
                    .containsExactly("John Doe");
            assertThat(pkg.entryNames().stream().filter(n -> n.startsWith("ppt/comments/"))
                    .flatMap(part -> {
                        try {
                            return PackageAssertions.selectNodes(pkg, part, "//p:cm/p:text")
                                    .stream().map(e -> e.getTextContent());
                        } catch (IOException e) {
                            throw new java.io.UncheckedIOException(e);
                        }
                    }).toList())
                    .as("comments left in %s", removed)
                    .containsExactly("From John");
        }
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
