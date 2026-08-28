package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.IComment;
import org.aspose.slides.foss.ICommentAuthor;
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
 * A reply must be a reply in the file, not only in memory.
 *
 * <p>Setting a parent comment reads back correctly from the object model and used to reach no
 * part of the package, so a reviewer's threaded discussion opened as a row of unrelated
 * comments. Threading is written in two places, and both are needed:</p>
 *
 * <ul>
 *   <li>{@code ppt/threadedComments/threadedCommentN.xml}, content type
 *       {@code application/vnd.ms-powerpoint.threadedcomments+xml}, whose {@code p188:cm}
 *       entries carry {@code @parentId} — the modern list; and</li>
 *   <li>a {@code p15:threadingInfo} extension on the classic {@code <p:cm>} element in
 *       {@code ppt/comments/commentN.xml} — which is where a reader that renders comments
 *       looks for the thread. A package that carries only the modern part opens with the
 *       reply shown as a second, unrelated comment: the file is valid and the thread is
 *       gone.</li>
 * </ul>
 *
 * <p>The classic {@code <p:cm>} element has no <em>attribute</em> for a parent, and inventing
 * one produces content no reader understands; the extension list is the schema's declared
 * place for markup like this.</p>
 *
 * <p>These tests state that a reply survives the save. If the decision is made instead to drop
 * reply support rather than implement it, the public API that accepts a parent comment goes at
 * the same time as these tests — what must not survive is an API that accepts a reply and
 * discards it.</p>
 */
class CommentThreadConformanceTest {

    /** The content type of the modern threaded-comments part. */
    private static final String THREADED_COMMENTS_CONTENT_TYPE =
            "application/vnd.ms-powerpoint.threadedcomments+xml";

    /** The {@code p:ext/@uri} that identifies a {@code <p15:threadingInfo>} extension. */
    private static final String THREADING_INFO_URI = "{C676402C-5697-4E1C-873F-D02D1690AC5C}";

    /** The XPath selecting a threading extension wherever it appears. */
    private static final String THREADING_INFO =
            "//p:cm/p:extLst/p:ext[@uri='" + THREADING_INFO_URI + "']/p15:threadingInfo";

    @TempDir
    Path tempDir;

    /** A reply must be recorded in the package as a reply to its parent. */
    @Test
    void aReplyMustBeRecordedInTheFileAsAReply() throws Exception {
        Path out = writeCommentAndReply();

        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<String> threadedParts = pkg.entryNames().stream()
                    .filter(name -> name.startsWith("ppt/threadedComments/"))
                    .toList();
            assertThat(threadedParts)
                    .as("threaded-comment parts in %s; parts present are %s",
                            out, pkg.entryNames())
                    .isNotEmpty();

            assertThat(PackageAssertions.contentTypeOverrides(pkg).values())
                    .as("content type declared for the threaded-comment part in %s", out)
                    .contains(THREADED_COMMENTS_CONTENT_TYPE);

            assertThat(PackageAssertions.selectNodes(
                    pkg, threadedParts.get(0), "//p188:cm[@parentId]"))
                    .as("threaded comments carrying a parent id in %s", out)
                    .isNotEmpty();
        }
    }

    /**
     * The classic comment list must carry the thread too, on the reply and only on the reply.
     *
     * <p>{@code p15:parentCm} names the parent by author and index rather than by a single id,
     * because a comment index is only unique within one author.</p>
     */
    @Test
    void theClassicCommentListMustCarryTheThreadOnTheReply() throws Exception {
        Path out = writeCommentAndReply();

        try (PptxPackage pkg = PptxPackage.open(out)) {
            String part = "ppt/comments/comment1.xml";
            List<Element> comments =
                    PackageAssertions.selectNodes(pkg, part, "//p:cmLst/p:cm");
            assertThat(comments)
                    .as("classic comments in %s of %s%n%s", part, out, pkg.text(part))
                    .hasSize(2);

            Element parent = comments.get(0);
            Element reply = comments.get(1);

            assertThat(PackageAssertions.selectNodes(pkg, part,
                    "//p:cm[@idx='" + parent.getAttribute("idx") + "']"
                            + "/p:extLst/p:ext/p15:threadingInfo"))
                    .as("the comment that starts the thread must not claim a parent, in %s of %s"
                            + "%n%s", part, out, pkg.text(part))
                    .isEmpty();

            List<Element> parentRefs = PackageAssertions.selectNodes(pkg, part,
                    THREADING_INFO + "/p15:parentCm");
            assertThat(parentRefs)
                    .as("p15:parentCm on the reply in %s of %s%n%s", part, out, pkg.text(part))
                    .hasSize(1);
            assertThat(parentRefs.get(0).getAttribute("authorId"))
                    .as("p15:parentCm/@authorId in %s of %s", part, out)
                    .isEqualTo(parent.getAttribute("authorId"));
            assertThat(parentRefs.get(0).getAttribute("idx"))
                    .as("p15:parentCm/@idx in %s of %s", part, out)
                    .isEqualTo(parent.getAttribute("idx"));

            assertThat(PackageAssertions.childNames(reply))
                    .as("children of the replying <p:cm> in %s of %s", part, out)
                    .containsExactly("p:pos", "p:text", "p:extLst");
        }
    }

    /**
     * A thread must survive being read back and written out again.
     *
     * <p>Asserted on the second package, not on the object model: a reader that drops the
     * thread and a writer that never wrote it are indistinguishable from inside the library.</p>
     */
    @Test
    void aThreadMustSurviveALoadAndASaveUnchanged() throws Exception {
        Path first = writeCommentAndReply();
        Path second = tempDir.resolve("comment-reply-resaved.pptx");
        try (var pres = new Presentation(first.toString())) {
            pres.save(second.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(second)) {
            assertThat(PackageAssertions.selectNodes(pkg, "ppt/comments/comment1.xml",
                    THREADING_INFO + "/p15:parentCm"))
                    .as("p15:parentCm after a load and a save, in %s%n%s",
                            second, pkg.text("ppt/comments/comment1.xml"))
                    .hasSize(1);
            assertThat(PackageAssertions.selectNodes(pkg,
                    "ppt/threadedComments/threadedComment1.xml", "//p188:cm[@parentId]"))
                    .as("p188:cm/@parentId after a load and a save, in %s", second)
                    .hasSize(1);
        }
    }

    /** No comment may carry an attribute the standard does not define for it. */
    @Test
    void aCommentMustNotCarryAnAttributeTheStandardDoesNotDefine() throws Exception {
        Path out = writeCommentAndReply();

        try (PptxPackage pkg = PptxPackage.open(out)) {
            for (String partName : pkg.entryNames()) {
                if (!partName.startsWith("ppt/comments/")) {
                    continue;
                }
                assertThat(pkg.text(partName))
                        .as("part %s of %s", partName, out)
                        .doesNotContain("parentCmId");
            }
        }
    }

    /** Comments that are not replies must not gain an empty extension list. */
    @Test
    void unrelatedCommentsMustNotCarryAThreadingExtension() throws Exception {
        Path out = tempDir.resolve("comments-unrelated.pptx");
        try (var pres = new Presentation()) {
            ICommentAuthor author = pres.getCommentAuthors().addAuthor("Jane Smith", "JS");
            ISlide slide = pres.getSlides().get(0);
            author.getComments().addComment("First", slide,
                    new PointF(1.0f, 1.0f), LocalDateTime.now());
            author.getComments().addComment("Second", slide,
                    new PointF(2.0f, 2.0f), LocalDateTime.now());
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            String part = "ppt/comments/comment1.xml";
            assertThat(PackageAssertions.selectNodes(pkg, part, "//p:cm/p:extLst"))
                    .as("extension lists on unrelated comments in %s of %s%n%s",
                            part, out, pkg.text(part))
                    .isEmpty();
        }
    }

    private Path writeCommentAndReply() throws IOException {
        Path out = tempDir.resolve("comment-reply.pptx");
        try (var pres = new Presentation()) {
            ICommentAuthor author = pres.getCommentAuthors().addAuthor("Jane Smith", "JS");
            ISlide slide = pres.getSlides().get(0);
            IComment parent = author.getComments().addComment("Review this slide", slide,
                    new PointF(2.0f, 2.0f), LocalDateTime.now());
            IComment reply = author.getComments().addComment("Agreed, will do", slide,
                    new PointF(2.5f, 2.5f), LocalDateTime.now());
            reply.setParentComment(parent);
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        return out;
    }
}
