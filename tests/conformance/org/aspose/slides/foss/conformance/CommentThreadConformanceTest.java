package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.IComment;
import org.aspose.slides.foss.ICommentAuthor;
import org.aspose.slides.foss.ISlide;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.drawing.PointF;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A reply must be a reply in the file, not only in memory.
 *
 * <p>Setting a parent comment reads back correctly from the object model and reaches no part of
 * the package, so a reviewer's threaded discussion opens in PowerPoint as a row of unrelated
 * comments. Threading is a part of its own — {@code ppt/threadedComments/threadedCommentN.xml},
 * content type {@code application/vnd.ms-powerpoint.threadedcomments+xml} — whose entries carry
 * {@code @parentId}; the classic {@code <p:cm>} element has no attribute for it, and inventing
 * one produces content no reader understands.</p>
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
