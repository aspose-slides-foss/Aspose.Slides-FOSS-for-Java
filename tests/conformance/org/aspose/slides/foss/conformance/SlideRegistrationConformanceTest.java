package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.IAutoShape;
import org.aspose.slides.foss.ILayoutSlide;
import org.aspose.slides.foss.ISlide;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.ShapeType;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A slide that was added must be in the saved file, and a slide that was removed must not.
 *
 * <p>Slides are counted from {@code ppt/presentation.xml}'s {@code p:sldIdLst}, never from the
 * number of {@code ppt/slides/slideN.xml} entries. A slide part that no {@code p:sldId}
 * references, that no relationship points at and that no content-type override types is not a
 * slide: PowerPoint, Apache POI and every other reader ignore it. Counting parts instead is
 * what lets a deck that has silently lost its content look complete.</p>
 */
class SlideRegistrationConformanceTest {

    @TempDir
    Path tempDir;

    /** A slide added through the API must appear in the saved file. */
    @Test
    void anAddedSlideMustBeRegisteredInTheSavedFile() throws Exception {
        Path out = tempDir.resolve("added.pptx");
        try (var pres = new Presentation()) {
            ILayoutSlide layout = pres.getLayoutSlides().get(0);
            pres.getSlides().addEmptySlide(layout);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertRegisteredSlideCount(pkg, 2);
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
        }
        assertThat(ThirdPartyReadBack.slideCount(out))
                .as("slides an independent reader finds in %s", out)
                .isEqualTo(2);
    }

    /** A slide inserted at a position must appear in the saved file. */
    @Test
    void anInsertedSlideMustBeRegisteredInTheSavedFile() throws Exception {
        Path out = tempDir.resolve("inserted.pptx");
        try (var pres = new Presentation()) {
            ILayoutSlide layout = pres.getLayoutSlides().get(0);
            pres.getSlides().insertEmptySlide(1, layout);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertRegisteredSlideCount(pkg, 2);
        }
    }

    /** A cloned slide must carry the shapes of the slide it was cloned from. */
    @Test
    void aClonedSlideMustCarryTheShapesOfItsSource() throws Exception {
        Path out = tempDir.resolve("cloned.pptx");
        try (var pres = new Presentation()) {
            ISlide source = pres.getSlides().get(0);
            IAutoShape shape = source.getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 200, 100);
            shape.addTextFrame("Cloned content");
            pres.getSlides().addClone(source);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertRegisteredSlideCount(pkg, 2);
        }
        assertThat(ThirdPartyReadBack.textOf(out, 1))
                .as("text an independent reader finds on the cloned slide of %s", out)
                .anySatisfy(text -> assertThat(text).contains("Cloned content"));
    }

    /** Removing a slide must remove it from the file, not only from the in-memory list. */
    @Test
    void aRemovedSlideMustBeGoneFromTheSavedFile() throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "three.pptx", "First", "Second", "Third");
        Path out = tempDir.resolve("removed.pptx");
        try (var pres = new Presentation(source.toString())) {
            pres.getSlides().removeAt(1);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertRegisteredSlideCount(pkg, 2);
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
        }
        assertThat(ThirdPartyReadBack.slideCount(out))
                .as("slides an independent reader finds after removing one from %s", out)
                .isEqualTo(2);
    }

    /** Removing a slide by reference must remove it from the file. */
    @Test
    void aSlideRemovedByReferenceMustBeGoneFromTheSavedFile() throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "three.pptx", "First", "Second", "Third");
        Path out = tempDir.resolve("removed-by-ref.pptx");
        try (var pres = new Presentation(source.toString())) {
            pres.getSlides().remove(pres.getSlides().get(2));
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertRegisteredSlideCount(pkg, 2);
        }
    }
}
