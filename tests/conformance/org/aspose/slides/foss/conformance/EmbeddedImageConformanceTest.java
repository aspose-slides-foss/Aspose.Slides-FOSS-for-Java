package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.IPPImage;
import org.aspose.slides.foss.ISlide;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.ShapeType;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Element;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * An embedded picture must be reachable from the slide that shows it.
 *
 * <p>The image bytes being in the package is not enough. A slide refers to an image through
 * {@code <a:blip r:embed="rIdN"/>}, and {@code rIdN} has to be declared in that slide's own
 * {@code .rels}. If it is not, the picture frame is an empty box: readers either raise on the
 * unknown id or draw nothing, and the user sees a blank rectangle where the photo was.</p>
 */
class EmbeddedImageConformanceTest {

    private static final Path PNG = Fixtures.testData("lotus.png");
    private static final Path JPEG = Fixtures.testData("image_example.jpg");

    @TempDir
    Path tempDir;

    /** A picture added from bytes must resolve its relationship. */
    @Test
    void aPictureAddedFromBytesMustResolveItsRelationship() throws Exception {
        Path out = writeDeckWithPicture("from-bytes.pptx", Files.readAllBytes(PNG));
        assertPictureIsReachable(out, "png");
    }

    /** A picture added from a stream must resolve its relationship. */
    @Test
    void aPictureAddedFromAStreamMustResolveItsRelationship() throws Exception {
        Path out = tempDir.resolve("from-stream.pptx");
        try (var pres = new Presentation()) {
            IPPImage image;
            try (var in = new ByteArrayInputStream(Files.readAllBytes(JPEG))) {
                image = pres.getImages().addImage(in);
            }
            ISlide slide = pres.getSlides().get(0);
            slide.getShapes().addPictureFrame(ShapeType.RECTANGLE, 50, 50, 200, 150, image);
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        assertPictureIsReachable(out, "jpeg", "jpg");
    }

    /** No slide may refer to a relationship id that its own .rels does not declare. */
    @Test
    void noSlideMayReferToARelationshipItDoesNotDeclare() throws Exception {
        Path out = writeDeckWithPicture("dangling.pptx", Files.readAllBytes(PNG));
        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertEveryRelationshipReferenceResolves(pkg);
        }
    }

    /** An independent reader must find the picture and be able to get at its bytes. */
    @Test
    void anIndependentReaderMustBeAbleToGetAtTheImageBehindThePicture() throws Exception {
        Path out = writeDeckWithPicture("readable.pptx", Files.readAllBytes(PNG));
        ThirdPartyReadBack.assertOpens(out);
        ThirdPartyReadBack.assertPicturesResolve(out, 0);
    }

    private Path writeDeckWithPicture(String fileName, byte[] imageBytes) throws IOException {
        Path out = tempDir.resolve(fileName);
        try (var pres = new Presentation()) {
            IPPImage image = pres.getImages().addImage(imageBytes);
            ISlide slide = pres.getSlides().get(0);
            slide.getShapes().addPictureFrame(ShapeType.RECTANGLE, 50, 50, 200, 150, image);
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        return out;
    }

    private static void assertPictureIsReachable(Path pptx, String... acceptedExtensions)
            throws IOException {
        try (PptxPackage pkg = PptxPackage.open(pptx)) {
            List<Element> blips = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:pic/p:blipFill/a:blip");
            assertThat(blips).as("a:blip elements in slide 1 of %s", pptx).hasSize(1);

            String embedId = blips.get(0).getAttributeNS(PptxPackage.NS_REL, "embed");
            assertThat(embedId).as("r:embed of the blip in %s", pptx).isNotEmpty();

            List<Element> relationships = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/_rels/slide1.xml.rels",
                    "//rel:Relationship[@Id='" + embedId + "']");
            assertThat(relationships)
                    .as("relationship '%s' declared in ppt/slides/_rels/slide1.xml.rels of %s%n%s",
                            embedId, pptx, pkg.text("ppt/slides/_rels/slide1.xml.rels"))
                    .hasSize(1);

            String target = PptxPackage.resolveTarget(
                    "ppt/slides/slide1.xml", relationships.get(0).getAttribute("Target"));
            assertThat(pkg.hasPart(target))
                    .as("image part %s referenced from slide 1 of %s", target, pptx)
                    .isTrue();
            assertThat(PackageAssertions.contentTypeDefaults(pkg).keySet())
                    .as("content type registered for the image extension in %s", pptx)
                    .containsAnyElementsOf(List.of(acceptedExtensions));
        }
    }
}
