package org.aspose.slides.foss;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilderFactory;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests for {@link Picture}.
 *
 * <p>Verifies picture fill behavior and image persistence.</p>
 */
class PictureTest {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private Document doc;

    @BeforeEach
    void setUp() throws Exception {
        doc = DocumentBuilderFactory.newInstance()
                .newDocumentBuilder()
                .newDocument();
    }

    private Element createBlip() {
        Element blip = doc.createElementNS(NS_A, "a:blip");
        doc.appendChild(blip);
        return blip;
    }

    @Test
    void asIPresentationComponent_returnsSelf() {
        var picture = new Picture();
        picture.initInternal(createBlip(), null);

        assertThat(picture.asIPresentationComponent()).isSameAs(picture);
    }

    @Test
    void asISlideComponent_returnsSelf() {
        var picture = new Picture();
        picture.initInternal(createBlip(), null);

        assertThat(picture.asISlideComponent()).isSameAs(picture);
    }

    @Test
    void image_returnsNullWhenNoEmbedAttribute() {
        var picture = new Picture();
        picture.initInternal(createBlip(), null);

        assertThat(picture.getImage()).isNull();
    }

    /**
     * A picture that is not bound to a package part keeps the image but writes no
     * {@code r:embed}.
     *
     * <p>An {@code r:embed} names a relationship in the owning part's {@code .rels}. With no
     * part there is no such relationship, and inventing an id produces a reference that
     * resolves to nothing — a picture that readers show as an empty box. The relationship is
     * written when the picture belongs to a part; that path is asserted against the produced
     * package in the conformance tests.</p>
     */
    @Test
    void setImage_withoutAPartKeepsTheImageAndInventsNoRelationship() {
        Element blip = createBlip();
        var picture = new Picture();
        picture.initInternal(blip, null);

        var stubImage = new StubPPImage();
        picture.setImage(stubImage);

        assertThat(picture.getImage()).isSameAs(stubImage);
        assertThat(blip.hasAttributeNS(NS_R, "embed")).isFalse();
    }

    @Test
    void setImage_throwsOnNull() {
        var picture = new Picture();
        picture.initInternal(createBlip(), null);

        assertThatThrownBy(() -> picture.setImage(null))
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void linkPathLong_returnsEmptyStringWhenNoLinkAttribute() {
        var picture = new Picture();
        picture.initInternal(createBlip(), null);

        assertThat(picture.getLinkPathLong()).isEmpty();
    }

    @Test
    void setLinkPathLong_setsLinkAttribute() {
        Element blip = createBlip();
        var picture = new Picture();
        picture.initInternal(blip, null);

        picture.setLinkPathLong("https://example.com/image.png");

        assertThat(blip.getAttributeNS(NS_R, "link")).isNotEmpty();
    }

    @Test
    void setLinkPathLong_emptyStringRemovesAttribute() {
        Element blip = createBlip();
        var picture = new Picture();
        picture.initInternal(blip, null);

        picture.setLinkPathLong("https://example.com/image.png");
        assertThat(blip.hasAttributeNS(NS_R, "link")).isTrue();

        picture.setLinkPathLong("");
        assertThat(blip.hasAttributeNS(NS_R, "link")).isFalse();
    }

    @Test
    void presentation_returnsNullWithoutParentSlide() {
        var picture = new Picture();
        picture.initInternal(createBlip(), null);

        assertThat(picture.getPresentation()).isNull();
    }

    @Test
    void slide_returnsParentSlide() {
        var picture = new Picture();
        var stubSlide = new BaseSlide();
        picture.initInternal(createBlip(), stubSlide);

        assertThat(picture.getSlide()).isSameAs(stubSlide);
    }

    /**
     * picture fill with an image persists.
     *
     * <p>Tests that setting fill type to PICTURE and configuring a picture
     * via PictureFillFormat persists the fill type when re-read from the same XML.</p>
     */
    @Test
    void pictureFill_withImage_persistsFillType() throws Exception {
        // Create a fresh document with <p:spPr> as root
        var freshDoc = DocumentBuilderFactory.newInstance()
                .newDocumentBuilder()
                .newDocument();
        Element spPr = freshDoc.createElementNS(
                "http://schemas.openxmlformats.org/presentationml/2006/main", "p:spPr");
        freshDoc.appendChild(spPr);

        // Set up picture fill
        var ff = new FillFormat(spPr, null);
        ff.setFillType(FillType.PICTURE);
        IPictureFillFormat pff = ff.getPictureFillFormat();
        assertThat(pff.getPictureFillMode()).isEqualTo(PictureFillMode.STRETCH);

        // Set the picture's image
        ISlidesPicture picture = pff.getPicture();
        picture.setImage(new StubPPImage());

        // Re-read from same XML (simulates save/reload)
        var ff2 = new FillFormat(spPr, null);
        assertThat(ff2.getFillType()).isEqualTo(FillType.PICTURE);
    }

    /**
     * Stub implementation of IPPImage for testing.
     */
    private static class StubPPImage implements IPPImage {
        @Override
        public byte[] getBinaryData() { return new byte[0]; }

        @Override
        public IImage getImage() { return null; }

        @Override
        public String getContentType() { return "image/png"; }

        @Override
        public int getWidth() { return 100; }

        @Override
        public int getHeight() { return 100; }

        @Override
        public int getX() { return 0; }

        @Override
        public int getY() { return 0; }

        @Override
        public void replaceImage(byte[] newImageData) { }

        @Override
        public void replaceImage(IImage newImage) { }

        @Override
        public void replaceImage(IPPImage newImage) { }
    }
}
