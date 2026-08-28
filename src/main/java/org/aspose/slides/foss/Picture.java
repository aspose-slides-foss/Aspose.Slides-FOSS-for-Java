package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.OpcPackage;
import org.aspose.slides.foss.internal.pptx.RelsHelper;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.List;

/**
 * Represents a picture in a presentation.
 *
 * <p>Wraps an OOXML {@code <a:blip>} element for reading and writing
 * embedded image references and linked image URLs.</p>
 */
public final class Picture implements ISlidesPicture, ISlideComponent, IPresentationComponent {

    private static final String NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    private static final String REL_TYPE_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    private Element blip;
    private IBaseSlide parentSlide;
    private IPresentation parentPresentation;
    private IPPImage cachedImage;
    private OpcPackage opcPackage;
    private String slidePartName;

    /**
     * Package-private constructor. Use {@link #initInternal} to initialise.
     */
    Picture() {
    }

    /**
     * Internal initialisation with the {@code <a:blip>} XML element.
     *
     * @param blipElement      the {@code <a:blip>} element containing {@code r:embed} reference
     * @param parentSlide      the parent slide object
     * @param parentPresentation the parent presentation, or {@code null}
     */
    public void initInternal(Element blipElement, IBaseSlide parentSlide, IPresentation parentPresentation) {
        this.blip = blipElement;
        this.parentSlide = parentSlide;
        this.parentPresentation = parentPresentation;
    }

    /**
     * Internal initialisation with the {@code <a:blip>} XML element (no presentation context).
     *
     * @param blipElement the {@code <a:blip>} element containing {@code r:embed} reference
     * @param parentSlide the parent slide object
     */
    public void initInternal(Element blipElement, IBaseSlide parentSlide) {
        initInternal(blipElement, parentSlide, null);
    }

    /**
     * Internal initialisation with the {@code <a:blip>} XML element and OPC package context.
     *
     * <p>This overload provides full relationship management support through the
     * OPC package and slide part name, enabling {@link #setBlipImage(Element, OpcPackage, String, IPPImage)}
     * to create proper OPC relationships.</p>
     *
     * @param blipElement        the {@code <a:blip>} element containing {@code r:embed} reference
     * @param opcPackage         the OPC package for relationship resolution
     * @param slidePartName      the part name of the parent slide (e.g. {@code "ppt/slides/slide1.xml"})
     * @param parentSlide        the parent slide object
     */
    public void initInternal(Element blipElement, OpcPackage opcPackage, String slidePartName, IBaseSlide parentSlide) {
        this.blip = blipElement;
        this.opcPackage = opcPackage;
        this.slidePartName = slidePartName;
        this.parentSlide = parentSlide;
    }

    @Override
    public IPresentationComponent asIPresentationComponent() {
        return this;
    }

    @Override
    public IPPImage getImage() {
        if (cachedImage != null) {
            return cachedImage;
        }
        String embedId = blip.getAttributeNS(NS_R, "embed");
        if (embedId == null || embedId.isEmpty()) {
            return null;
        }
        // Resolve through the slide's relationship and presentation image collection
        // when full relationship manager infrastructure is available
        return null;
    }

    /**
     * {@inheritDoc}
     *
     * @throws IllegalArgumentException if {@code value} is {@code null}
     * @throws IllegalStateException    if this picture is not bound to a package part, so that
     *                                  no relationship can be declared for the image
     */
    @Override
    public void setImage(IPPImage value) {
        if (value == null) {
            throw new IllegalArgumentException("Image must not be null");
        }
        if (opcPackage == null || slidePartName == null) {
            // An r:embed names a relationship in the owning part's .rels. With no part
            // there is none to declare, and the two things that could be done instead are
            // both worse than refusing: an invented id resolves to nothing, and accepting
            // the image silently discards it -- and removes whatever reference was there.
            throw new IllegalStateException(
                    "This picture is not attached to a package part, so an image set on it "
                            + "could not be given a relationship to refer to. Set the image "
                            + "through a shape, a fill or a bullet that belongs to a slide, "
                            + "layout or master.");
        }
        this.cachedImage = value;
        setBlipImage(blip, opcPackage, slidePartName, value);
    }

    @Override
    public String getLinkPathLong() {
        String linkId = blip.getAttributeNS(NS_R, "link");
        if (linkId == null || linkId.isEmpty()) {
            return "";
        }
        // Full resolution would require relationship manager lookup.
        // Returns the relationship ID as a placeholder when no rels manager is available.
        return linkId;
    }

    @Override
    public void setLinkPathLong(String value) {
        if (value == null || value.isEmpty()) {
            // Remove the link attribute
            if (blip.hasAttributeNS(NS_R, "link")) {
                blip.removeAttributeNS(NS_R, "link");
            }
        } else {
            blip.setAttributeNS(NS_R, "r:link", value);
        }
    }

    @Override
    public IPresentation getPresentation() {
        return parentPresentation;
    }

    @Override
    public IBaseSlide getSlide() {
        return parentSlide;
    }

    @Override
    public ISlideComponent asISlideComponent() {
        return this;
    }

    /**
     * Returns the underlying {@code <a:blip>} XML element.
     *
     * @return the blip element
     */
    Element getBlipElement() {
        return blip;
    }

    /**
     * Sets the image relationship on a blip element using OPC relationship management.
     *
     * <p>Looks up the slide part's existing image relationships. If a relationship
     * already targets the given image's part name, it reuses that relationship ID.
     * Otherwise, a new relationship is created and saved.</p>
     *
     * @param blipEl        the {@code <a:blip>} element
     * @param opcPackage    the OPC package containing the slide
     * @param slidePartName the part name of the slide (e.g. {@code "ppt/slides/slide1.xml"})
     * @param ppImage       the presentation image to reference
     */
    public static void setBlipImage(Element blipEl, OpcPackage opcPackage,
                                     String slidePartName, IPPImage ppImage) {
        String slideDir = slidePartName.contains("/")
                ? slidePartName.substring(0, slidePartName.lastIndexOf('/'))
                : "";
        String imagePath = ((PPImage) ppImage).getPartName();
        String relativeTarget = computeRelativePath(slideDir, imagePath);

        var rels = new RelsHelper(opcPackage, slidePartName);
        List<RelsHelper.RelEntry> imageRels = rels.getAllRelationships().stream()
                .filter(r -> REL_TYPE_IMAGE.equals(r.type()))
                .toList();

        String embedId = null;
        for (RelsHelper.RelEntry rel : imageRels) {
            String resolved = org.aspose.slides.foss.internal.pptx.SlidePart.resolveTargetStatic(
                    slidePartName, rel.target());
            if (resolved.equals(imagePath)) {
                embedId = rel.id();
                break;
            }
        }

        if (embedId == null) {
            embedId = rels.addRelationship(REL_TYPE_IMAGE, relativeTarget);
            rels.save();
        }

        blipEl.setAttributeNS(NS_R, "r:embed", embedId);
        opcPackage.serializeXml(slidePartName, blipEl.getOwnerDocument());
    }

    /**
     * Computes a relative path from a directory to a target path.
     *
     * <p>Both paths use forward-slash ({@code /}) separators. The method finds
     * the longest common prefix and builds a relative path with {@code ..} segments
     * as needed.</p>
     *
     * @param fromDir the source directory (e.g. {@code "ppt/slides"})
     * @param toPath  the target path (e.g. {@code "ppt/media/image1.png"})
     * @return the relative path from {@code fromDir} to {@code toPath}
     */
    public static String computeRelativePath(String fromDir, String toPath) {
        String[] fromParts = fromDir.isEmpty() ? new String[0] : fromDir.split("/");
        String[] toParts = toPath.split("/");

        int common = 0;
        for (int i = 0; i < Math.min(fromParts.length, toParts.length); i++) {
            if (fromParts[i].equals(toParts[i])) {
                common++;
            } else {
                break;
            }
        }

        int up = fromParts.length - common;
        var sb = new StringBuilder();
        for (int i = 0; i < up; i++) {
            if (!sb.isEmpty()) {
                sb.append('/');
            }
            sb.append("..");
        }
        for (int i = common; i < toParts.length; i++) {
            if (!sb.isEmpty()) {
                sb.append('/');
            }
            sb.append(toParts[i]);
        }
        return sb.toString();
    }
}
