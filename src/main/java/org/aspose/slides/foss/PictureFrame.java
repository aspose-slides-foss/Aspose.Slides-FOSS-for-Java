package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.util.Optional;

/**
 * Represents a frame with a picture inside.
 *
 * <p>A picture frame wraps an OOXML {@code <p:pic>} element and provides access to
 * the picture fill format, picture frame locks, and relative scale properties.</p>
 */
public final class PictureFrame extends GeometryShape implements IPictureFrame {

    private double relativeScaleHeight = 1.0;
    private double relativeScaleWidth = 1.0;

    /**
     * Creates a PictureFrame backed by the given XML element.
     *
     * @param xmlElement   the {@code <p:pic>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public PictureFrame(Element xmlElement, Runnable saveCallback) {
        super(xmlElement, saveCallback);
    }

    /**
     * Creates a PictureFrame with no backing element.
     */
    public PictureFrame() {
        super();
    }

    // --- Shape type (overrides GeometryShape to default to RECTANGLE) ---

    /**
     * {@inheritDoc}
     *
     * <p>Returns or sets the AutoShape type for a PictureFrame. All items of
     * {@link ShapeType} are allowable except line and connector types. Defaults
     * to {@link ShapeType#RECTANGLE} when no preset geometry is defined.</p>
     */
    @Override
    public ShapeType getShapeType() {
        if (xmlElement == null) return ShapeType.RECTANGLE;
        Element spPr = findChild(xmlElement, NS_P, "spPr");
        if (spPr == null) return ShapeType.RECTANGLE;
        Element prstGeom = findChild(spPr, NS_A, "prstGeom");
        if (prstGeom == null) return ShapeType.RECTANGLE;
        String prst = prstGeom.getAttribute("prst");
        if (prst == null || prst.isEmpty()) return ShapeType.RECTANGLE;
        return ShapeType.fromOoxml(prst).orElse(ShapeType.RECTANGLE);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setShapeType(ShapeType value) {
        if (xmlElement == null) return;
        Optional<String> prst = value.toOoxml();
        String prstName = prst.orElse("rect");
        Element spPr = findChild(xmlElement, NS_P, "spPr");
        if (spPr == null) return;
        Element prstGeom = findChild(spPr, NS_A, "prstGeom");
        if (prstGeom == null) {
            Document doc = spPr.getOwnerDocument();
            prstGeom = doc.createElementNS(NS_A, "a:prstGeom");
            spPr.appendChild(prstGeom);
        }
        prstGeom.setAttribute("prst", prstName);
        if (saveCallback != null) saveCallback.run();
    }

    // --- Picture frame lock ---

    /**
     * {@inheritDoc}
     *
     * <p>Navigates {@code <p:nvPicPr>/<p:cNvPicPr>/<a:picLocks>} and creates
     * the {@code <a:picLocks>} element if it does not exist.</p>
     */
    @Override
    public IPictureFrameLock getPictureFrameLock() {
        if (xmlElement == null) return null;
        Element nvPicPr = findChild(xmlElement, NS_P, "nvPicPr");
        Element picLocks = null;
        if (nvPicPr != null) {
            Element cNvPicPr = findChild(nvPicPr, NS_P, "cNvPicPr");
            if (cNvPicPr != null) {
                picLocks = findChild(cNvPicPr, NS_A, "picLocks");
                if (picLocks == null) {
                    picLocks = ensureChild(cNvPicPr, NS_A, "picLocks", "a:picLocks");
                }
            }
        }
        PictureFrameLock lock = new PictureFrameLock();
        lock.initInternal(picLocks, saveCallback);
        return lock;
    }

    // --- Picture format ---

    /**
     * {@inheritDoc}
     *
     * <p>Returns the {@link PictureFillFormat} wrapping the {@code <p:blipFill>}
     * child element, or {@code null} if no blipFill is present.</p>
     */
    @Override
    public IPictureFillFormat getPictureFormat() {
        if (xmlElement == null) return null;
        Element blipFill = findChild(xmlElement, NS_P, "blipFill");
        if (blipFill == null) return null;
        var pictureFill = new PictureFillFormat(blipFill, saveCallback);
        ShapeCollection shapes = owningShapes();
        if (shapes != null) {
            pictureFill.setPartContext(shapes.getPartPackage(), shapes.getPartName());
        }
        return pictureFill;
    }

    // --- Relative scale ---

    @Override
    public double getRelativeScaleHeight() {
        return relativeScaleHeight;
    }

    @Override
    public void setRelativeScaleHeight(double value) {
        this.relativeScaleHeight = value;
    }

    @Override
    public double getRelativeScaleWidth() {
        return relativeScaleWidth;
    }

    @Override
    public void setRelativeScaleWidth(double value) {
        this.relativeScaleWidth = value;
    }

    // --- Cameo ---

    @Override
    public boolean isCameo() {
        return false;
    }

    // --- cNvPr override for picture elements ---

    /**
     * Returns the {@code cNvPr} element from {@code <p:nvPicPr>/<p:cNvPr>}.
     *
     * @return the cNvPr element, or falls back to super implementation
     */
    @Override
    public Element getCNvPr() {
        if (xmlElement == null) return null;
        Element nvPicPr = findChild(xmlElement, NS_P, "nvPicPr");
        if (nvPicPr != null) {
            Element cNvPr = findChild(nvPicPr, NS_P, "cNvPr");
            if (cNvPr != null) return cNvPr;
        }
        return super.getCNvPr();
    }
}
