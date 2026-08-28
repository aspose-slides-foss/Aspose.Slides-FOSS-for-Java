package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.List;
import java.util.Optional;

/**
 * Base class for shapes with geometry, backed by an OOXML shape element.
 *
 * <p>Provides position/size (from {@code <a:xfrm>}), shape type (from {@code <a:prstGeom>}),
 * and adjustment values.</p>
 */
public class GeometryShape implements IGeometryShape {

    /** PresentationML namespace. */
    protected static final String NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    /** DrawingML namespace. */
    protected static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    /** EMUs per point. */
    protected static final double EMU_PER_POINT = 12700.0;

    /** The backing XML element (e.g. {@code <p:sp>} or {@code <p:cxnSp>}). */
    protected Element xmlElement;
    /** Callback to persist changes. */
    protected Runnable saveCallback;
    /** Sibling shapes for connection lookups. */
    protected Iterable<IShape> parentShapes;

    /**
     * Creates a GeometryShape backed by the given XML element.
     *
     * @param xmlElement   the shape XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public GeometryShape(Element xmlElement, Runnable saveCallback) {
        this.xmlElement = xmlElement;
        this.saveCallback = saveCallback;
        this.parentShapes = List.of();
    }

    /**
     * Creates a GeometryShape with no backing element.
     */
    public GeometryShape() {
        this(null, null);
    }

    /**
     * Sets the sibling shapes collection used for connection lookups.
     *
     * @param shapes the sibling shapes
     */
    public void setParentShapes(Iterable<IShape> shapes) {
        this.parentShapes = shapes != null ? shapes : List.of();
    }

    /**
     * Tells a fill format which package part this shape is serialized into, so that a
     * picture fill can declare its image relationship in the right {@code .rels}.
     *
     * @param fill the fill format to inform
     * @param <T>  the fill format type
     * @return the same fill format
     */
    <T extends FillFormat> T withPartContext(T fill) {
        if (parentShapes instanceof ShapeCollection shapes) {
            fill.setPartContext(shapes.getPartPackage(), shapes.getPartName());
        }
        return fill;
    }

    /**
     * @return the shape collection this shape belongs to, or {@code null} if it belongs to none
     */
    ShapeCollection owningShapes() {
        return parentShapes instanceof ShapeCollection shapes ? shapes : null;
    }

    // --- XML helpers ---

    /**
     * Finds a direct child element by namespace and local name.
     */
    protected Element findChild(Element parent, String nsUri, String localName) {
        if (parent == null) return null;
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el) {
                if (nsUri.equals(el.getNamespaceURI()) && localName.equals(el.getLocalName())) {
                    return el;
                }
            }
        }
        return null;
    }

    /**
     * Finds or creates a direct child element.
     */
    protected Element ensureChild(Element parent, String nsUri, String localName, String qualifiedName) {
        Element child = findChild(parent, nsUri, localName);
        if (child == null) {
            Document doc = parent.getOwnerDocument();
            child = doc.createElementNS(nsUri, qualifiedName);
            parent.appendChild(child);
        }
        return child;
    }

    // --- xfrm access ---

    private Element getSpPr() {
        Element spPr = findChild(xmlElement, NS_P, "spPr");
        if (spPr != null) return spPr;
        return findChild(xmlElement, NS_P, "grpSpPr");
    }

    private Element getXfrm() {
        // Try spPr (most common: sp, pic, cxnSp)
        Element spPr = getSpPr();
        if (spPr != null) {
            Element xfrm = findChild(spPr, NS_A, "xfrm");
            if (xfrm != null) return xfrm;
        }
        // Try grpSpPr (group shapes)
        if (xmlElement != null) {
            Element grpSpPr = findChild(xmlElement, NS_P, "grpSpPr");
            if (grpSpPr != null) {
                Element xfrm = findChild(grpSpPr, NS_A, "xfrm");
                if (xfrm != null) return xfrm;
            }
            // Try p:xfrm (graphic frames like tables, charts)
            Element pXfrm = findChild(xmlElement, NS_P, "xfrm");
            if (pXfrm != null) return pXfrm;
        }
        return null;
    }

    /**
     * Gets or creates the {@code <a:xfrm>} element.
     *
     * @return the xfrm element
     */
    protected Element ensureXfrm() {
        Element spPr = ensureChild(xmlElement, NS_P, "spPr", "p:spPr");
        return ensureChild(spPr, NS_A, "xfrm", "a:xfrm");
    }

    private long getEmu(String attrName, Element element) {
        if (element == null) return 0;
        String val = element.getAttribute(attrName);
        if (val == null || val.isEmpty()) return 0;
        try {
            return Long.parseLong(val);
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private Element getOff() {
        Element xfrm = getXfrm();
        return xfrm != null ? findChild(xfrm, NS_A, "off") : null;
    }

    private Element getExt() {
        Element xfrm = getXfrm();
        return xfrm != null ? findChild(xfrm, NS_A, "ext") : null;
    }

    @Override
    public double getX() {
        return getEmu("x", getOff()) / EMU_PER_POINT;
    }

    @Override
    public void setX(double x) {
        Element xfrm = ensureXfrm();
        Element off = ensureChild(xfrm, NS_A, "off", "a:off");
        off.setAttribute("x", String.valueOf(Math.round(x * EMU_PER_POINT)));
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public double getY() {
        return getEmu("y", getOff()) / EMU_PER_POINT;
    }

    @Override
    public void setY(double y) {
        Element xfrm = ensureXfrm();
        Element off = ensureChild(xfrm, NS_A, "off", "a:off");
        off.setAttribute("y", String.valueOf(Math.round(y * EMU_PER_POINT)));
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public double getWidth() {
        return getEmu("cx", getExt()) / EMU_PER_POINT;
    }

    @Override
    public void setWidth(double width) {
        Element xfrm = ensureXfrm();
        Element ext = ensureChild(xfrm, NS_A, "ext", "a:ext");
        ext.setAttribute("cx", String.valueOf(Math.round(width * EMU_PER_POINT)));
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public double getHeight() {
        return getEmu("cy", getExt()) / EMU_PER_POINT;
    }

    @Override
    public void setHeight(double height) {
        Element xfrm = ensureXfrm();
        Element ext = ensureChild(xfrm, NS_A, "ext", "a:ext");
        ext.setAttribute("cy", String.valueOf(Math.round(height * EMU_PER_POINT)));
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public double getRotation() {
        Element xfrm = getXfrm();
        if (xfrm == null) return 0;
        String rot = xfrm.getAttribute("rot");
        if (rot == null || rot.isEmpty()) return 0;
        try {
            return Long.parseLong(rot) / 60000.0;
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    @Override
    public void setRotation(double rotation) {
        Element xfrm = ensureXfrm();
        xfrm.setAttribute("rot", String.valueOf(Math.round(rotation * 60000)));
        if (saveCallback != null) saveCallback.run();
    }

    // --- Shape style ---

    @Override
    public IShapeStyle getShapeStyle() {
        return null;
    }

    // --- Shape type ---

    @Override
    public ShapeType getShapeType() {
        if (xmlElement == null) return ShapeType.NOT_DEFINED;
        Element spPr = getSpPr();
        if (spPr == null) return ShapeType.NOT_DEFINED;
        Element prstGeom = findChild(spPr, NS_A, "prstGeom");
        if (prstGeom == null) return ShapeType.NOT_DEFINED;
        String prst = prstGeom.getAttribute("prst");
        if (prst == null || prst.isEmpty()) return ShapeType.NOT_DEFINED;
        return ShapeType.fromOoxml(prst).orElse(ShapeType.NOT_DEFINED);
    }

    @Override
    public void setShapeType(ShapeType value) {
        if (xmlElement == null) return;
        if (value == ShapeType.NOT_DEFINED || value == ShapeType.CUSTOM) return;
        Optional<String> prst = value.toOoxml();
        if (prst.isEmpty()) return;
        Element spPr = ensureChild(xmlElement, NS_P, "spPr", "p:spPr");
        Element prstGeom = ensureChild(spPr, NS_A, "prstGeom", "a:prstGeom");
        prstGeom.setAttribute("prst", prst.get());
        if (saveCallback != null) saveCallback.run();
    }

    // --- Fill format ---

    @Override
    public IFillFormat getFillFormat() {
        if (xmlElement == null) return null;
        Element spPr = ensureChild(xmlElement, NS_P, "spPr", "p:spPr");
        return withPartContext(new FillFormat(spPr, saveCallback));
    }

    // --- Adjustments ---

    @Override
    public IAdjustValueCollection getAdjustments() {
        if (xmlElement == null) return null;
        Element spPr = getSpPr();
        if (spPr == null) return null;
        Element prstGeom = findChild(spPr, NS_A, "prstGeom");
        if (prstGeom == null) return null;
        Element avLst = findChild(prstGeom, NS_A, "avLst");
        if (avLst == null) {
            avLst = ensureChild(prstGeom, NS_A, "avLst", "a:avLst");
        }
        AdjustValueCollection avc = new AdjustValueCollection();
        avc.initInternal(avLst, saveCallback);
        return avc;
    }

    // --- cNvPr ---

    @Override
    public Element getCNvPr() {
        if (xmlElement == null) return null;
        // Check all shape type variants: sp, pic, graphicFrame, grpSp, cxnSp
        for (String nvName : new String[]{"nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"}) {
            Element nvElem = findChild(xmlElement, NS_P, nvName);
            if (nvElem != null) {
                Element cNvPr = findChild(nvElem, NS_P, "cNvPr");
                if (cNvPr != null) return cNvPr;
            }
        }
        return null;
    }

    // --- IShape default implementations ---

    @Override
    public boolean isTextHolder() {
        return false;
    }

    @Override
    public IPlaceholder getPlaceholder() {
        return null;
    }

    @Override
    public ICustomData getCustomData() {
        return null;
    }

    @Override
    public IShapeFrame getRawFrame() {
        return buildFrame();
    }

    @Override
    public void setRawFrame(IShapeFrame value) {
        applyFrame(value);
    }

    @Override
    public IShapeFrame getFrame() {
        return buildFrame();
    }

    @Override
    public void setFrame(IShapeFrame value) {
        applyFrame(value);
    }

    private IShapeFrame buildFrame() {
        if (xmlElement == null) return new ShapeFrame(0, 0, 0, 0, NullableBool.NOT_DEFINED, NullableBool.NOT_DEFINED, 0);
        Element xfrm = getXfrm();
        if (xfrm == null) {
            return new ShapeFrame(0, 0, 0, 0, NullableBool.NOT_DEFINED, NullableBool.NOT_DEFINED, 0);
        }
        Element off = findChild(xfrm, NS_A, "off");
        Element ext = findChild(xfrm, NS_A, "ext");
        double xVal = off != null ? getEmu("x", off) / EMU_PER_POINT : 0.0;
        double yVal = off != null ? getEmu("y", off) / EMU_PER_POINT : 0.0;
        double wVal = ext != null ? getEmu("cx", ext) / EMU_PER_POINT : 0.0;
        double hVal = ext != null ? getEmu("cy", ext) / EMU_PER_POINT : 0.0;
        double rot = 0;
        String rotStr = xfrm.getAttribute("rot");
        if (rotStr != null && !rotStr.isEmpty()) {
            try { rot = Long.parseLong(rotStr) / 60000.0; } catch (NumberFormatException ignored) { /* Non-numeric value; use default */ }
        }
        NullableBool flipH = "1".equals(xfrm.getAttribute("flipH")) ? NullableBool.TRUE : NullableBool.FALSE;
        NullableBool flipV = "1".equals(xfrm.getAttribute("flipV")) ? NullableBool.TRUE : NullableBool.FALSE;
        return new ShapeFrame(xVal, yVal, wVal, hVal, flipH, flipV, rot);
    }

    private void applyFrame(IShapeFrame value) {
        if (value == null || xmlElement == null) return;
        Element xfrm = ensureXfrm();
        Element off = ensureChild(xfrm, NS_A, "off", "a:off");
        Element ext = ensureChild(xfrm, NS_A, "ext", "a:ext");
        off.setAttribute("x", String.valueOf(Math.round(value.getX() * EMU_PER_POINT)));
        off.setAttribute("y", String.valueOf(Math.round(value.getY() * EMU_PER_POINT)));
        ext.setAttribute("cx", String.valueOf(Math.round(value.getWidth() * EMU_PER_POINT)));
        ext.setAttribute("cy", String.valueOf(Math.round(value.getHeight() * EMU_PER_POINT)));
        xfrm.setAttribute("rot", String.valueOf(Math.round(value.getRotation() * 60000)));
        if (value.getFlipH() == NullableBool.TRUE) {
            xfrm.setAttribute("flipH", "1");
        } else {
            xfrm.removeAttribute("flipH");
        }
        if (value.getFlipV() == NullableBool.TRUE) {
            xfrm.setAttribute("flipV", "1");
        } else {
            xfrm.removeAttribute("flipV");
        }
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public ILineFormat getLineFormat() {
        if (xmlElement == null) return null;
        Element spPr = ensureChild(xmlElement, NS_P, "spPr", "p:spPr");
        return new LineFormat(spPr, saveCallback);
    }

    @Override
    public IThreeDFormat getThreeDFormat() {
        if (xmlElement == null) return null;
        Element spPr = ensureChild(xmlElement, NS_P, "spPr", "p:spPr");
        return new ThreeDFormat(spPr, saveCallback);
    }

    @Override
    public IEffectFormat getEffectFormat() {
        if (xmlElement == null) return null;
        Element spPr = ensureChild(xmlElement, NS_P, "spPr", "p:spPr");
        return new EffectFormat(spPr, saveCallback);
    }

    @Override
    public boolean isHidden() {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return false;
        String hidden = cNvPr.getAttribute("hidden");
        return "1".equals(hidden) || "true".equals(hidden);
    }

    @Override
    public void setHidden(boolean value) {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return;
        if (value) {
            cNvPr.setAttribute("hidden", "1");
        } else {
            cNvPr.removeAttribute("hidden");
        }
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public int getZOrderPosition() {
        if (xmlElement == null) return 0;
        var parent = xmlElement.getParentNode();
        if (parent == null) return 0;
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) == xmlElement) return i;
        }
        return 0;
    }

    @Override
    public int getConnectionSiteCount() {
        if (xmlElement == null) return 0;
        Element spPr = findChild(xmlElement, NS_P, "spPr");
        if (spPr != null) {
            Element prstGeom = findChild(spPr, NS_A, "prstGeom");
            if (prstGeom != null) return 8;
        }
        return 0;
    }

    @Override
    public String getAlternativeText() {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return "";
        String descr = cNvPr.getAttribute("descr");
        return descr != null ? descr : "";
    }

    @Override
    public void setAlternativeText(String value) {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return;
        cNvPr.setAttribute("descr", value != null ? value : "");
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public String getAlternativeTextTitle() {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return "";
        String title = cNvPr.getAttribute("title");
        return title != null ? title : "";
    }

    @Override
    public void setAlternativeTextTitle(String value) {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return;
        cNvPr.setAttribute("title", value != null ? value : "");
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public String getName() {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return "";
        String name = cNvPr.getAttribute("name");
        return name != null ? name : "";
    }

    @Override
    public void setName(String value) {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return;
        cNvPr.setAttribute("name", value != null ? value : "");
        if (saveCallback != null) saveCallback.run();
    }

    private static final String DECORATIVE_URI = "http://schemas.microsoft.com/office/drawing/2017/decorative";
    private static final String DECORATIVE_EXT_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

    @Override
    public boolean isDecorative() {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return false;
        Element extLst = findChild(cNvPr, NS_A, "extLst");
        if (extLst == null) return false;
        NodeList exts = extLst.getChildNodes();
        for (int i = 0; i < exts.getLength(); i++) {
            if (exts.item(i) instanceof Element ext
                    && NS_A.equals(ext.getNamespaceURI()) && "ext".equals(ext.getLocalName())) {
                Element decorative = findChild(ext, DECORATIVE_URI, "decorative");
                if (decorative != null) {
                    return "1".equals(decorative.getAttribute("val"));
                }
            }
        }
        return false;
    }

    @Override
    public void setDecorative(boolean value) {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return;
        Document doc = cNvPr.getOwnerDocument();
        Element extLst = findChild(cNvPr, NS_A, "extLst");
        if (extLst == null) {
            extLst = doc.createElementNS(NS_A, "a:extLst");
            cNvPr.appendChild(extLst);
        }
        // Find existing decorative extension
        Element decExt = null;
        NodeList exts = extLst.getChildNodes();
        for (int i = 0; i < exts.getLength(); i++) {
            if (exts.item(i) instanceof Element ext
                    && NS_A.equals(ext.getNamespaceURI()) && "ext".equals(ext.getLocalName())) {
                if (findChild(ext, DECORATIVE_URI, "decorative") != null) {
                    decExt = ext;
                    break;
                }
            }
        }
        if (decExt == null) {
            decExt = doc.createElementNS(NS_A, "a:ext");
            decExt.setAttribute("uri", DECORATIVE_EXT_URI);
            extLst.appendChild(decExt);
        }
        Element decorative = findChild(decExt, DECORATIVE_URI, "decorative");
        if (decorative == null) {
            decorative = doc.createElementNS(DECORATIVE_URI, "adec:decorative");
            decExt.appendChild(decorative);
        }
        decorative.setAttribute("val", value ? "1" : "0");
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public int getUniqueId() {
        Element cNvPr = getCNvPr();
        if (cNvPr == null) return 0;
        String id = cNvPr.getAttribute("id");
        if (id == null || id.isEmpty()) return 0;
        try {
            return Integer.parseInt(id);
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    @Override
    public int getOfficeInteropShapeId() {
        return getUniqueId();
    }

    @Override
    public boolean isGrouped() {
        if (xmlElement == null) return false;
        var parent = xmlElement.getParentNode();
        if (parent instanceof Element parentEl) {
            return NS_P.equals(parentEl.getNamespaceURI()) && "grpSp".equals(parentEl.getLocalName());
        }
        return false;
    }

    @Override
    public ISlideComponent getAsISlideComponent() {
        return this;
    }

    @Override
    public IBaseSlide getSlide() {
        return null;
    }

    @Override
    public IPresentation getPresentation() {
        return null;
    }

    @Override
    public IPresentationComponent asIPresentationComponent() {
        return this;
    }
}
