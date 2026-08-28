package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.LayoutSlidePart;
import org.aspose.slides.foss.internal.pptx.OpcPackage;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.util.Optional;

/**
 * Abstract base class for shapes on a slide.
 *
 * <p>Extends {@link GeometryShape} and adds slide-awareness,
 * placeholder inheritance (layout &rarr; master chain), and
 * two-phase initialization via {@link #initInternal}.</p>
 */
public abstract class Shape extends GeometryShape implements IShape, ISlideComponent, IPresentationComponent {

    /** The parent slide containing this shape. */
    protected IBaseSlide parentSlide;

    /** The OPC package used for placeholder inheritance lookups. */
    private OpcPackage opcPackage;

    /** The layout part name for resolving inherited transforms. */
    private String layoutPartName;

    /**
     * Creates a Shape backed by the given XML element.
     *
     * @param xmlElement   the shape XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public Shape(Element xmlElement, Runnable saveCallback) {
        super(xmlElement, saveCallback);
    }

    /**
     * Creates a Shape with no backing element.
     */
    public Shape() {
        super();
    }

    /**
     * Internal initialization with XML element, OPC package, layout part, and parent slide.
     *
     * <p>Two-phase initialization entry point. It binds the shape to its XML element,
     * sets the save callback, configures the OPC package reference for
     * placeholder inheritance, and records the parent slide.</p>
     *
     * @param xmlElement     the XML element representing this shape
     * @param opcPackage     the OPC package for layout/master lookups; may be {@code null}
     * @param layoutPartName the layout slide part name; may be {@code null}
     * @param saveCallback   callback invoked after mutations; may be {@code null}
     * @param parentSlide    the parent slide object; may be {@code null}
     */
    public void initInternal(Element xmlElement, OpcPackage opcPackage,
                             String layoutPartName, Runnable saveCallback,
                             IBaseSlide parentSlide) {
        this.xmlElement = xmlElement;
        this.saveCallback = saveCallback;
        this.opcPackage = opcPackage;
        this.layoutPartName = layoutPartName;
        this.parentSlide = parentSlide;
    }

    // ── Frame building ──────────────────────────────────────────────────

    /**
     * Builds a {@link ShapeFrame} from the current {@code <a:xfrm>} element.
     *
     * <p>Reads offset, extents, rotation, and flip attributes from the
     * transform element. If no transform is found (including via placeholder
     * inheritance), returns a zeroed frame with {@link NullableBool#NOT_DEFINED}
     * flips.</p>
     *
     * @return the shape frame; never {@code null}
     */
    public IShapeFrame buildFrame() {
        Element xfrm = getXfrm();
        if (xfrm == null) {
            return new ShapeFrame(0, 0, 0, 0, NullableBool.NOT_DEFINED, NullableBool.NOT_DEFINED, 0);
        }
        Element off = findChild(xfrm, NS_A, "off");
        Element ext = findChild(xfrm, NS_A, "ext");
        double xVal = off != null ? emuToPoints(off, "x") : 0.0;
        double yVal = off != null ? emuToPoints(off, "y") : 0.0;
        double wVal = ext != null ? emuToPoints(ext, "cx") : 0.0;
        double hVal = ext != null ? emuToPoints(ext, "cy") : 0.0;
        double rot = parseRotation(xfrm);
        NullableBool flipH = "1".equals(xfrm.getAttribute("flipH"))
                ? NullableBool.TRUE : NullableBool.FALSE;
        NullableBool flipV = "1".equals(xfrm.getAttribute("flipV"))
                ? NullableBool.TRUE : NullableBool.FALSE;
        return new ShapeFrame(xVal, yVal, wVal, hVal, flipH, flipV, rot);
    }

    /**
     * Applies a {@link IShapeFrame}'s values to the {@code <a:xfrm>} element.
     *
     * <p>Creates offset and extent child elements if they do not exist.
     * Sets flip attributes only when {@link NullableBool#TRUE}; otherwise
     * removes them. Invokes the save callback after mutation.</p>
     *
     * @param value the shape frame to apply
     */
    public void applyFrame(IShapeFrame value) {
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

    // ── spPr access ─────────────────────────────────────────────────────

    /**
     * Returns the {@code <p:spPr>} (or {@code <p:grpSpPr>}) element for this shape.
     *
     * @return the shape properties element, or {@code null} if not present
     */
    public Element getSpPr() {
        if (xmlElement == null) return null;
        Element spPr = findChild(xmlElement, NS_P, "spPr");
        if (spPr != null) return spPr;
        return findChild(xmlElement, NS_P, "grpSpPr");
    }

    /**
     * Returns or creates the {@code <p:spPr>} element for this shape.
     *
     * @return the shape properties element; never {@code null}
     * @throws IllegalStateException if the shape has no XML element
     */
    public Element ensureSpPr() {
        Element spPr = getSpPr();
        if (spPr != null) return spPr;
        if (xmlElement == null) {
            throw new IllegalStateException("Shape has no XML element");
        }
        Document doc = xmlElement.getOwnerDocument();
        Element newSpPr = doc.createElementNS(NS_P, "p:spPr");
        xmlElement.appendChild(newSpPr);
        return newSpPr;
    }

    // ── cNvPr access ────────────────────────────────────────────────────

    /**
     * Returns the {@code <p:cNvPr>} element from the shape XML.
     *
     * <p>Handles different shape types (sp, pic, graphicFrame, grpSp, cxnSp)
     * which store the {@code cNvPr} under different {@code nvXxxPr} wrappers.</p>
     *
     * @return the cNvPr element, or {@code null}
     */
    @Override
    public Element getCNvPr() {
        if (xmlElement == null) return null;
        for (String nvName : new String[]{"nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"}) {
            Element nvElem = findChild(xmlElement, NS_P, nvName);
            if (nvElem != null) {
                Element cNvPr = findChild(nvElem, NS_P, "cNvPr");
                if (cNvPr != null) return cNvPr;
            }
        }
        return null;
    }

    // ── xfrm access with placeholder inheritance ────────────────────────

    /**
     * Finds the {@code <a:xfrm>} element directly within a shape XML element.
     *
     * <p>Checks in order: {@code spPr/a:xfrm}, {@code grpSpPr/a:xfrm},
     * then {@code p:xfrm} (for graphic frames like tables and charts).</p>
     *
     * @param shapeElement the shape XML element to search
     * @return the xfrm element, or {@code null} if not found
     */
    public static Element findXfrmInElement(Element shapeElement) {
        if (shapeElement == null) return null;
        // Try spPr (most common: sp, pic, cxnSp)
        Element spPr = findChildStatic(shapeElement, NS_P, "spPr");
        if (spPr != null) {
            Element xfrm = findChildStatic(spPr, NS_A, "xfrm");
            if (xfrm != null) return xfrm;
        }
        // Try grpSpPr (group shapes)
        Element grpSpPr = findChildStatic(shapeElement, NS_P, "grpSpPr");
        if (grpSpPr != null) {
            Element xfrm = findChildStatic(grpSpPr, NS_A, "xfrm");
            if (xfrm != null) return xfrm;
        }
        // Try p:xfrm (graphic frames like tables, charts)
        return findChildStatic(shapeElement, NS_P, "xfrm");
    }

    /**
     * Returns placeholder information for this shape, if it is a placeholder.
     *
     * <p>Searches all {@code nvXxxPr} variants for a {@code <p:nvPr><p:ph>} element
     * and extracts the {@code type} and {@code idx} attributes.</p>
     *
     * @return a two-element array {@code [type, idx]} where type may be {@code null}
     *         and idx defaults to {@code "0"}, or {@link Optional#empty()} if not a placeholder
     */
    public Optional<String[]> getPlaceholderInfo() {
        if (xmlElement == null) return Optional.empty();
        for (String nvName : new String[]{"nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"}) {
            Element nvElem = findChild(xmlElement, NS_P, nvName);
            if (nvElem != null) {
                Element nvPr = findChild(nvElem, NS_P, "nvPr");
                if (nvPr != null) {
                    Element ph = findChild(nvPr, NS_P, "ph");
                    if (ph != null) {
                        String type = ph.getAttribute("type");
                        if (type != null && type.isEmpty()) type = null;
                        String idx = ph.getAttribute("idx");
                        if (idx == null || idx.isEmpty()) idx = "0";
                        return Optional.of(new String[]{type, idx});
                    }
                }
            }
        }
        return Optional.empty();
    }

    /**
     * Finds the {@code <a:xfrm>} for a matching placeholder in a layout or master XML root.
     *
     * <p>Searches the shape tree ({@code <p:spTree>}) for a shape whose
     * {@code <p:ph>} element matches the given type and idx, then returns
     * its transform element.</p>
     *
     * @param root   the layout or master slide root element
     * @param phType the placeholder type (may be {@code null})
     * @param phIdx  the placeholder index
     * @return the matching xfrm element, or {@code null}
     */
    public static Element findPlaceholderXfrmInXml(Element root, String phType, String phIdx) {
        if (root == null) return null;
        // Find spTree
        NodeList spTrees = root.getElementsByTagNameNS(NS_P, "spTree");
        if (spTrees.getLength() == 0) return null;
        Element spTree = (Element) spTrees.item(0);

        NodeList children = spTree.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node node = children.item(i);
            if (!(node instanceof Element child)) continue;

            // Find the nvPr/ph element
            Element ph = null;
            for (String nvName : new String[]{"nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"}) {
                Element nvElem = findChildStatic(child, NS_P, nvName);
                if (nvElem != null) {
                    Element nvPr = findChildStatic(nvElem, NS_P, "nvPr");
                    if (nvPr != null) {
                        ph = findChildStatic(nvPr, NS_P, "ph");
                        break;
                    }
                }
            }
            if (ph == null) continue;

            // Match by type and idx
            String childType = ph.getAttribute("type");
            if (childType != null && childType.isEmpty()) childType = null;
            String childIdx = ph.getAttribute("idx");
            if (childIdx == null || childIdx.isEmpty()) childIdx = "0";

            boolean typeMatch = (phType == null && childType == null)
                    || (phType != null && phType.equals(childType));
            if (typeMatch && phIdx.equals(childIdx)) {
                Element xfrm = findXfrmInElement(child);
                if (xfrm != null) return xfrm;
            }
        }
        return null;
    }

    /**
     * Walks the layout &rarr; master chain to find an inherited {@code <a:xfrm>}
     * for placeholder shapes.
     *
     * @return the inherited xfrm element, or {@code null}
     */
    public Element getInheritedXfrm() {
        Optional<String[]> phInfo = getPlaceholderInfo();
        if (phInfo.isEmpty() || opcPackage == null) return null;
        String phType = phInfo.get()[0];
        String phIdx = phInfo.get()[1];

        // Try layout slide
        if (layoutPartName != null) {
            Document layoutDoc = opcPackage.parseXml(layoutPartName);
            if (layoutDoc != null) {
                Element xfrm = findPlaceholderXfrmInXml(layoutDoc.getDocumentElement(), phType, phIdx);
                if (xfrm != null) return xfrm;

                // Try master slide (resolve from layout's relationships)
                var layoutPart = new LayoutSlidePart(opcPackage, layoutPartName);
                String masterPartName = layoutPart.getMasterPartName();
                if (masterPartName != null) {
                    Document masterDoc = opcPackage.parseXml(masterPartName);
                    if (masterDoc != null) {
                        xfrm = findPlaceholderXfrmInXml(masterDoc.getDocumentElement(), phType, phIdx);
                        if (xfrm != null) return xfrm;
                    }
                }
            }
        }
        return null;
    }

    /**
     * Returns the {@code <a:xfrm>} element for this shape.
     *
     * <p>Handles different shape types:</p>
     * <ul>
     *   <li>{@code sp}, {@code pic}, {@code cxnSp}: {@code spPr/a:xfrm}</li>
     *   <li>{@code grpSp}: {@code grpSpPr/a:xfrm}</li>
     *   <li>{@code graphicFrame}: {@code p:xfrm} (direct child)</li>
     * </ul>
     * <p>For placeholder shapes with no local xfrm, walks the
     * layout &rarr; master inheritance chain.</p>
     *
     * @return the xfrm element, or {@code null}
     */
    public Element getXfrm() {
        if (xmlElement == null) return null;
        Element xfrm = findXfrmInElement(xmlElement);
        if (xfrm != null) return xfrm;
        // Placeholder inheritance: try layout, then master
        return getInheritedXfrm();
    }

    /**
     * Returns or creates the {@code <a:xfrm>} element for this shape.
     *
     * <p>Creates the element under {@code spPr} (or {@code grpSpPr} if present),
     * with zeroed offset and extent children.</p>
     *
     * @return the xfrm element; never {@code null}
     * @throws IllegalStateException if the shape has no XML element
     */
    @Override
    public Element ensureXfrm() {
        Element xfrm = getXfrm();
        if (xfrm != null) return xfrm;
        if (xmlElement == null) {
            throw new IllegalStateException("Shape has no XML element");
        }
        // Create xfrm under the appropriate parent
        Element spPr = findChild(xmlElement, NS_P, "spPr");
        if (spPr == null) {
            Element grpSpPr = findChild(xmlElement, NS_P, "grpSpPr");
            if (grpSpPr != null) {
                spPr = grpSpPr;
            } else {
                spPr = ensureChild(xmlElement, NS_P, "spPr", "p:spPr");
            }
        }
        Document doc = spPr.getOwnerDocument();
        xfrm = doc.createElementNS(NS_A, "a:xfrm");
        spPr.appendChild(xfrm);
        Element off = doc.createElementNS(NS_A, "a:off");
        off.setAttribute("x", "0");
        off.setAttribute("y", "0");
        xfrm.appendChild(off);
        Element ext = doc.createElementNS(NS_A, "a:ext");
        ext.setAttribute("cx", "0");
        ext.setAttribute("cy", "0");
        xfrm.appendChild(ext);
        return xfrm;
    }

    // ── Overrides from GeometryShape ────────────────────────────────────

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

    @Override
    public IBaseSlide getSlide() {
        return parentSlide;
    }

    @Override
    public IPresentation getPresentation() {
        if (parentSlide != null) {
            return parentSlide.getPresentation();
        }
        return null;
    }

    @Override
    public IPresentationComponent asIPresentationComponent() {
        return this;
    }

    /**
     * Sets the parent slide for this shape.
     *
     * @param slide the parent slide
     */
    public void setParentSlide(IBaseSlide slide) {
        this.parentSlide = slide;
    }

    // ── Private helpers ─────────────────────────────────────────────────

    private double emuToPoints(Element element, String attr) {
        String val = element.getAttribute(attr);
        if (val == null || val.isEmpty()) return 0.0;
        try {
            return Long.parseLong(val) / EMU_PER_POINT;
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private double parseRotation(Element xfrm) {
        String rotStr = xfrm.getAttribute("rot");
        if (rotStr == null || rotStr.isEmpty()) return 0.0;
        try {
            return Long.parseLong(rotStr) / 60000.0;
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    /**
     * Static variant of {@link #findChild} for use in static methods.
     */
    private static Element findChildStatic(Element parent, String nsUri, String localName) {
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
}
