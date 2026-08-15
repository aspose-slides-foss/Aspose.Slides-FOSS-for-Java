package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.SchemaOrder;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.Map;

/**
 * Contains the properties of shape's main face relief (bevel).
 *
 * <p>Wraps an OOXML {@code <a:sp3d>} element for reading and writing
 * bevel top or bottom properties.</p>
 */
public final class ShapeBevel extends PVIObject implements IShapeBevel {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final double EMU_PER_POINT = 12700.0;

    private static final Map<String, String> BEVEL_MAP = Map.ofEntries(
            Map.entry("angle", "ANGLE"),
            Map.entry("artDeco", "ART_DECO"),
            Map.entry("circle", "CIRCLE"),
            Map.entry("convex", "CONVEX"),
            Map.entry("coolSlant", "COOL_SLANT"),
            Map.entry("cross", "CROSS"),
            Map.entry("divot", "DIVOT"),
            Map.entry("hardEdge", "HARD_EDGE"),
            Map.entry("relaxedInset", "RELAXED_INSET"),
            Map.entry("riblet", "RIBLET"),
            Map.entry("slope", "SLOPE"),
            Map.entry("softRound", "SOFT_ROUND")
    );

    private static final Map<String, String> BEVEL_MAP_REV;
    static {
        var rev = new java.util.HashMap<String, String>();
        BEVEL_MAP.forEach((k, v) -> rev.put(v, k));
        BEVEL_MAP_REV = Map.copyOf(rev);
    }

    private Element sp3d;
    private Runnable saveCallback;
    private boolean isTopBevel;

    /**
     * Creates an unbound {@code ShapeBevel} for the top bevel.
     * Call {@link #initInternal} to bind to an XML element.
     */
    public ShapeBevel() {
        this.isTopBevel = true;
    }

    /**
     * Creates an unbound {@code ShapeBevel}.
     * Call {@link #initInternal} to bind to an XML element.
     *
     * @param isTopBevel {@code true} for top bevel, {@code false} for bottom bevel
     */
    public ShapeBevel(boolean isTopBevel) {
        this.isTopBevel = isTopBevel;
    }

    /**
     * Creates a new ShapeBevel.
     *
     * @param sp3d         the {@code <a:sp3d>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     * @param isTopBevel   {@code true} for top bevel, {@code false} for bottom bevel
     */
    public ShapeBevel(Element sp3d, Runnable saveCallback, boolean isTopBevel) {
        this.sp3d = sp3d;
        this.saveCallback = saveCallback;
        this.isTopBevel = isTopBevel;
    }

    /**
     * Initializes this bevel with its backing element and context.
     *
     * @param sp3d         the {@code <a:sp3d>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     * @param parentSlide  the slide containing this shape; may be {@code null}
     * @return this instance for fluent chaining
     */
    public ShapeBevel initInternal(Element sp3d, Runnable saveCallback, IBaseSlide parentSlide) {
        this.sp3d = sp3d;
        this.saveCallback = saveCallback;
        this.parentSlide = parentSlide;
        return this;
    }

    private void save() {
        if (saveCallback != null) saveCallback.run();
    }

    private String bevelElementName() {
        return isTopBevel ? "bevelT" : "bevelB";
    }

    private Element findChild(Element parent, String localName) {
        if (parent == null) return null;
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && NS_A.equals(el.getNamespaceURI())
                    && localName.equals(el.getLocalName())) {
                return el;
            }
        }
        return null;
    }

    private Element getBevelElement() {
        return findChild(sp3d, bevelElementName());
    }

    private Element ensureBevelElement() {
        Element el = getBevelElement();
        if (el != null) return el;
        Document doc = sp3d.getOwnerDocument();
        el = doc.createElementNS(NS_A, "a:" + bevelElementName());
        SchemaOrder.insert(sp3d, el);
        return el;
    }

    @Override
    public double getWidth() {
        Element el = getBevelElement();
        if (el == null) return 0;
        String val = el.getAttribute("w");
        if (val == null || val.isEmpty()) return 0;
        return Long.parseLong(val) / EMU_PER_POINT;
    }

    @Override
    public void setWidth(double value) {
        Element el = ensureBevelElement();
        el.setAttribute("w", String.valueOf(Math.round(value * EMU_PER_POINT)));
        save();
    }

    @Override
    public double getHeight() {
        Element el = getBevelElement();
        if (el == null) return 0;
        String val = el.getAttribute("h");
        if (val == null || val.isEmpty()) return 0;
        return Long.parseLong(val) / EMU_PER_POINT;
    }

    @Override
    public void setHeight(double value) {
        Element el = ensureBevelElement();
        el.setAttribute("h", String.valueOf(Math.round(value * EMU_PER_POINT)));
        save();
    }

    @Override
    public BevelPresetType getBevelType() {
        Element el = getBevelElement();
        if (el == null) return BevelPresetType.NOT_DEFINED;
        String val = el.getAttribute("prst");
        if (val == null || val.isEmpty()) return BevelPresetType.NOT_DEFINED;
        String name = BEVEL_MAP.get(val);
        if (name == null) return BevelPresetType.NOT_DEFINED;
        return BevelPresetType.valueOf(name);
    }

    @Override
    public void setBevelType(BevelPresetType value) {
        Element el = ensureBevelElement();
        if (value == BevelPresetType.NOT_DEFINED) {
            el.removeAttribute("prst");
        } else {
            String ooxmlVal = BEVEL_MAP_REV.get(value.name());
            if (ooxmlVal != null) el.setAttribute("prst", ooxmlVal);
        }
        save();
    }
}
