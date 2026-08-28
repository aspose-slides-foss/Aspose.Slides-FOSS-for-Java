package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.ParserConfigurationException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.aspose.slides.foss.internal.xml.SecureXml;

/**
 * Represents paragraph formatting properties.
 *
 * <p>Wraps the paragraph element ({@code <a:p>}) and manages its
 * {@code <a:pPr>} child for paragraph-level formatting.
 * Unlike read-only effective formats, all properties of this class are writeable.</p>
 */
public final class ParagraphFormat extends PVIObject implements IParagraphFormat {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final double EMU_PER_POINT = 12700.0;

    private static final Map<String, String> ALIGNMENT_MAP = Map.of(
            "l", "LEFT",
            "ctr", "CENTER",
            "r", "RIGHT",
            "just", "JUSTIFY",
            "justLow", "JUSTIFY_LOW",
            "dist", "DISTRIBUTED",
            "thaiDist", "THAI_DISTRIBUTED"
    );

    private static final Map<String, String> ALIGNMENT_MAP_REV;
    static {
        var rev = new HashMap<String, String>();
        ALIGNMENT_MAP.forEach((k, v) -> rev.put(v, k));
        ALIGNMENT_MAP_REV = Map.copyOf(rev);
    }

    private static final Map<String, String> FONT_ALIGN_MAP = Map.of(
            "auto", "AUTOMATIC",
            "t", "TOP",
            "ctr", "CENTER",
            "base", "BASELINE",
            "b", "BOTTOM"
    );

    private static final Map<String, String> FONT_ALIGN_MAP_REV;
    static {
        var rev = new HashMap<String, String>();
        FONT_ALIGN_MAP.forEach((k, v) -> rev.put(v, k));
        FONT_ALIGN_MAP_REV = Map.copyOf(rev);
    }

    // --- PPR child ordering per OOXML CT_TextParagraphProperties schema ---

    /**
     * Ordered groups of {@code <a:pPr>} child element local names,
     * reflecting the OOXML schema sequence for {@code CT_TextParagraphProperties}.
     */
    private static final List<List<String>> PPR_CHILD_ORDER = List.of(
            List.of("lnSpc"),
            List.of("spcBef"),
            List.of("spcAft"),
            List.of("buClrTx", "buClr"),
            List.of("buSzTx", "buSzPct", "buSzPts"),
            List.of("buFontTx", "buFont"),
            List.of("buNone", "buAutoNum", "buChar", "buBlip"),
            List.of("tabLst"),
            List.of("defRPr"),
            List.of("extLst")
    );

    /** Maps each local tag name to its positional index in the schema order. */
    private static final Map<String, Integer> PPR_TAG_INDEX;
    static {
        var idx = new HashMap<String, Integer>();
        for (int i = 0; i < PPR_CHILD_ORDER.size(); i++) {
            for (String tag : PPR_CHILD_ORDER.get(i)) {
                idx.put(tag, i);
            }
        }
        PPR_TAG_INDEX = Map.copyOf(idx);
    }

    private Element pElement;
    private Runnable saveCallback;
    // Direct pPr reference when initialized via initInternal (bypasses pElement)
    private Element directPPr;
    // The part this paragraph is serialized into, for a bullet that embeds an image.
    private PartContext partContext;

    /**
     * Creates a ParagraphFormat backed by the given paragraph element.
     *
     * @param pElement     the {@code <a:p>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public ParagraphFormat(Element pElement, Runnable saveCallback) {
        this.pElement = pElement;
        this.saveCallback = saveCallback;
    }

    /**
     * Creates a new detached ParagraphFormat.
     */
    public ParagraphFormat() {
        try {
            Document doc = SecureXml.documentBuilderFactory()
                    .newDocumentBuilder()
                    .newDocument();
            this.pElement = doc.createElementNS(NS_A, "a:p");
            doc.appendChild(this.pElement);
        } catch (ParserConfigurationException e) {
            throw new IllegalStateException("Failed to create XML document", e);
        }
        this.saveCallback = null;
    }

    /**
     * Initializes this format with a direct {@code <a:pPr>} element, save callback,
     * and parent slide reference.
     *
     * @param pprElement  the {@code <a:pPr>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     * @param parentSlide  the parent slide; may be {@code null}
     * @return this instance for method chaining
     */
    public ParagraphFormat initInternal(Element pprElement, Runnable saveCallback, IBaseSlide parentSlide) {
        this.directPPr = pprElement;
        this.saveCallback = saveCallback;
        this.parentSlide = parentSlide;
        return this;
    }

    /**
     * Returns the underlying {@code <a:p>} XML element.
     *
     * @return the paragraph element, or {@code null} if initialized via {@link #initInternal}
     */
    public Element getPElement() {
        return pElement;
    }

    /**
     * Creates and inserts a child element into {@code <a:pPr>} at the correct
     * OOXML schema position.
     *
     * @param ppr        the {@code <a:pPr>} element to insert into
     * @param localName  the local name of the child element (e.g. {@code "lnSpc"})
     * @param attributes optional attribute name-value pairs to set on the new element
     * @return the newly created and inserted element
     */
    public static Element pprInsertChild(Element ppr, String localName, Map<String, String> attributes) {
        int targetPos = PPR_TAG_INDEX.getOrDefault(localName, 999);
        Element insertBefore = null;
        NodeList children = ppr.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node node = children.item(i);
            if (node instanceof Element child) {
                int childPos = PPR_TAG_INDEX.getOrDefault(child.getLocalName(), 999);
                if (childPos > targetPos) {
                    insertBefore = child;
                    break;
                }
            }
        }
        Document doc = ppr.getOwnerDocument();
        Element el = doc.createElementNS(NS_A, "a:" + localName);
        if (attributes != null) {
            attributes.forEach(el::setAttribute);
        }
        if (insertBefore != null) {
            ppr.insertBefore(el, insertBefore);
        } else {
            ppr.appendChild(el);
        }
        return el;
    }

    /**
     * Overload of {@link #pprInsertChild(Element, String, Map)} with no attributes.
     *
     * @param ppr       the {@code <a:pPr>} element
     * @param localName the local name of the child element
     * @return the newly created and inserted element
     */
    public static Element pprInsertChild(Element ppr, String localName) {
        return pprInsertChild(ppr, localName, null);
    }

    /** Invokes the save callback if present. */
    void save() {
        if (saveCallback != null) saveCallback.run();
    }

    private Element findChild(Element parent, String localName) {
        if (parent == null) return null;
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && localName.equals(el.getLocalName())) {
                return el;
            }
        }
        return null;
    }

    private Element getPPr() {
        if (directPPr != null) return directPPr;
        return findChild(pElement, "pPr");
    }

    private Element ensurePPr() {
        if (directPPr != null) return directPPr;
        Element pPr = getPPr();
        if (pPr != null) return pPr;
        Document doc = pElement.getOwnerDocument();
        pPr = doc.createElementNS(NS_A, "a:pPr");
        // Insert pPr as the first child
        if (pElement.getFirstChild() != null) {
            pElement.insertBefore(pPr, pElement.getFirstChild());
        } else {
            pElement.appendChild(pPr);
        }
        return pPr;
    }

    private String getAttr(Element el, String name) {
        if (el == null || !el.hasAttribute(name)) return null;
        return el.getAttribute(name);
    }

    // --- NullableBool attribute helpers ---

    /**
     * Reads a {@link NullableBool} attribute from the {@code <a:pPr>} element.
     *
     * @param attr the attribute name
     * @return the nullable bool value
     */
    NullableBool getNullableBoolAttr(String attr) {
        Element pPr = getPPr();
        if (pPr == null) return NullableBool.NOT_DEFINED;
        String val = getAttr(pPr, attr);
        if (val == null) return NullableBool.NOT_DEFINED;
        return "1".equals(val) ? NullableBool.TRUE : NullableBool.FALSE;
    }

    /**
     * Writes a {@link NullableBool} attribute to the {@code <a:pPr>} element.
     *
     * @param attr  the attribute name
     * @param value the nullable bool value
     */
    void setNullableBoolAttr(String attr, NullableBool value) {
        if (value == NullableBool.NOT_DEFINED) {
            Element pPr = getPPr();
            if (pPr != null) pPr.removeAttribute(attr);
        } else {
            Element pPr = ensurePPr();
            pPr.setAttribute(attr, value == NullableBool.TRUE ? "1" : "0");
        }
        save();
    }

    // --- Spacing helpers ---

    /**
     * Reads spacing from a child element ({@code lnSpc}, {@code spcBef}, {@code spcAft}).
     * Positive return value means percentage, negative means points.
     *
     * @param elementName the local name of the spacing element
     * @return the spacing value, or {@code NaN} if not defined
     */
    double getSpacing(String elementName) {
        Element pPr = getPPr();
        if (pPr == null) return Double.NaN;
        Element el = findChild(pPr, elementName);
        if (el == null) return Double.NaN;
        Element pct = findChild(el, "spcPct");
        if (pct != null) {
            String val = pct.getAttribute("val");
            if (val != null && !val.isEmpty()) {
                return Integer.parseInt(val) / 1000.0;
            }
        }
        Element pts = findChild(el, "spcPts");
        if (pts != null) {
            String val = pts.getAttribute("val");
            if (val != null && !val.isEmpty()) {
                return -(Integer.parseInt(val) / 100.0);
            }
        }
        return Double.NaN;
    }

    /**
     * Writes spacing to a child element.
     * Positive value means percentage ({@code spcPct}), negative means points ({@code spcPts}).
     *
     * @param elementName the local name of the spacing element
     * @param value       the spacing value, or {@code NaN} to remove
     */
    void setSpacing(String elementName, double value) {
        Element pPr = ensurePPr();
        Element el = findChild(pPr, elementName);
        if (Double.isNaN(value)) {
            if (el != null) pPr.removeChild(el);
        } else {
            if (el == null) {
                el = pprInsertChild(pPr, elementName);
            }
            // Remove existing children
            while (el.getFirstChild() != null) {
                el.removeChild(el.getFirstChild());
            }
            Document doc = el.getOwnerDocument();
            if (value >= 0) {
                Element pct = doc.createElementNS(NS_A, "a:spcPct");
                pct.setAttribute("val", String.valueOf(Math.round(value * 1000)));
                el.appendChild(pct);
            } else {
                Element pts = doc.createElementNS(NS_A, "a:spcPts");
                pts.setAttribute("val", String.valueOf(Math.round(-value * 100)));
                el.appendChild(pts);
            }
        }
        save();
    }

    // --- EMU-based attribute helpers ---

    /**
     * Reads an EMU-valued attribute from the {@code <a:pPr>} element and converts to points.
     *
     * @param attr the attribute name
     * @return the value in points, or {@code NaN} if not defined
     */
    double getEmuAttr(String attr) {
        Element pPr = getPPr();
        if (pPr == null) return Double.NaN;
        String val = getAttr(pPr, attr);
        if (val == null) return Double.NaN;
        return Long.parseLong(val) / EMU_PER_POINT;
    }

    /**
     * Writes an EMU-valued attribute to the {@code <a:pPr>} element from a point value.
     *
     * @param attr  the attribute name
     * @param value the value in points, or {@code NaN} to remove
     */
    void setEmuAttr(String attr, double value) {
        if (Double.isNaN(value)) {
            Element pPr = getPPr();
            if (pPr != null) pPr.removeAttribute(attr);
        } else {
            Element pPr = ensurePPr();
            pPr.setAttribute(attr, String.valueOf(Math.round(value * EMU_PER_POINT)));
        }
        save();
    }

    // --- IParagraphFormat implementation ---

    /**
     * Tells this format which package part the paragraph is serialized into.
     *
     * @param partContext the part context, or {@code null} if it is not known
     */
    void setPartContext(PartContext partContext) {
        this.partContext = partContext;
    }

    @Override
    public IBulletFormat getBullet() {
        Element pPr = ensurePPr();
        return new BulletFormat()
                .initInternal(pPr, this::save, this.parentSlide)
                .withPartContext(partContext);
    }

    @Override
    public int getDepth() {
        Element pPr = getPPr();
        if (pPr == null) return 0;
        String val = getAttr(pPr, "lvl");
        if (val == null) return 0;
        return Integer.parseInt(val);
    }

    @Override
    public void setDepth(int value) {
        if (value == 0) {
            Element pPr = getPPr();
            if (pPr != null) pPr.removeAttribute("lvl");
        } else {
            Element pPr = ensurePPr();
            pPr.setAttribute("lvl", String.valueOf(value));
        }
        save();
    }

    @Override
    public TextAlignment getAlignment() {
        Element pPr = getPPr();
        if (pPr == null) return TextAlignment.NOT_DEFINED;
        String val = getAttr(pPr, "algn");
        if (val == null) return TextAlignment.NOT_DEFINED;
        String name = ALIGNMENT_MAP.get(val);
        if (name == null) return TextAlignment.NOT_DEFINED;
        return TextAlignment.valueOf(name);
    }

    @Override
    public void setAlignment(TextAlignment value) {
        if (value == TextAlignment.NOT_DEFINED) {
            Element pPr = getPPr();
            if (pPr != null) pPr.removeAttribute("algn");
        } else {
            Element pPr = ensurePPr();
            String ooxmlVal = ALIGNMENT_MAP_REV.get(value.name());
            if (ooxmlVal != null) pPr.setAttribute("algn", ooxmlVal);
        }
        save();
    }

    @Override
    public double getSpaceWithin() {
        return getSpacing("lnSpc");
    }

    @Override
    public void setSpaceWithin(double value) {
        setSpacing("lnSpc", value);
    }

    @Override
    public double getSpaceBefore() {
        return getSpacing("spcBef");
    }

    @Override
    public void setSpaceBefore(double value) {
        setSpacing("spcBef", value);
    }

    @Override
    public double getSpaceAfter() {
        return getSpacing("spcAft");
    }

    @Override
    public void setSpaceAfter(double value) {
        setSpacing("spcAft", value);
    }

    @Override
    public NullableBool getEastAsianLineBreak() {
        return getNullableBoolAttr("eaLnBrk");
    }

    @Override
    public void setEastAsianLineBreak(NullableBool value) {
        setNullableBoolAttr("eaLnBrk", value);
    }

    @Override
    public NullableBool getRightToLeft() {
        return getNullableBoolAttr("rtl");
    }

    @Override
    public void setRightToLeft(NullableBool value) {
        setNullableBoolAttr("rtl", value);
    }

    @Override
    public NullableBool getLatinLineBreak() {
        return getNullableBoolAttr("latinLnBrk");
    }

    @Override
    public void setLatinLineBreak(NullableBool value) {
        setNullableBoolAttr("latinLnBrk", value);
    }

    @Override
    public NullableBool getHangingPunctuation() {
        return getNullableBoolAttr("hangingPunct");
    }

    @Override
    public void setHangingPunctuation(NullableBool value) {
        setNullableBoolAttr("hangingPunct", value);
    }

    @Override
    public double getMarginLeft() {
        return getEmuAttr("marL");
    }

    @Override
    public void setMarginLeft(double value) {
        setEmuAttr("marL", value);
    }

    @Override
    public double getMarginRight() {
        return getEmuAttr("marR");
    }

    @Override
    public void setMarginRight(double value) {
        setEmuAttr("marR", value);
    }

    @Override
    public double getIndent() {
        return getEmuAttr("indent");
    }

    @Override
    public void setIndent(double value) {
        setEmuAttr("indent", value);
    }

    @Override
    public double getDefaultTabSize() {
        return getEmuAttr("defTabSz");
    }

    @Override
    public void setDefaultTabSize(double value) {
        setEmuAttr("defTabSz", value);
    }

    @Override
    public FontAlignment getFontAlignment() {
        Element pPr = getPPr();
        if (pPr == null) return FontAlignment.DEFAULT;
        String val = getAttr(pPr, "fontAlgn");
        if (val == null) return FontAlignment.DEFAULT;
        String name = FONT_ALIGN_MAP.get(val);
        if (name == null) return FontAlignment.DEFAULT;
        return FontAlignment.valueOf(name);
    }

    @Override
    public void setFontAlignment(FontAlignment value) {
        if (value == FontAlignment.DEFAULT) {
            Element pPr = getPPr();
            if (pPr != null) pPr.removeAttribute("fontAlgn");
        } else {
            Element pPr = ensurePPr();
            String ooxmlVal = FONT_ALIGN_MAP_REV.get(value.name());
            if (ooxmlVal != null) pPr.setAttribute("fontAlgn", ooxmlVal);
        }
        save();
    }

    @Override
    public IBasePortionFormat getDefaultPortionFormat() {
        Element pPr = ensurePPr();
        Element defRPr = findChild(pPr, "defRPr");
        if (defRPr == null) {
            defRPr = pprInsertChild(pPr, "defRPr");
        }
        return new BasePortionFormat(defRPr, this::save);
    }
}
