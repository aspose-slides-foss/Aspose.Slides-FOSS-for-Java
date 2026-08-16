package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.ParserConfigurationException;
import java.util.Map;
import org.aspose.slides.foss.internal.xml.SecureXml;

/**
 * Contains the TextFrame's formatting properties.
 *
 * <p>Wraps the text body element ({@code <a:txBody>}) and manages its
 * {@code <a:bodyPr>} child for text frame formatting.</p>
 */
public final class TextFrameFormat extends PVIObject implements ITextFrameFormat {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final double EMU_PER_POINT = 12700.0;
    private static final double ROTATION_UNIT = 60000.0;

    /** Default left/right margin in EMU (91440 EMU = 7.2pt). */
    private static final int DEFAULT_LR_MARGIN_EMU = 91440;
    /** Default top/bottom margin in EMU (45720 EMU = 3.6pt). */
    private static final int DEFAULT_TB_MARGIN_EMU = 45720;

    private static final Map<String, String> ANCHOR_MAP = Map.of(
            "t", "TOP",
            "ctr", "CENTER",
            "b", "BOTTOM",
            "just", "JUSTIFIED",
            "dist", "DISTRIBUTED"
    );

    private static final Map<String, String> ANCHOR_MAP_REV;
    static {
        var rev = new java.util.HashMap<String, String>();
        ANCHOR_MAP.forEach((k, v) -> rev.put(v, k));
        ANCHOR_MAP_REV = Map.copyOf(rev);
    }

    private static final Map<String, String> VERT_MAP = Map.ofEntries(
            Map.entry("horz", "HORIZONTAL"),
            Map.entry("vert", "VERTICAL"),
            Map.entry("vert270", "VERTICAL270"),
            Map.entry("wordArtVert", "WORD_ART_VERTICAL"),
            Map.entry("eaVert", "EAST_ASIAN_VERTICAL"),
            Map.entry("mongolianVert", "MONGOLIAN_VERTICAL"),
            Map.entry("wordArtVertRtl", "WORD_ART_VERTICAL_RIGHT_TO_LEFT")
    );

    private static final Map<String, String> VERT_MAP_REV;
    static {
        var rev = new java.util.HashMap<String, String>();
        VERT_MAP.forEach((k, v) -> rev.put(v, k));
        VERT_MAP_REV = Map.copyOf(rev);
    }

    private static final Map<String, String> WARP_MAP = Map.ofEntries(
            Map.entry("textNoShape", "NONE"),
            Map.entry("textPlain", "PLAIN"),
            Map.entry("textStop", "STOP"),
            Map.entry("textTriangle", "TRIANGLE"),
            Map.entry("textTriangleInverted", "TRIANGLE_INVERTED"),
            Map.entry("textChevron", "CHEVRON"),
            Map.entry("textChevronInverted", "CHEVRON_INVERTED"),
            Map.entry("textRingInside", "RING_INSIDE"),
            Map.entry("textRingOutside", "RING_OUTSIDE"),
            Map.entry("textArchUp", "ARCH_UP"),
            Map.entry("textArchDown", "ARCH_DOWN"),
            Map.entry("textCircle", "CIRCLE"),
            Map.entry("textButton", "BUTTON"),
            Map.entry("textArchUpPour", "ARCH_UP_POUR"),
            Map.entry("textArchDownPour", "ARCH_DOWN_POUR"),
            Map.entry("textCirclePour", "CIRCLE_POUR"),
            Map.entry("textButtonPour", "BUTTON_POUR"),
            Map.entry("textCurveUp", "CURVE_UP"),
            Map.entry("textCurveDown", "CURVE_DOWN"),
            Map.entry("textCanUp", "CAN_UP"),
            Map.entry("textCanDown", "CAN_DOWN"),
            Map.entry("textWave1", "WAVE1"),
            Map.entry("textWave2", "WAVE2"),
            Map.entry("textDoubleWave1", "DOUBLE_WAVE1"),
            Map.entry("textWave4", "WAVE4"),
            Map.entry("textInflate", "INFLATE"),
            Map.entry("textDeflate", "DEFLATE"),
            Map.entry("textInflateBottom", "INFLATE_BOTTOM"),
            Map.entry("textDeflateBottom", "DEFLATE_BOTTOM"),
            Map.entry("textInflateTop", "INFLATE_TOP"),
            Map.entry("textDeflateTop", "DEFLATE_TOP"),
            Map.entry("textDeflateInflate", "DEFLATE_INFLATE"),
            Map.entry("textDeflateInflateDeflate", "DEFLATE_INFLATE_DEFLATE"),
            Map.entry("textFadeRight", "FADE_RIGHT"),
            Map.entry("textFadeLeft", "FADE_LEFT"),
            Map.entry("textFadeUp", "FADE_UP"),
            Map.entry("textFadeDown", "FADE_DOWN"),
            Map.entry("textSlantUp", "SLANT_UP"),
            Map.entry("textSlantDown", "SLANT_DOWN"),
            Map.entry("textCascadeUp", "CASCADE_UP"),
            Map.entry("textCascadeDown", "CASCADE_DOWN")
    );

    private static final Map<String, String> WARP_MAP_REV;
    static {
        var rev = new java.util.HashMap<String, String>();
        WARP_MAP.forEach((k, v) -> rev.put(v, k));
        WARP_MAP_REV = Map.copyOf(rev);
    }

    private Element txBodyElement;
    private Runnable saveCallback;

    /**
     * Creates a TextFrameFormat backed by the given text body element.
     *
     * @param txBodyElement the {@code <a:txBody>} XML element
     * @param saveCallback  callback invoked after mutations; may be {@code null}
     */
    public TextFrameFormat(Element txBodyElement, Runnable saveCallback) {
        this.txBodyElement = txBodyElement;
        this.saveCallback = saveCallback;
    }

    /**
     * Returns the underlying {@code <a:txBody>} XML element.
     *
     * @return the text body element, or {@code null} if not set
     */
    public Element getTxBodyElement() {
        return txBodyElement;
    }

    /**
     * Creates a new detached TextFrameFormat with a standalone {@code <a:txBody>}
     * containing an empty {@code <a:bodyPr>}.
     */
    public TextFrameFormat() {
        try {
            Document doc = SecureXml.documentBuilderFactory()
                    .newDocumentBuilder()
                    .newDocument();
            this.txBodyElement = doc.createElementNS(NS_A, "a:txBody");
            doc.appendChild(this.txBodyElement);
            Element bodyPr = doc.createElementNS(NS_A, "a:bodyPr");
            this.txBodyElement.appendChild(bodyPr);
        } catch (ParserConfigurationException e) {
            throw new IllegalStateException("Failed to create XML document", e);
        }
        this.saveCallback = null;
    }

    /**
     * Initializes this format with the real XML element from a slide part.
     *
     * @param txBodyElement the {@code <a:txBody>} XML element
     * @param saveCallback  callback invoked after mutations; may be {@code null}
     * @param parentSlide   the parent slide
     * @return this instance for chaining
     */
    public TextFrameFormat initInternal(Element txBodyElement, Runnable saveCallback, IBaseSlide parentSlide) {
        this.txBodyElement = txBodyElement;
        this.saveCallback = saveCallback;
        this.parentSlide = parentSlide;
        return this;
    }

    @Override
    public IPresentation getPresentation() {
        if (parentSlide != null) {
            if (parentSlide instanceof IPresentationComponent pc) {
                return pc.getPresentation();
            }
        }
        return null;
    }

    /**
     * Returns this object as an {@link IPresentationComponent}.
     *
     * @return this instance
     */
    @Override
    public IPresentationComponent asIPresentationComponent() {
        return this;
    }

    private void save() {
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

    private Element getBodyPr() {
        if (txBodyElement == null) return null;
        return findChild(txBodyElement, "bodyPr");
    }

    private Element ensureBodyPr() {
        Element bodyPr = getBodyPr();
        if (bodyPr != null) return bodyPr;
        Document doc = txBodyElement.getOwnerDocument();
        bodyPr = doc.createElementNS(NS_A, "a:bodyPr");
        if (txBodyElement.getFirstChild() != null) {
            txBodyElement.insertBefore(bodyPr, txBodyElement.getFirstChild());
        } else {
            txBodyElement.appendChild(bodyPr);
        }
        return bodyPr;
    }

    private String getAttr(Element el, String name) {
        if (el == null || !el.hasAttribute(name)) return null;
        return el.getAttribute(name);
    }

    private double getMargin(String attr, int defaultEmu) {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return defaultEmu / EMU_PER_POINT;
        String val = getAttr(bodyPr, attr);
        if (val == null) return defaultEmu / EMU_PER_POINT;
        return Long.parseLong(val) / EMU_PER_POINT;
    }

    private void setMargin(String attr, double value) {
        Element bodyPr = ensureBodyPr();
        bodyPr.setAttribute(attr, String.valueOf(Math.round(value * EMU_PER_POINT)));
        save();
    }

    /** {@inheritDoc} */
    @Override
    public double getMarginLeft() {
        return getMargin("lIns", DEFAULT_LR_MARGIN_EMU);
    }

    /** {@inheritDoc} */
    @Override
    public void setMarginLeft(double value) {
        setMargin("lIns", value);
    }

    /** {@inheritDoc} */
    @Override
    public double getMarginRight() {
        return getMargin("rIns", DEFAULT_LR_MARGIN_EMU);
    }

    /** {@inheritDoc} */
    @Override
    public void setMarginRight(double value) {
        setMargin("rIns", value);
    }

    /** {@inheritDoc} */
    @Override
    public double getMarginTop() {
        return getMargin("tIns", DEFAULT_TB_MARGIN_EMU);
    }

    /** {@inheritDoc} */
    @Override
    public void setMarginTop(double value) {
        setMargin("tIns", value);
    }

    /** {@inheritDoc} */
    @Override
    public double getMarginBottom() {
        return getMargin("bIns", DEFAULT_TB_MARGIN_EMU);
    }

    /** {@inheritDoc} */
    @Override
    public void setMarginBottom(double value) {
        setMargin("bIns", value);
    }

    /** {@inheritDoc} */
    @Override
    public NullableBool getWrapText() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return NullableBool.NOT_DEFINED;
        String val = getAttr(bodyPr, "wrap");
        if (val == null) return NullableBool.NOT_DEFINED;
        if ("square".equals(val)) return NullableBool.TRUE;
        if ("none".equals(val)) return NullableBool.FALSE;
        return NullableBool.NOT_DEFINED;
    }

    /** {@inheritDoc} */
    @Override
    public void setWrapText(NullableBool value) {
        Element bodyPr = ensureBodyPr();
        if (value == NullableBool.NOT_DEFINED) {
            bodyPr.removeAttribute("wrap");
        } else {
            bodyPr.setAttribute("wrap", value == NullableBool.TRUE ? "square" : "none");
        }
        save();
    }

    /** {@inheritDoc} */
    @Override
    public TextAnchorType getAnchoringType() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return TextAnchorType.NOT_DEFINED;
        String val = getAttr(bodyPr, "anchor");
        if (val == null) return TextAnchorType.NOT_DEFINED;
        String name = ANCHOR_MAP.get(val);
        if (name == null) return TextAnchorType.NOT_DEFINED;
        return TextAnchorType.valueOf(name);
    }

    /** {@inheritDoc} */
    @Override
    public void setAnchoringType(TextAnchorType value) {
        Element bodyPr = ensureBodyPr();
        if (value == TextAnchorType.NOT_DEFINED) {
            bodyPr.removeAttribute("anchor");
        } else {
            String ooxmlVal = ANCHOR_MAP_REV.get(value.name());
            if (ooxmlVal != null) bodyPr.setAttribute("anchor", ooxmlVal);
        }
        save();
    }

    /** {@inheritDoc} */
    @Override
    public NullableBool getCenterText() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return NullableBool.NOT_DEFINED;
        String val = getAttr(bodyPr, "anchorCtr");
        if (val == null) return NullableBool.NOT_DEFINED;
        return "1".equals(val) ? NullableBool.TRUE : NullableBool.FALSE;
    }

    /** {@inheritDoc} */
    @Override
    public void setCenterText(NullableBool value) {
        Element bodyPr = ensureBodyPr();
        if (value == NullableBool.NOT_DEFINED) {
            bodyPr.removeAttribute("anchorCtr");
        } else {
            bodyPr.setAttribute("anchorCtr", value == NullableBool.TRUE ? "1" : "0");
        }
        save();
    }

    /** {@inheritDoc} */
    @Override
    public TextVerticalType getTextVerticalType() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return TextVerticalType.NOT_DEFINED;
        String val = getAttr(bodyPr, "vert");
        if (val == null) return TextVerticalType.NOT_DEFINED;
        String name = VERT_MAP.get(val);
        if (name == null) return TextVerticalType.NOT_DEFINED;
        return TextVerticalType.valueOf(name);
    }

    /** {@inheritDoc} */
    @Override
    public void setTextVerticalType(TextVerticalType value) {
        Element bodyPr = ensureBodyPr();
        if (value == TextVerticalType.NOT_DEFINED) {
            bodyPr.removeAttribute("vert");
        } else {
            String ooxmlVal = VERT_MAP_REV.get(value.name());
            if (ooxmlVal != null) bodyPr.setAttribute("vert", ooxmlVal);
        }
        save();
    }

    /** {@inheritDoc} */
    @Override
    public TextAutofitType getAutofitType() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return TextAutofitType.NOT_DEFINED;
        if (findChild(bodyPr, "noAutofit") != null) return TextAutofitType.NONE;
        if (findChild(bodyPr, "spAutoFit") != null) return TextAutofitType.SHAPE;
        if (findChild(bodyPr, "normAutofit") != null) return TextAutofitType.NORMAL;
        return TextAutofitType.NOT_DEFINED;
    }

    /** {@inheritDoc} */
    @Override
    public void setAutofitType(TextAutofitType value) {
        Element bodyPr = ensureBodyPr();
        removeChild(bodyPr, "noAutofit");
        removeChild(bodyPr, "spAutoFit");
        removeChild(bodyPr, "normAutofit");
        Document doc = bodyPr.getOwnerDocument();
        switch (value) {
            case NONE -> bodyPr.appendChild(doc.createElementNS(NS_A, "a:noAutofit"));
            case SHAPE -> {
                bodyPr.appendChild(doc.createElementNS(NS_A, "a:spAutoFit"));
                resizeShapeToFitText(bodyPr);
            }
            case NORMAL -> bodyPr.appendChild(doc.createElementNS(NS_A, "a:normAutofit"));
            default -> { /* NOT_DEFINED — no element added */ }
        }
        save();
    }

    private void removeChild(Element parent, String localName) {
        Element el = findChild(parent, localName);
        if (el != null) parent.removeChild(el);
    }

    /** {@inheritDoc} */
    @Override
    public int getColumnCount() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return 0;
        String val = getAttr(bodyPr, "numCol");
        if (val == null) return 0;
        return Math.max(0, Integer.parseInt(val));
    }

    /** {@inheritDoc} */
    @Override
    public void setColumnCount(int value) {
        Element bodyPr = ensureBodyPr();
        value = Math.max(0, value);
        if (value == 0) {
            bodyPr.removeAttribute("numCol");
        } else {
            bodyPr.setAttribute("numCol", String.valueOf(value));
        }
        save();
    }

    /** {@inheritDoc} */
    @Override
    public double getColumnSpacing() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return 0.0;
        String val = getAttr(bodyPr, "spcCol");
        if (val == null) return 0.0;
        return Math.max(0.0, Long.parseLong(val) / EMU_PER_POINT);
    }

    /** {@inheritDoc} */
    @Override
    public void setColumnSpacing(double value) {
        Element bodyPr = ensureBodyPr();
        value = Math.max(0.0, value);
        bodyPr.setAttribute("spcCol", String.valueOf(Math.round(value * EMU_PER_POINT)));
        save();
    }

    /** {@inheritDoc} */
    @Override
    public double getRotationAngle() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return 0.0;
        String val = getAttr(bodyPr, "rot");
        if (val == null) return 0.0;
        return Long.parseLong(val) / ROTATION_UNIT;
    }

    /** {@inheritDoc} */
    @Override
    public void setRotationAngle(double value) {
        Element bodyPr = ensureBodyPr();
        bodyPr.setAttribute("rot", String.valueOf(Math.round(value * ROTATION_UNIT)));
        save();
    }

    /** {@inheritDoc} */
    @Override
    public TextShapeType getTransform() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return TextShapeType.NOT_DEFINED;
        Element prstTxWarp = findChild(bodyPr, "prstTxWarp");
        if (prstTxWarp == null) return TextShapeType.NOT_DEFINED;
        String val = prstTxWarp.getAttribute("prst");
        if (val == null || val.isEmpty()) return TextShapeType.NOT_DEFINED;
        String name = WARP_MAP.get(val);
        if (name == null) return TextShapeType.NOT_DEFINED;
        return TextShapeType.valueOf(name);
    }

    /** {@inheritDoc} */
    @Override
    public void setTransform(TextShapeType value) {
        Element bodyPr = ensureBodyPr();
        if (value == TextShapeType.NOT_DEFINED) {
            removeChild(bodyPr, "prstTxWarp");
        } else {
            Element prstTxWarp = findChild(bodyPr, "prstTxWarp");
            if (prstTxWarp == null) {
                Document doc = bodyPr.getOwnerDocument();
                prstTxWarp = doc.createElementNS(NS_A, "a:prstTxWarp");
                bodyPr.appendChild(prstTxWarp);
            }
            String ooxmlVal = WARP_MAP_REV.get(value.name());
            if (ooxmlVal != null) prstTxWarp.setAttribute("prst", ooxmlVal);
        }
        save();
    }

    /** {@inheritDoc} */
    @Override
    public boolean isKeepTextFlat() {
        Element bodyPr = getBodyPr();
        if (bodyPr == null) return false;
        return "1".equals(getAttr(bodyPr, "upright"));
    }

    /** {@inheritDoc} */
    @Override
    public void setKeepTextFlat(boolean value) {
        Element bodyPr = ensureBodyPr();
        if (value) {
            bodyPr.setAttribute("upright", "1");
        } else {
            bodyPr.removeAttribute("upright");
        }
        save();
    }

    /** {@inheritDoc} */
    @Override
    public IThreeDFormat getThreeDFormat() {
        Element bodyPr = ensureBodyPr();
        return new ThreeDFormat(bodyPr, this::save);
    }

    /**
     * Resizes the parent shape to fit text content, preserving vertical center.
     *
     * <p>Calculates the required height based on paragraph count, default font size
     * (18pt), line spacing (120%), and top/bottom margins, then adjusts the shape's
     * extent and offset to keep it vertically centered.</p>
     *
     * @param bodyPr the {@code <a:bodyPr>} element containing margin attributes
     */
    private void resizeShapeToFitText(Element bodyPr) {
        if (txBodyElement == null) return;

        org.w3c.dom.Node parentNode = txBodyElement.getParentNode();
        if (!(parentNode instanceof Element spElement)) return;

        Element spPr = findChild(spElement, "spPr");
        if (spPr == null) return;

        Element xfrm = findChild(spPr, "xfrm");
        if (xfrm == null) return;

        Element ext = findChild(xfrm, "ext");
        Element off = findChild(xfrm, "off");
        if (ext == null || off == null) return;

        // Count paragraphs
        int numParagraphs = 0;
        NodeList children = txBodyElement.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el && "p".equals(el.getLocalName())) {
                numParagraphs++;
            }
        }
        numParagraphs = Math.max(1, numParagraphs);

        // Default font size 18pt, single line spacing = 120% of font size
        int defaultFontSizeEmu = (int) (18.0 * EMU_PER_POINT);
        int lineHeightEmu = (int) (defaultFontSizeEmu * 1.2);
        int textHeightEmu = numParagraphs * lineHeightEmu;

        // Top and bottom margins (default 45720 EMU = 3.6pt each)
        int topMargin = parseInt(bodyPr, "tIns", 45720);
        int bottomMargin = parseInt(bodyPr, "bIns", 45720);
        int requiredHeight = textHeightEmu + topMargin + bottomMargin;

        // Resize preserving vertical center
        int currentY = parseInt(off, "y", 0);
        int currentCy = parseInt(ext, "cy", 0);
        int centerY = currentY + currentCy / 2;
        int newY = centerY - requiredHeight / 2;

        ext.setAttribute("cy", String.valueOf(requiredHeight));
        off.setAttribute("y", String.valueOf(newY));
    }

    /**
     * Parses an integer attribute from an element, returning a default if absent.
     */
    private int parseInt(Element el, String attr, int defaultValue) {
        String val = getAttr(el, attr);
        if (val == null) return defaultValue;
        return Integer.parseInt(val);
    }
}
