package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.ParserConfigurationException;
import java.util.Map;
import org.aspose.slides.foss.internal.xml.SecureXml;

/**
 * Common text portion formatting properties.
 *
 * <p>Wraps an OOXML {@code <a:rPr>} element and provides read/write access to
 * run-level text formatting attributes and child elements.</p>
 */
public class BasePortionFormat implements IBasePortionFormat {

    /** OOXML Drawing namespace. */
    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // Child element local names
    private static final String LN_HIGHLIGHT = "highlight";
    private static final String LN_U_LN = "uLn";
    private static final String LN_U_LN_TX = "uLnTx";
    private static final String LN_U_FILL = "uFill";
    private static final String LN_U_FILL_TX = "uFillTx";
    private static final String LN_LATIN = "latin";
    private static final String LN_EA = "ea";
    private static final String LN_CS = "cs";
    private static final String LN_SYM = "sym";

    // OOXML underline type attribute values -> TextUnderlineType enum names
    private static final Map<String, TextUnderlineType> UNDERLINE_FROM_OOXML = Map.ofEntries(
            Map.entry("none", TextUnderlineType.NONE),
            Map.entry("words", TextUnderlineType.WORDS),
            Map.entry("sng", TextUnderlineType.SINGLE),
            Map.entry("dbl", TextUnderlineType.DOUBLE),
            Map.entry("heavy", TextUnderlineType.HEAVY),
            Map.entry("dotted", TextUnderlineType.DOTTED),
            Map.entry("dottedHeavy", TextUnderlineType.HEAVY_DOTTED),
            Map.entry("dash", TextUnderlineType.DASHED),
            Map.entry("dashHeavy", TextUnderlineType.HEAVY_DASHED),
            Map.entry("dashLong", TextUnderlineType.LONG_DASHED),
            Map.entry("dashLongHeavy", TextUnderlineType.HEAVY_LONG_DASHED),
            Map.entry("dotDash", TextUnderlineType.DOT_DASH),
            Map.entry("dotDashHeavy", TextUnderlineType.HEAVY_DOT_DASH),
            Map.entry("dotDotDash", TextUnderlineType.DOT_DOT_DASH),
            Map.entry("dotDotDashHeavy", TextUnderlineType.HEAVY_DOT_DOT_DASH),
            Map.entry("wavy", TextUnderlineType.WAVY),
            Map.entry("wavyHeavy", TextUnderlineType.HEAVY_WAVY),
            Map.entry("wavyDbl", TextUnderlineType.DOUBLE_WAVY)
    );

    private static final Map<TextUnderlineType, String> UNDERLINE_TO_OOXML = Map.ofEntries(
            Map.entry(TextUnderlineType.NONE, "none"),
            Map.entry(TextUnderlineType.WORDS, "words"),
            Map.entry(TextUnderlineType.SINGLE, "sng"),
            Map.entry(TextUnderlineType.DOUBLE, "dbl"),
            Map.entry(TextUnderlineType.HEAVY, "heavy"),
            Map.entry(TextUnderlineType.DOTTED, "dotted"),
            Map.entry(TextUnderlineType.HEAVY_DOTTED, "dottedHeavy"),
            Map.entry(TextUnderlineType.DASHED, "dash"),
            Map.entry(TextUnderlineType.HEAVY_DASHED, "dashHeavy"),
            Map.entry(TextUnderlineType.LONG_DASHED, "dashLong"),
            Map.entry(TextUnderlineType.HEAVY_LONG_DASHED, "dashLongHeavy"),
            Map.entry(TextUnderlineType.DOT_DASH, "dotDash"),
            Map.entry(TextUnderlineType.HEAVY_DOT_DASH, "dotDashHeavy"),
            Map.entry(TextUnderlineType.DOT_DOT_DASH, "dotDotDash"),
            Map.entry(TextUnderlineType.HEAVY_DOT_DOT_DASH, "dotDotDashHeavy"),
            Map.entry(TextUnderlineType.WAVY, "wavy"),
            Map.entry(TextUnderlineType.HEAVY_WAVY, "wavyHeavy"),
            Map.entry(TextUnderlineType.DOUBLE_WAVY, "wavyDbl")
    );

    // OOXML cap attribute values -> TextCapType
    private static final Map<String, TextCapType> CAP_FROM_OOXML = Map.of(
            "none", TextCapType.NONE,
            "small", TextCapType.SMALL,
            "all", TextCapType.ALL
    );

    private static final Map<TextCapType, String> CAP_TO_OOXML = Map.of(
            TextCapType.NONE, "none",
            TextCapType.SMALL, "small",
            TextCapType.ALL, "all"
    );

    // OOXML strike attribute values -> TextStrikethroughType
    private static final Map<String, TextStrikethroughType> STRIKE_FROM_OOXML = Map.of(
            "noStrike", TextStrikethroughType.NONE,
            "sngStrike", TextStrikethroughType.SINGLE,
            "dblStrike", TextStrikethroughType.DOUBLE
    );

    private static final Map<TextStrikethroughType, String> STRIKE_TO_OOXML = Map.of(
            TextStrikethroughType.NONE, "noStrike",
            TextStrikethroughType.SINGLE, "sngStrike",
            TextStrikethroughType.DOUBLE, "dblStrike"
    );

    private Element rprElement;
    private Runnable saveCallback;

    /**
     * Creates a new BasePortionFormat with a detached {@code <a:rPr>} element.
     */
    public BasePortionFormat() {
        try {
            Document doc = SecureXml.documentBuilderFactory()
                    .newDocumentBuilder()
                    .newDocument();
            this.rprElement = doc.createElementNS(NS_A, "a:rPr");
            doc.appendChild(this.rprElement);
        } catch (ParserConfigurationException e) {
            throw new IllegalStateException("Failed to create XML document", e);
        }
        this.saveCallback = null;
    }

    /**
     * Creates a new BasePortionFormat backed by the given {@code <a:rPr>} element.
     *
     * @param rprElement   the run properties XML element
     * @param saveCallback callback invoked after mutations to persist changes; may be {@code null}
     */
    public BasePortionFormat(Element rprElement, Runnable saveCallback) {
        this.rprElement = rprElement;
        this.saveCallback = saveCallback;
    }

    /**
     * Returns the underlying {@code <a:rPr>} XML element.
     *
     * @return the run properties element, or {@code null} if not set
     */
    public Element getRprElement() {
        return rprElement;
    }

    // ---- internal helpers ----

    private void save() {
        if (saveCallback != null) {
            saveCallback.run();
        }
    }

    private Element findChild(String localName) {
        NodeList children = rprElement.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if (child instanceof Element el && localName.equals(el.getLocalName())) {
                return el;
            }
        }
        return null;
    }

    private Element getOrCreateChild(String localName) {
        Element el = findChild(localName);
        if (el == null) {
            Document doc = rprElement.getOwnerDocument();
            el = doc.createElementNS(NS_A, "a:" + localName);
            rprElement.appendChild(el);
        }
        return el;
    }

    private void removeChild(String localName) {
        Element el = findChild(localName);
        if (el != null) {
            rprElement.removeChild(el);
        }
    }

    private String getAttr(String name) {
        return rprElement.hasAttribute(name) ? rprElement.getAttribute(name) : null;
    }

    // ---- NullableBool helpers ----

    private NullableBool getNullableBoolAttr(String attr) {
        if (rprElement == null) {
            return NullableBool.NOT_DEFINED;
        }
        String val = getAttr(attr);
        if (val == null) {
            return NullableBool.NOT_DEFINED;
        }
        return "1".equals(val) ? NullableBool.TRUE : NullableBool.FALSE;
    }

    private void setNullableBoolAttr(String attr, NullableBool value) {
        if (rprElement == null) {
            return;
        }
        if (value == NullableBool.NOT_DEFINED) {
            rprElement.removeAttribute(attr);
        } else {
            rprElement.setAttribute(attr, value == NullableBool.TRUE ? "1" : "0");
        }
        save();
    }

    // ---- Font helpers ----

    private IFontData getFont(String localName) {
        if (rprElement == null) {
            return null;
        }
        Element el = findChild(localName);
        if (el == null) {
            return null;
        }
        String typeface = el.hasAttribute("typeface") ? el.getAttribute("typeface") : null;
        if (typeface == null) {
            return null;
        }
        return new FontData(typeface);
    }

    private void setFont(String localName, IFontData value) {
        if (rprElement == null) {
            return;
        }
        if (value == null) {
            removeChild(localName);
        } else {
            Element el = getOrCreateChild(localName);
            el.setAttribute("typeface", value.getFontName());
        }
        save();
    }

    // ---- Read-only format object properties ----

    @Override
    public ILineFormat getLineFormat() {
        if (rprElement == null) {
            return null;
        }
        return new LineFormat(rprElement, this::save);
    }

    @Override
    public IFillFormat getFillFormat() {
        if (rprElement == null) {
            return null;
        }
        return new FillFormat(rprElement, this::save);
    }

    @Override
    public IEffectFormat getEffectFormat() {
        if (rprElement == null) {
            return null;
        }
        return new EffectFormat(rprElement, this::save);
    }

    @Override
    public IColorFormat getHighlightColor() {
        if (rprElement == null) {
            return null;
        }
        Element highlightEl = getOrCreateChild(LN_HIGHLIGHT);
        return new ColorFormat(highlightEl, this::save);
    }

    @Override
    public ILineFormat getUnderlineLineFormat() {
        if (rprElement == null) {
            return null;
        }
        return new LineFormat(rprElement, this::save, LN_U_LN);
    }

    @Override
    public IFillFormat getUnderlineFillFormat() {
        if (rprElement == null) {
            return null;
        }
        Element uFillEl = getOrCreateChild(LN_U_FILL);
        return new FillFormat(uFillEl, this::save);
    }

    // ---- NullableBool attribute properties ----

    @Override
    public NullableBool getFontBold() {
        return getNullableBoolAttr("b");
    }

    @Override
    public void setFontBold(NullableBool value) {
        setNullableBoolAttr("b", value);
    }

    @Override
    public NullableBool getFontItalic() {
        return getNullableBoolAttr("i");
    }

    @Override
    public void setFontItalic(NullableBool value) {
        setNullableBoolAttr("i", value);
    }

    @Override
    public NullableBool getKumimoji() {
        return getNullableBoolAttr("kumimoji");
    }

    @Override
    public void setKumimoji(NullableBool value) {
        setNullableBoolAttr("kumimoji", value);
    }

    @Override
    public NullableBool getNormaliseHeight() {
        return getNullableBoolAttr("normalizeH");
    }

    @Override
    public void setNormaliseHeight(NullableBool value) {
        setNullableBoolAttr("normalizeH", value);
    }

    @Override
    public NullableBool getProofDisabled() {
        return getNullableBoolAttr("noProof");
    }

    @Override
    public void setProofDisabled(NullableBool value) {
        setNullableBoolAttr("noProof", value);
    }

    // ---- Enum attribute properties ----

    @Override
    public TextUnderlineType getFontUnderline() {
        if (rprElement == null) {
            return TextUnderlineType.NOT_DEFINED;
        }
        String val = getAttr("u");
        if (val == null) {
            return TextUnderlineType.NOT_DEFINED;
        }
        return UNDERLINE_FROM_OOXML.getOrDefault(val, TextUnderlineType.NOT_DEFINED);
    }

    @Override
    public void setFontUnderline(TextUnderlineType value) {
        if (rprElement == null) {
            return;
        }
        if (value == TextUnderlineType.NOT_DEFINED) {
            rprElement.removeAttribute("u");
        } else {
            String ooxmlVal = UNDERLINE_TO_OOXML.get(value);
            if (ooxmlVal != null) {
                rprElement.setAttribute("u", ooxmlVal);
            }
        }
        save();
    }

    @Override
    public TextCapType getTextCapType() {
        if (rprElement == null) {
            return TextCapType.NOT_DEFINED;
        }
        String val = getAttr("cap");
        if (val == null) {
            return TextCapType.NOT_DEFINED;
        }
        return CAP_FROM_OOXML.getOrDefault(val, TextCapType.NOT_DEFINED);
    }

    @Override
    public void setTextCapType(TextCapType value) {
        if (rprElement == null) {
            return;
        }
        if (value == TextCapType.NOT_DEFINED) {
            rprElement.removeAttribute("cap");
        } else {
            String ooxmlVal = CAP_TO_OOXML.get(value);
            if (ooxmlVal != null) {
                rprElement.setAttribute("cap", ooxmlVal);
            }
        }
        save();
    }

    @Override
    public TextStrikethroughType getStrikethroughType() {
        if (rprElement == null) {
            return TextStrikethroughType.NOT_DEFINED;
        }
        String val = getAttr("strike");
        if (val == null) {
            return TextStrikethroughType.NOT_DEFINED;
        }
        return STRIKE_FROM_OOXML.getOrDefault(val, TextStrikethroughType.NOT_DEFINED);
    }

    @Override
    public void setStrikethroughType(TextStrikethroughType value) {
        if (rprElement == null) {
            return;
        }
        if (value == TextStrikethroughType.NOT_DEFINED) {
            rprElement.removeAttribute("strike");
        } else {
            String ooxmlVal = STRIKE_TO_OOXML.get(value);
            if (ooxmlVal != null) {
                rprElement.setAttribute("strike", ooxmlVal);
            }
        }
        save();
    }

    // ---- Underline hard/soft properties ----

    @Override
    public NullableBool getIsHardUnderlineLine() {
        if (rprElement == null) {
            return NullableBool.NOT_DEFINED;
        }
        if (findChild(LN_U_LN) != null) {
            return NullableBool.TRUE;
        }
        if (findChild(LN_U_LN_TX) != null) {
            return NullableBool.FALSE;
        }
        return NullableBool.NOT_DEFINED;
    }

    @Override
    public void setIsHardUnderlineLine(NullableBool value) {
        if (rprElement == null) {
            return;
        }
        removeChild(LN_U_LN);
        removeChild(LN_U_LN_TX);
        if (value == NullableBool.TRUE) {
            getOrCreateChild(LN_U_LN);
        } else if (value == NullableBool.FALSE) {
            getOrCreateChild(LN_U_LN_TX);
        }
        save();
    }

    @Override
    public NullableBool getIsHardUnderlineFill() {
        if (rprElement == null) {
            return NullableBool.NOT_DEFINED;
        }
        if (findChild(LN_U_FILL) != null) {
            return NullableBool.TRUE;
        }
        if (findChild(LN_U_FILL_TX) != null) {
            return NullableBool.FALSE;
        }
        return NullableBool.NOT_DEFINED;
    }

    @Override
    public void setIsHardUnderlineFill(NullableBool value) {
        if (rprElement == null) {
            return;
        }
        removeChild(LN_U_FILL);
        removeChild(LN_U_FILL_TX);
        if (value == NullableBool.TRUE) {
            getOrCreateChild(LN_U_FILL);
        } else if (value == NullableBool.FALSE) {
            getOrCreateChild(LN_U_FILL_TX);
        }
        save();
    }

    // ---- Float attribute properties (NaN = undefined) ----

    @Override
    public float getFontHeight() {
        if (rprElement == null) {
            return Float.NaN;
        }
        String val = getAttr("sz");
        if (val == null) {
            return Float.NaN;
        }
        // OOXML stores font size in hundredths of a point
        return Integer.parseInt(val) / 100.0f;
    }

    @Override
    public void setFontHeight(float value) {
        if (rprElement == null) {
            return;
        }
        if (Float.isNaN(value)) {
            rprElement.removeAttribute("sz");
        } else {
            rprElement.setAttribute("sz", String.valueOf(Math.round(value * 100)));
        }
        save();
    }

    @Override
    public float getEscapement() {
        if (rprElement == null) {
            return Float.NaN;
        }
        String val = getAttr("baseline");
        if (val == null) {
            return Float.NaN;
        }
        // OOXML stores as thousandths of percent (e.g., 30000 = 30%)
        return Integer.parseInt(val) / 1000.0f;
    }

    @Override
    public void setEscapement(float value) {
        if (rprElement == null) {
            return;
        }
        if (Float.isNaN(value)) {
            rprElement.removeAttribute("baseline");
        } else {
            rprElement.setAttribute("baseline", String.valueOf(Math.round(value * 1000)));
        }
        save();
    }

    @Override
    public float getKerningMinimalSize() {
        if (rprElement == null) {
            return Float.NaN;
        }
        String val = getAttr("kern");
        if (val == null) {
            return Float.NaN;
        }
        // OOXML stores in hundredths of a point
        return Integer.parseInt(val) / 100.0f;
    }

    @Override
    public void setKerningMinimalSize(float value) {
        if (rprElement == null) {
            return;
        }
        if (Float.isNaN(value)) {
            rprElement.removeAttribute("kern");
        } else {
            rprElement.setAttribute("kern", String.valueOf(Math.round(value * 100)));
        }
        save();
    }

    @Override
    public float getSpacing() {
        if (rprElement == null) {
            return Float.NaN;
        }
        String val = getAttr("spc");
        if (val == null) {
            return Float.NaN;
        }
        // OOXML stores in hundredths of a point
        return Integer.parseInt(val) / 100.0f;
    }

    @Override
    public void setSpacing(float value) {
        if (rprElement == null) {
            return;
        }
        if (Float.isNaN(value)) {
            rprElement.removeAttribute("spc");
        } else {
            rprElement.setAttribute("spc", String.valueOf(Math.round(value * 100)));
        }
        save();
    }

    // ---- Font properties ----

    @Override
    public IFontData getLatinFont() {
        return getFont(LN_LATIN);
    }

    @Override
    public void setLatinFont(IFontData value) {
        setFont(LN_LATIN, value);
    }

    @Override
    public IFontData getEastAsianFont() {
        return getFont(LN_EA);
    }

    @Override
    public void setEastAsianFont(IFontData value) {
        setFont(LN_EA, value);
    }

    @Override
    public IFontData getComplexScriptFont() {
        return getFont(LN_CS);
    }

    @Override
    public void setComplexScriptFont(IFontData value) {
        setFont(LN_CS, value);
    }

    @Override
    public IFontData getSymbolFont() {
        return getFont(LN_SYM);
    }

    @Override
    public void setSymbolFont(IFontData value) {
        setFont(LN_SYM, value);
    }

    // ---- String attribute properties ----

    @Override
    public String getLanguageId() {
        if (rprElement == null) {
            return null;
        }
        return getAttr("lang");
    }

    @Override
    public void setLanguageId(String value) {
        if (rprElement == null) {
            return;
        }
        if (value == null) {
            rprElement.removeAttribute("lang");
        } else {
            rprElement.setAttribute("lang", value);
        }
        save();
    }

    @Override
    public String getAlternativeLanguageId() {
        if (rprElement == null) {
            return null;
        }
        return getAttr("altLang");
    }

    @Override
    public void setAlternativeLanguageId(String value) {
        if (rprElement == null) {
            return;
        }
        if (value == null) {
            rprElement.removeAttribute("altLang");
        } else {
            rprElement.setAttribute("altLang", value);
        }
        save();
    }

    // ---- spell_check property ----

    @Override
    public boolean getSpellCheck() {
        if (rprElement == null) {
            return false;
        }
        String val = getAttr("noProof");
        if ("1".equals(val)) {
            return false;
        }
        // If noProof is '0', spell check is enabled; also enabled if 'err' attribute is present
        return ("0".equals(val)) || rprElement.hasAttribute("err");
    }

    @Override
    public void setSpellCheck(boolean value) {
        if (rprElement == null) {
            return;
        }
        if (!value) {
            rprElement.setAttribute("noProof", "1");
        } else {
            rprElement.removeAttribute("noProof");
        }
        save();
    }
}
