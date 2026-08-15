package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.Map;

/**
 * Represents paragraph bullet formatting properties.
 *
 * <p>Wraps an OOXML {@code <a:pPr>} element for reading and writing bullet properties.</p>
 */
public final class BulletFormat extends PVIObject implements IBulletFormat {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static final Map<String, String> AUTO_NUM_MAP = Map.ofEntries(
            Map.entry("alphaLcParenBoth", "BULLET_ALPHA_LC_PAREN_BOTH"),
            Map.entry("alphaLcParenR", "BULLET_ALPHA_LC_PAREN_RIGHT"),
            Map.entry("alphaLcPeriod", "BULLET_ALPHA_LC_PERIOD"),
            Map.entry("alphaUcParenBoth", "BULLET_ALPHA_UC_PAREN_BOTH"),
            Map.entry("alphaUcParenR", "BULLET_ALPHA_UC_PAREN_RIGHT"),
            Map.entry("alphaUcPeriod", "BULLET_ALPHA_UC_PERIOD"),
            Map.entry("arabicParenBoth", "BULLET_ARABIC_PAREN_BOTH"),
            Map.entry("arabicParenR", "BULLET_ARABIC_PAREN_RIGHT"),
            Map.entry("arabicPeriod", "BULLET_ARABIC_PERIOD"),
            Map.entry("arabicPlain", "BULLET_ARABIC_PLAIN"),
            Map.entry("romanLcParenBoth", "BULLET_ROMAN_LC_PAREN_BOTH"),
            Map.entry("romanLcParenR", "BULLET_ROMAN_LC_PAREN_RIGHT"),
            Map.entry("romanLcPeriod", "BULLET_ROMAN_LC_PERIOD"),
            Map.entry("romanUcParenBoth", "BULLET_ROMAN_UC_PAREN_BOTH"),
            Map.entry("romanUcParenR", "BULLET_ROMAN_UC_PAREN_RIGHT"),
            Map.entry("romanUcPeriod", "BULLET_ROMAN_UC_PERIOD"),
            Map.entry("circleNumDbPlain", "BULLET_CIRCLE_NUM_DB_PLAIN"),
            Map.entry("circleNumWdBlackPlain", "BULLET_CIRCLE_NUM_WD_BLACK_PLAIN"),
            Map.entry("circleNumWdWhitePlain", "BULLET_CIRCLE_NUM_WD_WHITE_PLAIN"),
            Map.entry("ea1ChsPeriod", "BULLET_SIMP_CHIN_PERIOD"),
            Map.entry("ea1ChsPlain", "BULLET_SIMP_CHIN_PLAIN"),
            Map.entry("ea1ChtPeriod", "BULLET_TRAD_CHIN_PERIOD"),
            Map.entry("ea1ChtPlain", "BULLET_TRAD_CHIN_PLAIN"),
            Map.entry("ea1JpnChsDbPeriod", "BULLET_KANJI_SIMP_CHIN_DB_PERIOD"),
            Map.entry("ea1JpnKorPeriod", "BULLET_KANJI_KOREAN_PERIOD"),
            Map.entry("ea1JpnKorPlain", "BULLET_KANJI_KOREAN_PLAIN"),
            Map.entry("arabic1Minus", "BULLET_ARABIC_ALPHA_DASH"),
            Map.entry("arabic2Minus", "BULLET_ARABIC_ABJAD_DASH"),
            Map.entry("hebrew2Minus", "BULLET_HEBREW_ALPHA_DASH"),
            Map.entry("thaiAlphaPeriod", "BULLET_THAI_ALPHA_PERIOD"),
            Map.entry("thaiAlphaParenR", "BULLET_THAI_ALPHA_PAREN_RIGHT"),
            Map.entry("thaiAlphaParenBoth", "BULLET_THAI_ALPHA_PAREN_BOTH"),
            Map.entry("thaiNumPeriod", "BULLET_THAI_NUM_PERIOD"),
            Map.entry("thaiNumParenR", "BULLET_THAI_NUM_PAREN_RIGHT"),
            Map.entry("thaiNumParenBoth", "BULLET_THAI_NUM_PAREN_BOTH"),
            Map.entry("hindiAlphaPeriod", "BULLET_HINDI_ALPHA_PERIOD"),
            Map.entry("hindiNumPeriod", "BULLET_HINDI_NUM_PERIOD"),
            Map.entry("hindiNumParenR", "BULLET_HINDI_NUM_PAREN_RIGHT"),
            Map.entry("hindiAlpha1Period", "BULLET_HINDI_ALPHA_1_PERIOD"),
            Map.entry("arabicDbPeriod", "BULLET_ARABIC_DB_PERIOD"),
            Map.entry("arabicDbPlain", "BULLET_ARABIC_DB_PLAIN")
    );

    private static final Map<String, String> AUTO_NUM_MAP_REV;
    static {
        var rev = new java.util.HashMap<String, String>();
        AUTO_NUM_MAP.forEach((k, v) -> rev.put(v, k));
        AUTO_NUM_MAP_REV = Map.copyOf(rev);
    }

    private Element pprElement;
    private Runnable saveCallback;
    private PartContext partContext;

    /**
     * Creates an unbound {@code BulletFormat}. Call {@link #initInternal} to bind
     * to an XML element.
     */
    public BulletFormat() {
        // no-op; call initInternal to bind to an XML element
    }

    /**
     * Initializes this bullet format from the given {@code <a:pPr>} element.
     *
     * @param pprElement   the {@code <a:pPr>} XML element; may be {@code null}
     * @param saveCallback callback invoked after mutations; may be {@code null}
     * @param parentSlide  the parent slide this object is associated with; may be {@code null}
     * @return this instance, for fluent chaining
     */
    public BulletFormat initInternal(Element pprElement, Runnable saveCallback, IBaseSlide parentSlide) {
        this.pprElement = pprElement;
        this.saveCallback = saveCallback;
        this.parentSlide = parentSlide;
        return this;
    }

    /**
     * Tells this bullet format which package part the paragraph is serialized into, so that
     * a picture bullet can declare its image relationship in the right {@code .rels}.
     *
     * @param partContext the part context, or {@code null} if it is not known
     * @return this instance, for fluent chaining
     */
    BulletFormat withPartContext(PartContext partContext) {
        this.partContext = partContext;
        return this;
    }

    /**
     * Persists changes by invoking the save callback if present.
     */
    private void save() {
        if (saveCallback != null) saveCallback.run();
    }

    private Element findChild(String localName) {
        if (pprElement == null) return null;
        NodeList children = pprElement.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && NS_A.equals(el.getNamespaceURI())
                    && localName.equals(el.getLocalName())) {
                return el;
            }
        }
        return null;
    }

    /**
     * Removes all bullet type elements ({@code buNone}, {@code buChar},
     * {@code buAutoNum}, {@code buBlip}) from the paragraph properties element.
     */
    void removeBulletTypeElements() {
        for (String tag : new String[]{"buNone", "buChar", "buAutoNum", "buBlip"}) {
            Element el = findChild(tag);
            if (el != null) pprElement.removeChild(el);
        }
    }

    private Element createChild(String localName) {
        Document doc = pprElement.getOwnerDocument();
        Element el = doc.createElementNS(NS_A, "a:" + localName);
        pprElement.appendChild(el);
        return el;
    }

    @Override
    public BulletType getType() {
        if (pprElement == null) return BulletType.NOT_DEFINED;
        if (findChild("buNone") != null) return BulletType.NONE;
        if (findChild("buChar") != null) return BulletType.SYMBOL;
        if (findChild("buAutoNum") != null) return BulletType.NUMBERED;
        if (findChild("buBlip") != null) return BulletType.PICTURE;
        return BulletType.NOT_DEFINED;
    }

    @Override
    public void setType(BulletType value) {
        if (pprElement == null) return;
        removeBulletTypeElements();
        switch (value) {
            case NONE -> createChild("buNone");
            case SYMBOL -> createChild("buChar").setAttribute("char", "\u2022");
            case NUMBERED -> createChild("buAutoNum").setAttribute("type", "arabicPeriod");
            case PICTURE -> createChild("buBlip");
            default -> { /* NOT_DEFINED: just removing is enough */ }
        }
        save();
    }

    @Override
    public String getChar() {
        Element el = findChild("buChar");
        if (el == null) return "";
        String val = el.getAttribute("char");
        return val != null ? val : "";
    }

    @Override
    public void setChar(String value) {
        if (pprElement == null) return;
        Element el = findChild("buChar");
        if (el == null) {
            removeBulletTypeElements();
            el = createChild("buChar");
        }
        el.setAttribute("char", value);
        save();
    }

    @Override
    public IFontData getFont() {
        Element el = findChild("buFont");
        if (el == null) return null;
        String typeface = el.getAttribute("typeface");
        if (typeface == null || typeface.isEmpty()) return null;
        return new FontData(typeface);
    }

    @Override
    public void setFont(IFontData value) {
        if (pprElement == null) return;
        Element el = findChild("buFont");
        if (value == null) {
            if (el != null) pprElement.removeChild(el);
        } else {
            if (el == null) el = createChild("buFont");
            el.setAttribute("typeface", value.getFontName());
        }
        save();
    }

    @Override
    public float getHeight() {
        Element pct = findChild("buSzPct");
        if (pct != null) {
            String val = pct.getAttribute("val");
            if (val != null && !val.isEmpty()) {
                return Integer.parseInt(val) / 1000.0f;
            }
        }
        Element pts = findChild("buSzPts");
        if (pts != null) {
            String val = pts.getAttribute("val");
            if (val != null && !val.isEmpty()) {
                return Integer.parseInt(val) / 100.0f;
            }
        }
        return Float.NaN;
    }

    @Override
    public void setHeight(float value) {
        if (pprElement == null) return;
        for (String tag : new String[]{"buSzPct", "buSzPts", "buSzTx"}) {
            Element el = findChild(tag);
            if (el != null) pprElement.removeChild(el);
        }
        if (!Float.isNaN(value)) {
            Element el = createChild("buSzPct");
            el.setAttribute("val", String.valueOf(Math.round(value * 1000)));
        }
        save();
    }

    @Override
    public IColorFormat getColor() {
        Element buClr = findChild("buClr");
        if (buClr == null) {
            buClr = createChild("buClr");
        }
        return new ColorFormat(buClr, saveCallback);
    }

    @Override
    public int getNumberedBulletStartWith() {
        Element el = findChild("buAutoNum");
        if (el == null) return 1;
        String val = el.getAttribute("startAt");
        if (val == null || val.isEmpty()) return 1;
        return Integer.parseInt(val);
    }

    @Override
    public void setNumberedBulletStartWith(int value) {
        if (pprElement == null) return;
        Element el = findChild("buAutoNum");
        if (el == null) {
            removeBulletTypeElements();
            el = createChild("buAutoNum");
            el.setAttribute("type", "arabicPeriod");
        }
        if (value <= 1) {
            el.removeAttribute("startAt");
        } else {
            el.setAttribute("startAt", String.valueOf(value));
        }
        save();
    }

    @Override
    public NumberedBulletStyle getNumberedBulletStyle() {
        Element el = findChild("buAutoNum");
        if (el == null) return NumberedBulletStyle.NOT_DEFINED;
        String val = el.getAttribute("type");
        if (val == null || val.isEmpty()) return NumberedBulletStyle.NOT_DEFINED;
        String name = AUTO_NUM_MAP.get(val);
        if (name == null) return NumberedBulletStyle.NOT_DEFINED;
        return NumberedBulletStyle.valueOf(name);
    }

    @Override
    public void setNumberedBulletStyle(NumberedBulletStyle value) {
        if (pprElement == null) return;
        Element el = findChild("buAutoNum");
        if (value == NumberedBulletStyle.NOT_DEFINED) {
            if (el != null) pprElement.removeChild(el);
            return;
        }
        if (el == null) {
            removeBulletTypeElements();
            el = createChild("buAutoNum");
        }
        String ooxmlVal = AUTO_NUM_MAP_REV.get(value.name());
        if (ooxmlVal != null) el.setAttribute("type", ooxmlVal);
        save();
    }

    @Override
    public NullableBool getIsBulletHardColor() {
        if (findChild("buClr") != null) return NullableBool.TRUE;
        if (findChild("buClrTx") != null) return NullableBool.FALSE;
        return NullableBool.NOT_DEFINED;
    }

    @Override
    public void setIsBulletHardColor(NullableBool value) {
        if (pprElement == null) return;
        if (value == NullableBool.TRUE) {
            Element el = findChild("buClrTx");
            if (el != null) pprElement.removeChild(el);
            if (findChild("buClr") == null) createChild("buClr");
        } else if (value == NullableBool.FALSE) {
            Element el = findChild("buClr");
            if (el != null) pprElement.removeChild(el);
            if (findChild("buClrTx") == null) createChild("buClrTx");
        } else {
            for (String tag : new String[]{"buClr", "buClrTx"}) {
                Element el = findChild(tag);
                if (el != null) pprElement.removeChild(el);
            }
        }
        save();
    }

    @Override
    public ISlidesPicture getPicture() {
        if (pprElement == null) return null;
        Element buBlip = findChild("buBlip");
        if (buBlip == null) {
            buBlip = createChild("buBlip");
        }
        // Find or create <a:blip> inside <a:buBlip>
        Element blip = null;
        var children = buBlip.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && NS_A.equals(el.getNamespaceURI())
                    && "blip".equals(el.getLocalName())) {
                blip = el;
                break;
            }
        }
        if (blip == null) {
            Document doc = pprElement.getOwnerDocument();
            blip = doc.createElementNS(NS_A, "a:blip");
            buBlip.appendChild(blip);
        }
        var pic = new Picture();
        if (partContext != null && partContext.isBound()) {
            pic.initInternal(blip, partContext.pkg(), partContext.partName(), parentSlide);
        } else {
            // No part is known, so setting an image on this picture will refuse rather
            // than write a reference nothing can resolve.
            pic.initInternal(blip, parentSlide);
        }
        return pic;
    }

    @Override
    public NullableBool getIsBulletHardFont() {
        if (findChild("buFont") != null) return NullableBool.TRUE;
        if (findChild("buFontTx") != null) return NullableBool.FALSE;
        return NullableBool.NOT_DEFINED;
    }

    @Override
    public void setIsBulletHardFont(NullableBool value) {
        if (pprElement == null) return;
        if (value == NullableBool.TRUE) {
            Element el = findChild("buFontTx");
            if (el != null) pprElement.removeChild(el);
            if (findChild("buFont") == null) createChild("buFont");
        } else if (value == NullableBool.FALSE) {
            Element el = findChild("buFont");
            if (el != null) pprElement.removeChild(el);
            if (findChild("buFontTx") == null) createChild("buFontTx");
        } else {
            for (String tag : new String[]{"buFont", "buFontTx"}) {
                Element el = findChild(tag);
                if (el != null) pprElement.removeChild(el);
            }
        }
        save();
    }
}
