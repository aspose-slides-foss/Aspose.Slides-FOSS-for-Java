package org.aspose.slides.foss;

import org.aspose.slides.foss.effects.*;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.List;

/**
 * Represents effect formatting properties of a shape.
 *
 * <p>Wraps an OOXML parent element (e.g., {@code <p:spPr>}) and manages the
 * {@code <a:effectLst>} child element for reading and writing effect properties.</p>
 */
public final class EffectFormat extends PVIObject implements IEffectFormat {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final double EMU_PER_POINT = 12700.0;

    /** OOXML effectLst child order. */
    private static final List<String> EFFECT_LST_ORDER = List.of(
            "blur", "fillOverlay", "glow", "innerShdw", "outerShdw", "prstShdw", "reflection", "softEdge"
    );

    private Element parentElement;
    private Runnable saveCallback;

    /**
     * Creates a new EffectFormat with no backing element.
     * Call {@link #initInternal} to complete initialization.
     */
    public EffectFormat() {
        super();
    }

    /**
     * Creates a new EffectFormat backed by the given parent XML element.
     *
     * @param parentElement the parent XML element (e.g., {@code <p:spPr>})
     * @param saveCallback  callback invoked after mutations; may be {@code null}
     */
    public EffectFormat(Element parentElement, Runnable saveCallback) {
        super();
        this.parentElement = parentElement;
        this.saveCallback = saveCallback;
    }

    /**
     * Internal initialization that binds this EffectFormat to its XML element,
     * save callback, and parent slide.
     *
     * @param parentElement the XML element that may contain {@code <a:effectLst>}
     *                      (e.g., {@code <p:spPr>})
     * @param saveCallback  callback invoked after mutations; may be {@code null}
     * @param parentSlide   the parent slide object
     */
    public void initInternal(Element parentElement, Runnable saveCallback, IBaseSlide parentSlide) {
        this.parentElement = parentElement;
        this.saveCallback = saveCallback;
        this.parentSlide = parentSlide;
    }

    // --- XML helpers ---

    private Element findChild(Element parent, String localName) {
        if (parent == null) return null;
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el) {
                if (NS_A.equals(el.getNamespaceURI()) && localName.equals(el.getLocalName())) {
                    return el;
                }
            }
        }
        return null;
    }

    private Element getEffectLst() {
        return findChild(parentElement, "effectLst");
    }

    private Element ensureEffectLst() {
        Element el = getEffectLst();
        if (el != null) return el;

        Document doc = parentElement.getOwnerDocument();
        el = doc.createElementNS(NS_A, "a:effectLst");

        // Insert before scene3d, sp3d, extLst
        NodeList children = parentElement.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element child) {
                String name = child.getLocalName();
                if ("scene3d".equals(name) || "sp3d".equals(name) || "extLst".equals(name)) {
                    parentElement.insertBefore(el, child);
                    return el;
                }
            }
        }
        parentElement.appendChild(el);
        return el;
    }

    private Element getEffectChild(String localName) {
        Element effectLst = getEffectLst();
        if (effectLst == null) return null;
        return findChild(effectLst, localName);
    }

    private Element ensureEffectChild(String localName) {
        Element effectLst = ensureEffectLst();
        Element existing = findChild(effectLst, localName);
        if (existing != null) return existing;

        Document doc = effectLst.getOwnerDocument();
        Element el = doc.createElementNS(NS_A, "a:" + localName);
        applyRequiredDefaults(el);

        int newRank = EFFECT_LST_ORDER.indexOf(localName);
        if (newRank < 0) {
            effectLst.appendChild(el);
            return el;
        }

        NodeList children = effectLst.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element child) {
                int childRank = EFFECT_LST_ORDER.indexOf(child.getLocalName());
                if (childRank >= 0 && childRank > newRank) {
                    effectLst.insertBefore(el, child);
                    return el;
                }
            }
        }
        effectLst.appendChild(el);
        return el;
    }

    /**
     * Gives a newly created effect element the children and attributes ECMA-376 §20.1.8
     * marks mandatory for it.
     *
     * <p>Without them the element is schema-incomplete and PowerPoint refuses the whole
     * file rather than ignoring the effect: {@code CT_OuterShadowEffect} and
     * {@code CT_InnerShadowEffect} require a colour, {@code CT_GlowEffect} requires a
     * colour, {@code CT_PresetShadowEffect} requires {@code prst} and a colour,
     * {@code CT_SoftEdgesEffect} requires {@code rad}, and {@code CT_FillOverlayEffect}
     * requires {@code blend} and a fill. {@code a:blur} and {@code a:reflection} have no
     * mandatory content and are left as they are.</p>
     *
     * <p>Applied only when the element is created, so it never overwrites what a caller
     * has already set.</p>
     *
     * @param effect the newly created effect element
     */
    private static void applyRequiredDefaults(Element effect) {
        Document doc = effect.getOwnerDocument();
        switch (effect.getLocalName()) {
            case "outerShdw", "innerShdw" -> effect.appendChild(defaultEffectColor(doc));
            case "glow" -> {
                effect.setAttribute("rad", "63500");
                effect.appendChild(defaultEffectColor(doc));
            }
            case "prstShdw" -> {
                effect.setAttribute("prst", "shdw1");
                effect.appendChild(defaultEffectColor(doc));
            }
            case "softEdge" -> effect.setAttribute("rad", "63500");
            case "fillOverlay" -> {
                effect.setAttribute("blend", "over");
                effect.appendChild(doc.createElementNS(NS_A, "a:noFill"));
            }
            default -> {
                // blur and reflection carry no mandatory content
            }
        }
    }

    /** @return {@code <a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr>} */
    private static Element defaultEffectColor(Document doc) {
        Element color = doc.createElementNS(NS_A, "a:srgbClr");
        color.setAttribute("val", "000000");
        Element alpha = doc.createElementNS(NS_A, "a:alpha");
        alpha.setAttribute("val", "40000");
        color.appendChild(alpha);
        return color;
    }

    private void removeEffectChild(String localName) {
        Element effectLst = getEffectLst();
        if (effectLst == null) return;
        Element child = findChild(effectLst, localName);
        if (child != null) {
            effectLst.removeChild(child);
        }
        // If effectLst is now empty, remove it too
        if (!hasElementChildren(effectLst)) {
            parentElement.removeChild(effectLst);
        }
    }

    private boolean hasElementChildren(Element el) {
        NodeList children = el.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element) return true;
        }
        return false;
    }

    private void save() {
        if (saveCallback != null) saveCallback.run();
    }

    // --- Properties ---

    @Override
    public boolean isNoEffects() {
        Element effectLst = getEffectLst();
        if (effectLst == null) return true;
        return !hasElementChildren(effectLst);
    }

    @Override
    public IBlur getBlurEffect() {
        Element el = getEffectChild("blur");
        if (el == null) return null;
        return new Blur(el, saveCallback);
    }

    @Override
    public void setBlurEffect(IBlur value) {
        if (value == null) {
            removeEffectChild("blur");
        } else {
            Element el = ensureEffectChild("blur");
            el.setAttribute("rad", String.valueOf(Math.round(value.getRadius() * EMU_PER_POINT)));
            el.setAttribute("grow", value.isGrow() ? "1" : "0");
        }
        save();
    }

    @Override
    public IFillOverlay getFillOverlayEffect() {
        Element el = getEffectChild("fillOverlay");
        if (el == null) return null;
        return new FillOverlay(el, saveCallback);
    }

    @Override
    public void setFillOverlayEffect(IFillOverlay value) {
        if (value == null) {
            removeEffectChild("fillOverlay");
        } else {
            ensureEffectChild("fillOverlay");
        }
        save();
    }

    @Override
    public IGlow getGlowEffect() {
        Element el = getEffectChild("glow");
        if (el == null) return null;
        return new Glow(el, saveCallback);
    }

    @Override
    public void setGlowEffect(IGlow value) {
        if (value == null) {
            removeEffectChild("glow");
        } else {
            Element el = ensureEffectChild("glow");
            el.setAttribute("rad", String.valueOf(Math.round(value.getRadius() * EMU_PER_POINT)));
        }
        save();
    }

    @Override
    public IInnerShadow getInnerShadowEffect() {
        Element el = getEffectChild("innerShdw");
        if (el == null) return null;
        return new InnerShadow(el, saveCallback);
    }

    @Override
    public void setInnerShadowEffect(IInnerShadow value) {
        if (value == null) {
            removeEffectChild("innerShdw");
        } else {
            Element el = ensureEffectChild("innerShdw");
            el.setAttribute("blurRad", String.valueOf(Math.round(value.getBlurRadius() * EMU_PER_POINT)));
            el.setAttribute("dist", String.valueOf(Math.round(value.getDistance() * EMU_PER_POINT)));
            el.setAttribute("dir", String.valueOf(Math.round(value.getDirection() * 60000)));
        }
        save();
    }

    @Override
    public IOuterShadow getOuterShadowEffect() {
        Element el = getEffectChild("outerShdw");
        if (el == null) return null;
        return new OuterShadow(el, saveCallback);
    }

    @Override
    public void setOuterShadowEffect(IOuterShadow value) {
        if (value == null) {
            removeEffectChild("outerShdw");
        } else {
            Element el = ensureEffectChild("outerShdw");
            el.setAttribute("blurRad", String.valueOf(Math.round(value.getBlurRadius() * EMU_PER_POINT)));
            el.setAttribute("dist", String.valueOf(Math.round(value.getDistance() * EMU_PER_POINT)));
            el.setAttribute("dir", String.valueOf(Math.round(value.getDirection() * 60000)));
        }
        save();
    }

    @Override
    public IPresetShadow getPresetShadowEffect() {
        Element el = getEffectChild("prstShdw");
        if (el == null) return null;
        return new PresetShadow(el, saveCallback);
    }

    @Override
    public void setPresetShadowEffect(IPresetShadow value) {
        if (value == null) {
            removeEffectChild("prstShdw");
        } else {
            ensureEffectChild("prstShdw");
        }
        save();
    }

    @Override
    public IReflection getReflectionEffect() {
        Element el = getEffectChild("reflection");
        if (el == null) return null;
        return new Reflection(el, saveCallback);
    }

    @Override
    public void setReflectionEffect(IReflection value) {
        if (value == null) {
            removeEffectChild("reflection");
        } else {
            ensureEffectChild("reflection");
        }
        save();
    }

    @Override
    public ISoftEdge getSoftEdgeEffect() {
        Element el = getEffectChild("softEdge");
        if (el == null) return null;
        return new SoftEdge(el, saveCallback);
    }

    @Override
    public void setSoftEdgeEffect(ISoftEdge value) {
        if (value == null) {
            removeEffectChild("softEdge");
        } else {
            Element el = ensureEffectChild("softEdge");
            el.setAttribute("rad", String.valueOf(Math.round(value.getRadius() * EMU_PER_POINT)));
        }
        save();
    }

    @Override
    public IEffectParamSource asIEffectParamSource() {
        return this;
    }

    // --- Enable/disable convenience methods ---

    @Override
    public void setBlurEffect(double radius, boolean grow) {
        Element el = ensureEffectChild("blur");
        el.setAttribute("rad", String.valueOf(Math.round(radius * EMU_PER_POINT)));
        el.setAttribute("grow", grow ? "1" : "0");
        save();
    }

    @Override
    public void enableBlurEffect() {
        ensureEffectChild("blur");
        save();
    }

    @Override
    public void enableFillOverlayEffect() {
        ensureEffectChild("fillOverlay");
        save();
    }

    @Override
    public void enableGlowEffect() {
        ensureEffectChild("glow");
        save();
    }

    @Override
    public void enableInnerShadowEffect() {
        ensureEffectChild("innerShdw");
        save();
    }

    @Override
    public void enableOuterShadowEffect() {
        ensureEffectChild("outerShdw");
        save();
    }

    @Override
    public void enablePresetShadowEffect() {
        ensureEffectChild("prstShdw");
        save();
    }

    @Override
    public void enableReflectionEffect() {
        ensureEffectChild("reflection");
        save();
    }

    @Override
    public void enableSoftEdgeEffect() {
        ensureEffectChild("softEdge");
        save();
    }

    @Override
    public void disableBlurEffect() {
        removeEffectChild("blur");
        save();
    }

    @Override
    public void disableFillOverlayEffect() {
        removeEffectChild("fillOverlay");
        save();
    }

    @Override
    public void disableGlowEffect() {
        removeEffectChild("glow");
        save();
    }

    @Override
    public void disableInnerShadowEffect() {
        removeEffectChild("innerShdw");
        save();
    }

    @Override
    public void disableOuterShadowEffect() {
        removeEffectChild("outerShdw");
        save();
    }

    @Override
    public void disablePresetShadowEffect() {
        removeEffectChild("prstShdw");
        save();
    }

    @Override
    public void disableReflectionEffect() {
        removeEffectChild("reflection");
        save();
    }

    @Override
    public void disableSoftEdgeEffect() {
        removeEffectChild("softEdge");
        save();
    }
}
