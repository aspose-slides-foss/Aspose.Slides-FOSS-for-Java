package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.SchemaOrder;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.Map;

/**
 * Represents 3-D formatting properties for a shape.
 *
 * <p>Wraps the parent shape property element and manages {@code <a:sp3d>}
 * and {@code <a:scene3d>} child elements.</p>
 */
public final class ThreeDFormat extends PVIObject implements IThreeDFormat, IThreeDParamSource {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final double EMU_PER_POINT = 12700.0;

    private static final Map<String, String> MATERIAL_MAP = Map.ofEntries(
            Map.entry("clear", "CLEAR"),
            Map.entry("dkEdge", "DK_EDGE"),
            Map.entry("flat", "FLAT"),
            Map.entry("legacyMatte", "LEGACY_MATTE"),
            Map.entry("legacyMetal", "LEGACY_METAL"),
            Map.entry("legacyPlastic", "LEGACY_PLASTIC"),
            Map.entry("legacyWireframe", "LEGACY_WIREFRAME"),
            Map.entry("matte", "MATTE"),
            Map.entry("metal", "METAL"),
            Map.entry("plastic", "PLASTIC"),
            Map.entry("powder", "POWDER"),
            Map.entry("softEdge", "SOFT_EDGE"),
            Map.entry("softmetal", "SOFTMETAL"),
            Map.entry("translucentPowder", "TRANSLUCENT_POWDER"),
            Map.entry("warmMatte", "WARM_MATTE")
    );

    private static final Map<String, String> MATERIAL_MAP_REV;
    static {
        var rev = new java.util.HashMap<String, String>();
        MATERIAL_MAP.forEach((k, v) -> rev.put(v, k));
        MATERIAL_MAP_REV = Map.copyOf(rev);
    }

    private Element parentElement;
    private Runnable saveCallback;

    /**
     * Creates a new ThreeDFormat.
     *
     * @param parentElement the parent element (e.g. {@code <p:spPr>} or {@code <a:bodyPr>})
     * @param saveCallback  callback invoked after mutations; may be {@code null}
     */
    public ThreeDFormat(Element parentElement, Runnable saveCallback) {
        this.parentElement = parentElement;
        this.saveCallback = saveCallback;
    }

    /**
     * Creates a ThreeDFormat with no initial state.
     *
     * <p>Call {@link #initInternal} to complete initialization.</p>
     */
    public ThreeDFormat() {
    }

    /**
     * Internal initialization.
     *
     * <p>Sets the XML parent element, save callback, and parent slide for this
     * 3-D format object. Uses the two-phase initialization pattern where
     * construction and initialization are separate steps.</p>
     *
     * @param parentElement the XML element containing scene3d/sp3d (e.g. {@code <p:spPr>})
     * @param saveCallback  callback invoked after mutations; may be {@code null}
     * @param parentSlide   the parent slide object
     */
    void initInternal(Element parentElement, Runnable saveCallback, IBaseSlide parentSlide) {
        this.parentElement = parentElement;
        this.saveCallback = saveCallback;
        this.parentSlide = parentSlide;
    }

    /**
     * Persists changes by invoking the save callback, if one was provided.
     */
    void save() {
        if (saveCallback != null) saveCallback.run();
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

    private Element ensureChild(Element parent, String localName) {
        Element el = findChild(parent, localName);
        if (el != null) return el;
        Document doc = parent.getOwnerDocument();
        el = doc.createElementNS(NS_A, "a:" + localName);
        SchemaOrder.insert(parent, el);
        return el;
    }

    /**
     * Returns the {@code <a:sp3d>} child element, or {@code null} if absent.
     *
     * @return the sp3d element, or {@code null}
     */
    Element getSp3d() {
        return findChild(parentElement, "sp3d");
    }

    /**
     * Returns the {@code <a:sp3d>} child element, creating it if absent.
     *
     * @return the sp3d element (never {@code null})
     */
    Element ensureSp3d() {
        return ensureChild(parentElement, "sp3d");
    }

    /**
     * Returns the {@code <a:scene3d>} child element, or {@code null} if absent.
     *
     * @return the scene3d element, or {@code null}
     */
    Element getScene3d() {
        return findChild(parentElement, "scene3d");
    }

    /**
     * Returns the {@code <a:scene3d>} child element, creating it if absent.
     *
     * <p>{@code CT_Scene3D} is {@code camera, lightRig, backdrop?, extLst?} — both
     * of the first two are required, so a scene created because the caller asked
     * for one of them is populated with schema defaults for the other. A scene
     * carrying only the property that was set is not a partial scene; it is
     * invalid content, and a validating reader rejects the whole package.</p>
     *
     * @return the scene3d element (never {@code null})
     */
    Element ensureScene3d() {
        Element existing = getScene3d();
        if (existing != null) return existing;
        Element scene3d = ensureChild(parentElement, "scene3d");
        Camera.ensureCameraElement(scene3d);
        LightRig.ensureLightRigElement(scene3d);
        return scene3d;
    }

    private double getEmuAttr(Element el, String attr) {
        if (el == null) return 0;
        String val = el.getAttribute(attr);
        if (val == null || val.isEmpty()) return 0;
        return Long.parseLong(val) / EMU_PER_POINT;
    }

    private void setEmuAttr(Element el, String attr, double value) {
        el.setAttribute(attr, String.valueOf(Math.round(value * EMU_PER_POINT)));
    }

    @Override
    public double getContourWidth() {
        return getEmuAttr(getSp3d(), "contourW");
    }

    @Override
    public void setContourWidth(double value) {
        Element sp3d = ensureSp3d();
        setEmuAttr(sp3d, "contourW", value);
        save();
    }

    @Override
    public double getExtrusionHeight() {
        return getEmuAttr(getSp3d(), "extrusionH");
    }

    @Override
    public void setExtrusionHeight(double value) {
        Element sp3d = ensureSp3d();
        setEmuAttr(sp3d, "extrusionH", value);
        save();
    }

    @Override
    public double getDepth() {
        return getEmuAttr(getSp3d(), "z");
    }

    @Override
    public void setDepth(double value) {
        Element sp3d = ensureSp3d();
        setEmuAttr(sp3d, "z", value);
        save();
    }

    @Override
    public IShapeBevel getBevelTop() {
        Element sp3d = ensureSp3d();
        return new ShapeBevel(sp3d, this::save, true);
    }

    @Override
    public IShapeBevel getBevelBottom() {
        Element sp3d = ensureSp3d();
        return new ShapeBevel(sp3d, this::save, false);
    }

    @Override
    public IColorFormat getContourColor() {
        Element sp3d = ensureSp3d();
        Element contourClr = ensureChild(sp3d, "contourClr");
        return new ColorFormat(contourClr, this::save);
    }

    @Override
    public IColorFormat getExtrusionColor() {
        Element sp3d = ensureSp3d();
        Element extrusionClr = ensureChild(sp3d, "extrusionClr");
        return new ColorFormat(extrusionClr, this::save);
    }

    @Override
    public ICamera getCamera() {
        Element scene3d = ensureScene3d();
        return new Camera(scene3d, this::save);
    }

    @Override
    public ILightRig getLightRig() {
        Element scene3d = ensureScene3d();
        return new LightRig(scene3d, this::save);
    }

    @Override
    public MaterialPresetType getMaterial() {
        Element sp3d = getSp3d();
        if (sp3d == null) return MaterialPresetType.NOT_DEFINED;
        String val = sp3d.getAttribute("prstMaterial");
        if (val == null || val.isEmpty()) return MaterialPresetType.NOT_DEFINED;
        String name = MATERIAL_MAP.get(val);
        if (name == null) return MaterialPresetType.NOT_DEFINED;
        return MaterialPresetType.valueOf(name);
    }

    @Override
    public void setMaterial(MaterialPresetType value) {
        Element sp3d = ensureSp3d();
        if (value == MaterialPresetType.NOT_DEFINED) {
            sp3d.removeAttribute("prstMaterial");
        } else {
            String ooxmlVal = MATERIAL_MAP_REV.get(value.name());
            if (ooxmlVal != null) sp3d.setAttribute("prstMaterial", ooxmlVal);
        }
        save();
    }
}
