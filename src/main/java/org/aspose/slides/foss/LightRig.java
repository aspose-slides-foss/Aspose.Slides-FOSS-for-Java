package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.SchemaOrder;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.Map;

/**
 * Represents a light rig for 3D scene.
 *
 * <p>Wraps an OOXML {@code <a:scene3d>} element for reading and writing light rig properties.</p>
 */
public final class LightRig extends PVIObject implements ILightRig {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final double ROTATION_UNIT = 60000.0;

    /** The rig written when a light rig is created or its type is cleared; {@code @rig} is required. */
    private static final String DEFAULT_RIG = "threePt";

    /** The direction written when a light rig is created or cleared; {@code @dir} is required. */
    private static final String DEFAULT_DIRECTION = "t";

    private static final Map<String, String> LIGHT_TYPE_MAP = Map.ofEntries(
            Map.entry("balanced", "BALANCED"),
            Map.entry("brightRoom", "BRIGHT_ROOM"),
            Map.entry("chilly", "CHILLY"),
            Map.entry("contrasting", "CONTRASTING"),
            Map.entry("flat", "FLAT"),
            Map.entry("flood", "FLOOD"),
            Map.entry("freezing", "FREEZING"),
            Map.entry("glow", "GLOW"),
            Map.entry("harsh", "HARSH"),
            Map.entry("legacyFlat1", "LEGACY_FLAT1"),
            Map.entry("legacyFlat2", "LEGACY_FLAT2"),
            Map.entry("legacyFlat3", "LEGACY_FLAT3"),
            Map.entry("legacyFlat4", "LEGACY_FLAT4"),
            Map.entry("legacyHarsh1", "LEGACY_HARSH1"),
            Map.entry("legacyHarsh2", "LEGACY_HARSH2"),
            Map.entry("legacyHarsh3", "LEGACY_HARSH3"),
            Map.entry("legacyHarsh4", "LEGACY_HARSH4"),
            Map.entry("legacyNormal1", "LEGACY_NORMAL1"),
            Map.entry("legacyNormal2", "LEGACY_NORMAL2"),
            Map.entry("legacyNormal3", "LEGACY_NORMAL3"),
            Map.entry("legacyNormal4", "LEGACY_NORMAL4"),
            Map.entry("morning", "MORNING"),
            Map.entry("soft", "SOFT"),
            Map.entry("sunrise", "SUNRISE"),
            Map.entry("sunset", "SUNSET"),
            Map.entry("threePt", "THREE_PT"),
            Map.entry("twoPt", "TWO_PT")
    );

    private static final Map<String, String> LIGHT_TYPE_MAP_REV;
    static {
        var rev = new java.util.HashMap<String, String>();
        LIGHT_TYPE_MAP.forEach((k, v) -> rev.put(v, k));
        LIGHT_TYPE_MAP_REV = Map.copyOf(rev);
    }

    private static final Map<String, String> DIRECTION_MAP = Map.ofEntries(
            Map.entry("t", "TOP"),
            Map.entry("tl", "TOP_LEFT"),
            Map.entry("tr", "TOP_RIGHT"),
            Map.entry("b", "BOTTOM"),
            Map.entry("bl", "BOTTOM_LEFT"),
            Map.entry("br", "BOTTOM_RIGHT"),
            Map.entry("l", "LEFT"),
            Map.entry("r", "RIGHT")
    );

    private static final Map<String, String> DIRECTION_MAP_REV;
    static {
        var rev = new java.util.HashMap<String, String>();
        DIRECTION_MAP.forEach((k, v) -> rev.put(v, k));
        DIRECTION_MAP_REV = Map.copyOf(rev);
    }

    private Element scene3d;
    private Runnable saveCallback;

    /**
     * Creates a new LightRig backed by the given scene3d element.
     *
     * @param scene3d      the {@code <a:scene3d>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public LightRig(Element scene3d, Runnable saveCallback) {
        this.scene3d = scene3d;
        this.saveCallback = saveCallback;
    }

    /**
     * Initializes internal state for this light rig.
     *
     * <p>Two-phase initialization method that sets the backing XML element,
     * save callback, and parent slide association.</p>
     *
     * @param scene3dElement the {@code <a:scene3d>} XML element
     * @param saveCallback   callback invoked after mutations; may be {@code null}
     * @param parentSlide    the parent slide this light rig belongs to
     */
    public void initInternal(Element scene3dElement, Runnable saveCallback, IBaseSlide parentSlide) {
        this.scene3d = scene3dElement;
        this.saveCallback = saveCallback;
        this.parentSlide = parentSlide;
    }

    /**
     * Persists changes by invoking the save callback, if present.
     */
    void save() {
        if (saveCallback != null) saveCallback.run();
    }

    private Element findChild(Element parent, String localName) {
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

    /**
     * Returns the {@code <a:lightRig>} child element, or {@code null} if absent.
     *
     * @return the light rig element, or {@code null}
     */
    Element getLightRig() {
        return findChild(scene3d, "lightRig");
    }

    /**
     * Returns the {@code <a:lightRig>} child element, creating it with defaults if absent.
     *
     * <p>The default light rig uses preset type {@code threePt} and direction {@code t} (top).</p>
     *
     * @return the light rig element, never {@code null}
     */
    Element ensureLightRig() {
        return ensureLightRigElement(scene3d);
    }

    /**
     * Returns the {@code <a:lightRig>} child of the given {@code <a:scene3d>},
     * creating it in schema position with schema defaults if absent.
     *
     * <p>{@code CT_LightRig} requires both {@code @rig} and {@code @dir}, so the
     * element is never written without them.</p>
     *
     * @param scene3d the {@code <a:scene3d>} element
     * @return the light rig element, never {@code null}
     */
    static Element ensureLightRigElement(Element scene3d) {
        NodeList children = scene3d.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && NS_A.equals(el.getNamespaceURI())
                    && "lightRig".equals(el.getLocalName())) {
                return el;
            }
        }
        Document doc = scene3d.getOwnerDocument();
        Element lr = doc.createElementNS(NS_A, "a:lightRig");
        lr.setAttribute("rig", DEFAULT_RIG);
        lr.setAttribute("dir", DEFAULT_DIRECTION);
        SchemaOrder.insert(scene3d, lr);
        return lr;
    }

    @Override
    public LightRigPresetType getLightType() {
        Element lr = getLightRig();
        if (lr == null) return LightRigPresetType.NOT_DEFINED;
        String val = lr.getAttribute("rig");
        if (val == null || val.isEmpty()) return LightRigPresetType.NOT_DEFINED;
        String name = LIGHT_TYPE_MAP.get(val);
        if (name == null) return LightRigPresetType.NOT_DEFINED;
        return LightRigPresetType.valueOf(name);
    }

    @Override
    public void setLightType(LightRigPresetType value) {
        Element lr = ensureLightRig();
        if (value == LightRigPresetType.NOT_DEFINED) {
            // @rig is required on CT_LightRig: clearing the type restores the
            // schema default rather than dropping the attribute.
            lr.setAttribute("rig", DEFAULT_RIG);
        } else {
            String ooxmlVal = LIGHT_TYPE_MAP_REV.get(value.name());
            if (ooxmlVal != null) lr.setAttribute("rig", ooxmlVal);
        }
        save();
    }

    @Override
    public LightingDirection getDirection() {
        Element lr = getLightRig();
        if (lr == null) return LightingDirection.NOT_DEFINED;
        String val = lr.getAttribute("dir");
        if (val == null || val.isEmpty()) return LightingDirection.NOT_DEFINED;
        String name = DIRECTION_MAP.get(val);
        if (name == null) return LightingDirection.NOT_DEFINED;
        return LightingDirection.valueOf(name);
    }

    @Override
    public void setDirection(LightingDirection value) {
        Element lr = ensureLightRig();
        if (value == LightingDirection.NOT_DEFINED) {
            // @dir is required on CT_LightRig, as @rig is.
            lr.setAttribute("dir", DEFAULT_DIRECTION);
        } else {
            String ooxmlVal = DIRECTION_MAP_REV.get(value.name());
            if (ooxmlVal != null) lr.setAttribute("dir", ooxmlVal);
        }
        save();
    }

    @Override
    public void setRotation(double latitude, double longitude, double revolution) {
        Element lr = ensureLightRig();
        Element rot = findChild(lr, "rot");
        if (rot == null) {
            Document doc = lr.getOwnerDocument();
            rot = doc.createElementNS(NS_A, "a:rot");
            lr.appendChild(rot);
        }
        rot.setAttribute("lat", String.valueOf(Math.round(latitude * ROTATION_UNIT)));
        rot.setAttribute("lon", String.valueOf(Math.round(longitude * ROTATION_UNIT)));
        rot.setAttribute("rev", String.valueOf(Math.round(revolution * ROTATION_UNIT)));
        save();
    }

    @Override
    public double[] getRotation() {
        Element lr = getLightRig();
        if (lr == null) return new double[]{0.0, 0.0, 0.0};
        Element rot = findChild(lr, "rot");
        if (rot == null) return new double[]{0.0, 0.0, 0.0};
        double lat = parseLong(rot.getAttribute("lat")) / ROTATION_UNIT;
        double lon = parseLong(rot.getAttribute("lon")) / ROTATION_UNIT;
        double rev = parseLong(rot.getAttribute("rev")) / ROTATION_UNIT;
        return new double[]{lat, lon, rev};
    }

    private static long parseLong(String s) {
        if (s == null || s.isEmpty()) return 0;
        return Long.parseLong(s);
    }
}
