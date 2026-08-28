package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.SchemaOrder;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.Map;

/**
 * Represents 3D camera settings.
 *
 * <p>Wraps an OOXML {@code <a:scene3d>} element for reading and writing camera properties.</p>
 */
public final class Camera implements ICamera {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final double ROTATION_UNIT = 60000.0;

    /** The preset written when a camera is created or its type is cleared; {@code @prst} is required. */
    private static final String DEFAULT_PRESET = "orthographicFront";

    private static final Map<String, String> CAMERA_PRST_MAP = Map.ofEntries(
            Map.entry("isometricBottomDown", "ISOMETRIC_BOTTOM_DOWN"),
            Map.entry("isometricBottomUp", "ISOMETRIC_BOTTOM_UP"),
            Map.entry("isometricLeftDown", "ISOMETRIC_LEFT_DOWN"),
            Map.entry("isometricLeftUp", "ISOMETRIC_LEFT_UP"),
            Map.entry("isometricOffAxis1Left", "ISOMETRIC_OFF_AXIS_1_LEFT"),
            Map.entry("isometricOffAxis1Right", "ISOMETRIC_OFF_AXIS_1_RIGHT"),
            Map.entry("isometricOffAxis1Top", "ISOMETRIC_OFF_AXIS_1_TOP"),
            Map.entry("isometricOffAxis2Left", "ISOMETRIC_OFF_AXIS_2_LEFT"),
            Map.entry("isometricOffAxis2Right", "ISOMETRIC_OFF_AXIS_2_RIGHT"),
            Map.entry("isometricOffAxis2Top", "ISOMETRIC_OFF_AXIS_2_TOP"),
            Map.entry("isometricOffAxis3Bottom", "ISOMETRIC_OFF_AXIS_3_BOTTOM"),
            Map.entry("isometricOffAxis3Left", "ISOMETRIC_OFF_AXIS_3_LEFT"),
            Map.entry("isometricOffAxis3Right", "ISOMETRIC_OFF_AXIS_3_RIGHT"),
            Map.entry("isometricOffAxis4Bottom", "ISOMETRIC_OFF_AXIS_4_BOTTOM"),
            Map.entry("isometricOffAxis4Left", "ISOMETRIC_OFF_AXIS_4_LEFT"),
            Map.entry("isometricOffAxis4Right", "ISOMETRIC_OFF_AXIS_4_RIGHT"),
            Map.entry("isometricRightDown", "ISOMETRIC_RIGHT_DOWN"),
            Map.entry("isometricRightUp", "ISOMETRIC_RIGHT_UP"),
            Map.entry("isometricTopDown", "ISOMETRIC_TOP_DOWN"),
            Map.entry("isometricTopUp", "ISOMETRIC_TOP_UP"),
            Map.entry("legacyObliqueBottom", "LEGACY_OBLIQUE_BOTTOM"),
            Map.entry("legacyObliqueBottomLeft", "LEGACY_OBLIQUE_BOTTOM_LEFT"),
            Map.entry("legacyObliqueBottomRight", "LEGACY_OBLIQUE_BOTTOM_RIGHT"),
            Map.entry("legacyObliqueFront", "LEGACY_OBLIQUE_FRONT"),
            Map.entry("legacyObliqueLeft", "LEGACY_OBLIQUE_LEFT"),
            Map.entry("legacyObliqueRight", "LEGACY_OBLIQUE_RIGHT"),
            Map.entry("legacyObliqueTop", "LEGACY_OBLIQUE_TOP"),
            Map.entry("legacyObliqueTopLeft", "LEGACY_OBLIQUE_TOP_LEFT"),
            Map.entry("legacyObliqueTopRight", "LEGACY_OBLIQUE_TOP_RIGHT"),
            Map.entry("legacyPerspectiveBottom", "LEGACY_PERSPECTIVE_BOTTOM"),
            Map.entry("legacyPerspectiveBottomLeft", "LEGACY_PERSPECTIVE_BOTTOM_LEFT"),
            Map.entry("legacyPerspectiveBottomRight", "LEGACY_PERSPECTIVE_BOTTOM_RIGHT"),
            Map.entry("legacyPerspectiveFront", "LEGACY_PERSPECTIVE_FRONT"),
            Map.entry("legacyPerspectiveLeft", "LEGACY_PERSPECTIVE_LEFT"),
            Map.entry("legacyPerspectiveRight", "LEGACY_PERSPECTIVE_RIGHT"),
            Map.entry("legacyPerspectiveTop", "LEGACY_PERSPECTIVE_TOP"),
            Map.entry("legacyPerspectiveTopLeft", "LEGACY_PERSPECTIVE_TOP_LEFT"),
            Map.entry("legacyPerspectiveTopRight", "LEGACY_PERSPECTIVE_TOP_RIGHT"),
            Map.entry("obliqueBottom", "OBLIQUE_BOTTOM"),
            Map.entry("obliqueBottomLeft", "OBLIQUE_BOTTOM_LEFT"),
            Map.entry("obliqueBottomRight", "OBLIQUE_BOTTOM_RIGHT"),
            Map.entry("obliqueLeft", "OBLIQUE_LEFT"),
            Map.entry("obliqueRight", "OBLIQUE_RIGHT"),
            Map.entry("obliqueTop", "OBLIQUE_TOP"),
            Map.entry("obliqueTopLeft", "OBLIQUE_TOP_LEFT"),
            Map.entry("obliqueTopRight", "OBLIQUE_TOP_RIGHT"),
            Map.entry("orthographicFront", "ORTHOGRAPHIC_FRONT"),
            Map.entry("perspectiveAbove", "PERSPECTIVE_ABOVE"),
            Map.entry("perspectiveAboveLeftFacing", "PERSPECTIVE_ABOVE_LEFT_FACING"),
            Map.entry("perspectiveAboveRightFacing", "PERSPECTIVE_ABOVE_RIGHT_FACING"),
            Map.entry("perspectiveBelow", "PERSPECTIVE_BELOW"),
            Map.entry("perspectiveContrastingLeftFacing", "PERSPECTIVE_CONTRASTING_LEFT_FACING"),
            Map.entry("perspectiveContrastingRightFacing", "PERSPECTIVE_CONTRASTING_RIGHT_FACING"),
            Map.entry("perspectiveFront", "PERSPECTIVE_FRONT"),
            Map.entry("perspectiveHeroicExtremeLeftFacing", "PERSPECTIVE_HEROIC_EXTREME_LEFT_FACING"),
            Map.entry("perspectiveHeroicExtremeRightFacing", "PERSPECTIVE_HEROIC_EXTREME_RIGHT_FACING"),
            Map.entry("perspectiveHeroicLeftFacing", "PERSPECTIVE_HEROIC_LEFT_FACING"),
            Map.entry("perspectiveHeroicRightFacing", "PERSPECTIVE_HEROIC_RIGHT_FACING"),
            Map.entry("perspectiveLeft", "PERSPECTIVE_LEFT"),
            Map.entry("perspectiveRelaxed", "PERSPECTIVE_RELAXED"),
            Map.entry("perspectiveRelaxedModerately", "PERSPECTIVE_RELAXED_MODERATELY"),
            Map.entry("perspectiveRight", "PERSPECTIVE_RIGHT")
    );

    private static final Map<String, String> CAMERA_PRST_MAP_REV;
    static {
        var rev = new java.util.HashMap<String, String>();
        CAMERA_PRST_MAP.forEach((k, v) -> rev.put(v, k));
        CAMERA_PRST_MAP_REV = Map.copyOf(rev);
    }

    private final Element scene3d;
    private final Runnable saveCallback;

    /**
     * Creates a new Camera backed by the given scene3d element.
     *
     * @param scene3d      the {@code <a:scene3d>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public Camera(Element scene3d, Runnable saveCallback) {
        this.scene3d = scene3d;
        this.saveCallback = saveCallback;
    }

    private void save() {
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

    private Element getCamera() {
        return findChild(scene3d, "camera");
    }

    private Element ensureCamera() {
        return ensureCameraElement(scene3d);
    }

    /**
     * Returns the {@code <a:camera>} child of the given {@code <a:scene3d>},
     * creating it in schema position with the default preset if absent.
     *
     * <p>{@code CT_Camera/@prst} is required, so the element is never written
     * without one.</p>
     *
     * @param scene3d the {@code <a:scene3d>} element
     * @return the camera element, never {@code null}
     */
    static Element ensureCameraElement(Element scene3d) {
        NodeList children = scene3d.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && NS_A.equals(el.getNamespaceURI())
                    && "camera".equals(el.getLocalName())) {
                return el;
            }
        }
        Document doc = scene3d.getOwnerDocument();
        Element cam = doc.createElementNS(NS_A, "a:camera");
        cam.setAttribute("prst", DEFAULT_PRESET);
        SchemaOrder.insert(scene3d, cam);
        return cam;
    }

    @Override
    public CameraPresetType getCameraType() {
        Element cam = getCamera();
        if (cam == null) return CameraPresetType.NOT_DEFINED;
        String val = cam.getAttribute("prst");
        if (val == null || val.isEmpty()) return CameraPresetType.NOT_DEFINED;
        String name = CAMERA_PRST_MAP.get(val);
        if (name == null) return CameraPresetType.NOT_DEFINED;
        return CameraPresetType.valueOf(name);
    }

    @Override
    public void setCameraType(CameraPresetType value) {
        Element cam = ensureCamera();
        if (value == CameraPresetType.NOT_DEFINED) {
            // @prst is required on CT_Camera: clearing the type restores the
            // schema default rather than dropping the attribute.
            cam.setAttribute("prst", DEFAULT_PRESET);
        } else {
            String ooxmlVal = CAMERA_PRST_MAP_REV.get(value.name());
            if (ooxmlVal != null) cam.setAttribute("prst", ooxmlVal);
        }
        save();
    }

    @Override
    public double getFieldOfViewAngle() {
        Element cam = getCamera();
        if (cam == null) return 0.0;
        String fov = cam.getAttribute("fov");
        if (fov == null || fov.isEmpty()) return 0.0;
        return Long.parseLong(fov) / ROTATION_UNIT;
    }

    @Override
    public void setFieldOfViewAngle(double value) {
        Element cam = ensureCamera();
        cam.setAttribute("fov", String.valueOf(Math.round(value * ROTATION_UNIT)));
        save();
    }

    @Override
    public double getZoom() {
        Element cam = getCamera();
        if (cam == null) return 100.0;
        String val = cam.getAttribute("zoom");
        if (val == null || val.isEmpty()) return 100.0;
        return Long.parseLong(val) / 1000.0;
    }

    @Override
    public void setZoom(double value) {
        Element cam = ensureCamera();
        cam.setAttribute("zoom", String.valueOf(Math.round(value * 1000)));
        save();
    }

    @Override
    public void setRotation(double latitude, double longitude, double revolution) {
        Element cam = ensureCamera();
        Element rot = findChild(cam, "rot");
        if (rot == null) {
            Document doc = cam.getOwnerDocument();
            rot = doc.createElementNS(NS_A, "a:rot");
            cam.appendChild(rot);
        }
        rot.setAttribute("lat", String.valueOf(Math.round(latitude * ROTATION_UNIT)));
        rot.setAttribute("lon", String.valueOf(Math.round(longitude * ROTATION_UNIT)));
        rot.setAttribute("rev", String.valueOf(Math.round(revolution * ROTATION_UNIT)));
        save();
    }

    @Override
    public double[] getRotation() {
        Element cam = getCamera();
        if (cam == null) return new double[]{0.0, 0.0, 0.0};
        Element rot = findChild(cam, "rot");
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
