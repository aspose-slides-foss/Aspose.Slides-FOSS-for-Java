package org.aspose.slides.foss;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilderFactory;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for {@link Camera}: camera type, FOV, zoom, and rotation.
 *
 * <p>Verifies camera type, field of view, zoom, and rotation behavior.</p>
 */
class CameraTest {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private Document doc;
    private Element scene3d;

    @BeforeEach
    void setUp() throws Exception {
        doc = DocumentBuilderFactory.newInstance()
                .newDocumentBuilder()
                .getDOMImplementation()
                .createDocument(NS_A, "a:scene3d", null);
        scene3d = doc.getDocumentElement();
    }

    private Camera createCamera() {
        return new Camera(scene3d, null);
    }

    // --- camera type ---

    @Test
    void cameraType_defaultIsNotDefined() {
        var cam = createCamera();
        assertThat(cam.getCameraType()).isEqualTo(CameraPresetType.NOT_DEFINED);
    }

    @Test
    void cameraType_perspectiveAbove_persists() {
        var cam = createCamera();
        cam.setCameraType(CameraPresetType.PERSPECTIVE_ABOVE);
        assertThat(cam.getCameraType()).isEqualTo(CameraPresetType.PERSPECTIVE_ABOVE);
    }

    @Test
    void cameraType_persistsAcrossReRead() {
        var cam = createCamera();
        cam.setCameraType(CameraPresetType.PERSPECTIVE_ABOVE);

        // Simulate re-reading from the same XML (as save/reload would)
        var cam2 = new Camera(scene3d, null);
        assertThat(cam2.getCameraType()).isEqualTo(CameraPresetType.PERSPECTIVE_ABOVE);
    }

    /**
     * {@code prst} is required on {@code CT_Camera}, so clearing the camera type has to
     * restore the schema default; dropping the attribute leaves content a validating
     * reader rejects. This test previously asserted the removal.
     */
    @Test
    void cameraType_notDefined_restoresTheDefaultPresetAndKeepsTheRequiredAttribute() {
        var cam = createCamera();
        cam.setCameraType(CameraPresetType.PERSPECTIVE_FRONT);
        cam.setCameraType(CameraPresetType.NOT_DEFINED);

        var camera = (Element) scene3d.getElementsByTagNameNS(NS_A, "camera").item(0);
        assertThat(camera.hasAttribute("prst")).isTrue();
        assertThat(camera.getAttribute("prst")).isEqualTo("orthographicFront");
        assertThat(cam.getCameraType()).isEqualTo(CameraPresetType.ORTHOGRAPHIC_FRONT);
    }

    static Stream<CameraPresetType> cameraPresetTypes() {
        return Stream.of(
                CameraPresetType.ISOMETRIC_BOTTOM_DOWN, CameraPresetType.ISOMETRIC_TOP_UP,
                CameraPresetType.LEGACY_OBLIQUE_FRONT, CameraPresetType.LEGACY_PERSPECTIVE_FRONT,
                CameraPresetType.OBLIQUE_TOP, CameraPresetType.ORTHOGRAPHIC_FRONT,
                CameraPresetType.PERSPECTIVE_ABOVE, CameraPresetType.PERSPECTIVE_FRONT,
                CameraPresetType.PERSPECTIVE_LEFT, CameraPresetType.PERSPECTIVE_RIGHT,
                CameraPresetType.PERSPECTIVE_RELAXED, CameraPresetType.PERSPECTIVE_HEROIC_LEFT_FACING,
                CameraPresetType.ISOMETRIC_OFF_AXIS_1_LEFT, CameraPresetType.ISOMETRIC_OFF_AXIS_2_RIGHT);
    }

    @ParameterizedTest
    @MethodSource("cameraPresetTypes")
    void cameraType_allPresetsRoundTrip(CameraPresetType type) {
        var cam = createCamera();
        cam.setCameraType(type);
        assertThat(cam.getCameraType()).isEqualTo(type);
    }

    // --- field of view angle ---

    @Test
    void fieldOfViewAngle_defaultIsZero() {
        var cam = createCamera();
        assertThat(cam.getFieldOfViewAngle()).isEqualTo(0.0);
    }

    @Test
    void fieldOfViewAngle_setAndGet() {
        var cam = createCamera();
        cam.setFieldOfViewAngle(90.0);
        assertThat(cam.getFieldOfViewAngle()).isCloseTo(90.0, org.assertj.core.data.Offset.offset(0.001));
    }

    // --- zoom ---

    @Test
    void zoom_defaultIs100() {
        var cam = createCamera();
        assertThat(cam.getZoom()).isEqualTo(100.0);
    }

    @Test
    void zoom_setAndGet() {
        var cam = createCamera();
        cam.setZoom(150.0);
        assertThat(cam.getZoom()).isCloseTo(150.0, org.assertj.core.data.Offset.offset(0.001));
    }

    // --- rotation ---

    @Test
    void rotation_defaultIsZeros() {
        var cam = createCamera();
        double[] rot = cam.getRotation();
        assertThat(rot).containsExactly(0.0, 0.0, 0.0);
    }

    @Test
    void rotation_setAndGet() {
        var cam = createCamera();
        cam.setRotation(45.0, 90.0, 180.0);
        double[] rot = cam.getRotation();
        assertThat(rot[0]).isCloseTo(45.0, org.assertj.core.data.Offset.offset(0.001));
        assertThat(rot[1]).isCloseTo(90.0, org.assertj.core.data.Offset.offset(0.001));
        assertThat(rot[2]).isCloseTo(180.0, org.assertj.core.data.Offset.offset(0.001));
    }

    // --- save callback ---

    @Test
    void setCameraType_invokesSaveCallback() {
        int[] callCount = {0};
        var cam = new Camera(scene3d, () -> callCount[0]++);
        cam.setCameraType(CameraPresetType.PERSPECTIVE_ABOVE);
        assertThat(callCount[0]).isEqualTo(1);
    }

    @Test
    void setFieldOfViewAngle_invokesSaveCallback() {
        int[] callCount = {0};
        var cam = new Camera(scene3d, () -> callCount[0]++);
        cam.setFieldOfViewAngle(45.0);
        assertThat(callCount[0]).isEqualTo(1);
    }

    @Test
    void setZoom_invokesSaveCallback() {
        int[] callCount = {0};
        var cam = new Camera(scene3d, () -> callCount[0]++);
        cam.setZoom(200.0);
        assertThat(callCount[0]).isEqualTo(1);
    }

    @Test
    void setRotation_invokesSaveCallback() {
        int[] callCount = {0};
        var cam = new Camera(scene3d, () -> callCount[0]++);
        cam.setRotation(10.0, 20.0, 30.0);
        assertThat(callCount[0]).isEqualTo(1);
    }

    // --- reads from existing XML ---

    @Test
    void readsExistingCameraElement() {
        Element camera = doc.createElementNS(NS_A, "a:camera");
        camera.setAttribute("prst", "perspectiveAbove");
        scene3d.appendChild(camera);

        var cam = createCamera();
        assertThat(cam.getCameraType()).isEqualTo(CameraPresetType.PERSPECTIVE_ABOVE);
    }
}
