package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.CameraPresetType;
import org.aspose.slides.foss.IAutoShape;
import org.aspose.slides.foss.LightRigPresetType;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.ShapeType;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;
import org.w3c.dom.Element;

import java.nio.file.Path;
import java.util.List;
import java.util.Locale;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A 3-D scene must be complete and must be written where the schema puts it.
 *
 * <p>{@code CT_Scene3D} requires both a camera and a light rig; either one alone is invalid
 * content, not a partial scene. And {@code CT_ShapeProperties} orders its children
 * {@code xfrm, geometry, fill, ln, effect, scene3d, sp3d, extLst} — a {@code scene3d} appended
 * after an {@code sp3d} that already exists is out of sequence, so whether the file is valid
 * depends on which property the caller happened to set first.</p>
 */
class ThreeDConformanceTest {

    @TempDir
    Path tempDir;

    /**
     * A camera on its own must still produce a complete scene.
     *
     * @param camera the camera preset under test
     */
    @ParameterizedTest(name = "{0}")
    @EnumSource(CameraPresetType.class)
    void aCameraMustBeWrittenWithTheLightRigTheSchemaRequires(CameraPresetType camera)
            throws Exception {
        Path out = tempDir.resolve("camera-" + camera.name().toLowerCase(Locale.ROOT) + ".pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 200, 100);
            shape.getThreeDFormat().getCamera().setCameraType(camera);
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        assertSceneIsComplete(out);
    }

    /**
     * A light rig on its own must still produce a complete scene.
     *
     * @param lightRig the light rig preset under test
     */
    @ParameterizedTest(name = "{0}")
    @EnumSource(LightRigPresetType.class)
    void aLightRigMustBeWrittenWithTheCameraTheSchemaRequires(LightRigPresetType lightRig)
            throws Exception {
        Path out = tempDir.resolve("lightrig-"
                + lightRig.name().toLowerCase(Locale.ROOT) + ".pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 200, 100);
            shape.getThreeDFormat().getLightRig().setLightType(lightRig);
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        assertSceneIsComplete(out);
    }

    /**
     * The scene must be written before the shape 3-D properties, whichever order the caller
     * set them in.
     */
    @Test
    void theSceneMustPrecedeTheShapeThreeDPropertiesWhicheverOrderTheyWereSetIn()
            throws Exception {
        Path out = tempDir.resolve("threed-order.pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 200, 100);
            // Extrusion first: it creates <a:sp3d>, which <a:scene3d> must then be placed before.
            shape.getThreeDFormat().setExtrusionHeight(20);
            shape.getThreeDFormat().getCamera().setCameraType(CameraPresetType.PERSPECTIVE_FRONT);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            Element shapeProperties = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:sp/p:spPr").get(0);
            List<String> children = PackageAssertions.childNames(shapeProperties);
            assertThat(children).as("children of p:spPr in %s", out)
                    .contains("a:scene3d", "a:sp3d");
            assertThat(children.indexOf("a:scene3d"))
                    .as("a:scene3d must come before a:sp3d; p:spPr children are %s in %s",
                            children, out)
                    .isLessThan(children.indexOf("a:sp3d"));
        }
    }

    private static void assertSceneIsComplete(Path pptx) throws Exception {
        try (PptxPackage pkg = PptxPackage.open(pptx)) {
            List<Element> scenes = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//p:sp/p:spPr/a:scene3d");
            assertThat(scenes).as("a:scene3d elements in %s", pptx).hasSize(1);
            PackageAssertions.assertChildrenInSchemaOrder(
                    scenes.get(0), "a:camera", "a:lightRig", "a:backdrop", "a:extLst");
            assertThat(PackageAssertions.childNames(scenes.get(0)))
                    .as("a:scene3d requires both a camera and a light rig, in %s", pptx)
                    .contains("a:camera", "a:lightRig");

            Element camera = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//a:scene3d/a:camera").get(0);
            assertThat(camera.hasAttribute("prst"))
                    .as("a:camera/@prst is required, in %s", pptx)
                    .isTrue();

            Element lightRig = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//a:scene3d/a:lightRig").get(0);
            assertThat(lightRig.hasAttribute("rig"))
                    .as("a:lightRig/@rig is required, in %s", pptx)
                    .isTrue();
            assertThat(lightRig.hasAttribute("dir"))
                    .as("a:lightRig/@dir is required, in %s", pptx)
                    .isTrue();
        }
    }
}
