package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.IAutoShape;
import org.aspose.slides.foss.IEffectFormat;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.ShapeType;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;
import org.w3c.dom.Element;

import java.nio.file.Path;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A one-line effect enabler must write an effect the schema accepts.
 *
 * <p>ECMA-376 §20.1.8 makes some children and attributes of the effect elements mandatory: a
 * shadow needs a colour, a soft edge needs a radius, a preset shadow needs a preset. An element
 * written without them is incomplete, and PowerPoint refuses the whole file rather than
 * ignoring the effect — so the cost of a missing default is the user's entire document.</p>
 */
class EffectElementConformanceTest {

    /** The colour choice group of DrawingML: any one of these satisfies an EG_ColorChoice. */
    private static final List<String> COLOR_CHOICE = List.of(
            "a:scrgbClr", "a:srgbClr", "a:hslClr", "a:sysClr", "a:schemeClr", "a:prstClr");

    /** The fill group of DrawingML: any one of these satisfies an EG_FillProperties. */
    private static final List<String> FILL_CHOICE = List.of(
            "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill");

    @TempDir
    Path tempDir;

    static Stream<Arguments> effects() {
        return Stream.of(
                effect("blur", "a:blur", c -> c.enableBlurEffect(), List.of(), List.of()),
                effect("fill overlay", "a:fillOverlay", c -> c.enableFillOverlayEffect(),
                        List.of("blend"), FILL_CHOICE),
                effect("glow", "a:glow", c -> c.enableGlowEffect(), List.of(), COLOR_CHOICE),
                effect("inner shadow", "a:innerShdw", c -> c.enableInnerShadowEffect(),
                        List.of(), COLOR_CHOICE),
                effect("outer shadow", "a:outerShdw", c -> c.enableOuterShadowEffect(),
                        List.of(), COLOR_CHOICE),
                effect("preset shadow", "a:prstShdw", c -> c.enablePresetShadowEffect(),
                        List.of("prst"), COLOR_CHOICE),
                effect("reflection", "a:reflection", c -> c.enableReflectionEffect(),
                        List.of(), List.of()),
                effect("soft edge", "a:softEdge", c -> c.enableSoftEdgeEffect(),
                        List.of("rad"), List.of()));
    }

    private static Arguments effect(String label, String tag, Consumer<IEffectFormat> enable,
                                    List<String> requiredAttributes,
                                    List<String> requiredChildGroup) {
        return Arguments.of(label, tag, enable, requiredAttributes, requiredChildGroup);
    }

    /**
     * Enabling an effect writes an element carrying everything the schema makes mandatory.
     *
     * @param label              the effect, for the test name
     * @param tag                the element the enabler is expected to write
     * @param enable             the one-line enabler under test
     * @param requiredAttributes attributes ECMA-376 marks required on that element
     * @param requiredChildGroup a choice group of which one member must be present, or empty
     */
    @ParameterizedTest(name = "{0}")
    @MethodSource("effects")
    void enablingAnEffectMustWriteACompleteElement(String label, String tag,
                                                   Consumer<IEffectFormat> enable,
                                                   List<String> requiredAttributes,
                                                   List<String> requiredChildGroup)
            throws Exception {
        Path out = tempDir.resolve(tag.replace(':', '-') + ".pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 200, 100);
            enable.accept(shape.getEffectFormat());
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<Element> effects = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//a:effectLst/" + tag);
            assertThat(effects)
                    .as("<%s> elements written by the %s enabler in %s", tag, label, out)
                    .hasSize(1);
            Element effect = effects.get(0);

            for (String attribute : requiredAttributes) {
                assertThat(effect.hasAttribute(attribute))
                        .as("<%s> must carry the required attribute '%s'; it was written as %s",
                                tag, attribute, describe(effect))
                        .isTrue();
            }

            if (!requiredChildGroup.isEmpty()) {
                assertThat(PackageAssertions.childNames(effect))
                        .as("<%s> must carry one of %s; it was written as %s",
                                tag, requiredChildGroup, describe(effect))
                        .containsAnyElementsOf(requiredChildGroup);
            }
        }
    }

    private static String describe(Element element) {
        StringBuilder text = new StringBuilder("<").append(element.getTagName());
        var attributes = element.getAttributes();
        for (int i = 0; i < attributes.getLength(); i++) {
            text.append(' ').append(attributes.item(i).getNodeName())
                    .append("=\"").append(attributes.item(i).getNodeValue()).append('"');
        }
        List<String> children = PackageAssertions.childNames(element);
        return children.isEmpty()
                ? text.append("/>").toString()
                : text.append('>').append(children).append("</").append(element.getTagName())
                .append('>').toString();
    }
}
