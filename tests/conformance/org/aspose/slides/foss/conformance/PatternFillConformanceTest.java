package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.FillType;
import org.aspose.slides.foss.IAutoShape;
import org.aspose.slides.foss.PatternStyle;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.ShapeType;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;
import org.w3c.dom.Element;

import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A pattern fill must name a pattern OOXML actually defines.
 *
 * <p>{@code ST_PresetPatternVal} is a closed list of 54 tokens. A {@code prst} outside it is not
 * an unknown pattern that a reader falls back on — it is invalid content, and PowerPoint
 * refuses the file. Deriving the token from the Java constant name produces
 * {@code diagonalBrick} where the standard says {@code diagBrick}, which looks right in a code
 * review and is wrong in every file.</p>
 */
class PatternFillConformanceTest {

    /** Every token ECMA-376 §20.1.10.50 permits as {@code a:pattFill/@prst}. */
    private static final Set<String> ST_PRESET_PATTERN_VAL = Set.of(
            "pct5", "pct10", "pct20", "pct25", "pct30", "pct40", "pct50", "pct60", "pct70",
            "pct75", "pct80", "pct90", "horz", "vert", "ltHorz", "ltVert", "dkHorz", "dkVert",
            "narHorz", "narVert", "dashHorz", "dashVert", "cross", "dnDiag", "upDiag",
            "ltDnDiag", "ltUpDiag", "dkDnDiag", "dkUpDiag", "wdDnDiag", "wdUpDiag",
            "dashDnDiag", "dashUpDiag", "diagCross", "smCheck", "lgCheck", "smGrid", "lgGrid",
            "dotGrid", "smConfetti", "lgConfetti", "horzBrick", "diagBrick", "solidDmnd",
            "openDmnd", "dotDmnd", "plaid", "sphere", "weave", "divot", "shingle", "wave",
            "trellis", "zigZag");

    @TempDir
    Path tempDir;

    /**
     * Every pattern style that reaches the file names a pattern the standard defines.
     *
     * <p>A style the library cannot express must write no {@code prst} at all; writing an
     * invented token is the one outcome that is never acceptable.</p>
     *
     * @param style the pattern style under test
     */
    @ParameterizedTest(name = "{0}")
    @EnumSource(PatternStyle.class)
    void aPatternFillMustNameAPatternTheStandardDefines(PatternStyle style) throws Exception {
        Path out = tempDir.resolve("pattern-"
                + style.name().toLowerCase(Locale.ROOT) + ".pptx");
        try (var pres = new Presentation()) {
            IAutoShape shape = pres.getSlides().get(0).getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 200, 100);
            shape.getFillFormat().setFillType(FillType.PATTERN);
            shape.getFillFormat().getPatternFormat().setPatternStyle(style);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<Element> fills = PackageAssertions.selectNodes(
                    pkg, "ppt/slides/slide1.xml", "//a:pattFill");
            if (fills.isEmpty()) {
                return;
            }
            Element fill = fills.get(0);
            if (!fill.hasAttribute("prst")) {
                return;
            }
            assertThat(fill.getAttribute("prst"))
                    .as("a:pattFill/@prst written for PatternStyle.%s", style)
                    .isIn(ST_PRESET_PATTERN_VAL);
        }
    }
}
