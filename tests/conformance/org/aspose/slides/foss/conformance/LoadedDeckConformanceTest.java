package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.ILayoutSlide;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.SlideLayoutType;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;
import org.w3c.dom.Element;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Opening a deck must expose the masters and layouts the deck actually contains.
 *
 * <p>This is the one place where the library's own report is the thing under test rather than
 * the file it wrote — a reader defect is only visible by comparing what the reader says with
 * what the file says. So the expected numbers are taken from the package itself, by counting
 * the layout parts in the ZIP, and never from the library.</p>
 *
 * <p>What breaks when this is wrong: {@code addEmptySlide(pres.getLayoutSlides().get(0))} — the
 * form every documented example uses — hands the new slide a layout that is not in the
 * document, so placeholder inheritance and theme resolution have nothing to resolve
 * against.</p>
 */
class LoadedDeckConformanceTest {

    @TempDir
    Path tempDir;

    /** The layouts reported must be the layouts the file contains. */
    @Test
    void openingADeckMustExposeTheLayoutsItContains() throws Exception {
        Path deck = Fixtures.authoredDeck(tempDir, "real.pptx", "First", "Second");

        long layoutPartsInFile;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            layoutPartsInFile = pkg.entryNames().stream()
                    .filter(name -> name.matches("ppt/slideLayouts/slideLayout\\d+\\.xml"))
                    .count();
        }
        assertThat(layoutPartsInFile)
                .as("the fixture must have several layouts for this test to mean anything")
                .isEqualTo(Fixtures.LAYOUTS_IN_FIXTURE);

        try (var pres = new Presentation(deck.toString())) {
            assertThat(pres.getLayoutSlides().size())
                    .as("layouts reported for a deck whose package contains %d of them",
                            layoutPartsInFile)
                    .isEqualTo((int) layoutPartsInFile);
        }
    }

    /** The master reported must be the master the file contains. */
    @Test
    void openingADeckMustExposeTheMasterItContains() throws Exception {
        Path deck = Fixtures.authoredDeck(tempDir, "real.pptx", "First");

        long masterPartsInFile;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            masterPartsInFile = pkg.entryNames().stream()
                    .filter(name -> name.matches("ppt/slideMasters/slideMaster\\d+\\.xml"))
                    .count();
        }

        long layoutPartsInFile;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            layoutPartsInFile = pkg.entryNames().stream()
                    .filter(name -> name.matches("ppt/slideLayouts/slideLayout\\d+\\.xml"))
                    .count();
        }

        try (var pres = new Presentation(deck.toString())) {
            assertThat(pres.getMasters().size())
                    .as("masters reported for a deck whose package contains %d of them",
                            masterPartsInFile)
                    .isEqualTo((int) masterPartsInFile);
            assertThat(pres.getMasters().get(0).getLayoutSlides().size())
                    .as("layouts on the master of a deck whose package contains %d of them",
                            layoutPartsInFile)
                    .isEqualTo((int) layoutPartsInFile);
        }
    }

    /** A slide must report the layout it is actually related to. */
    @Test
    void aSlideMustReportTheLayoutItIsRelatedTo() throws Exception {
        Path deck = Fixtures.authoredDeck(tempDir, "real.pptx", "First");

        String layoutNameInFile;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            String target = PackageAssertions.selectNodes(pkg,
                            "ppt/slides/_rels/slide1.xml.rels",
                            "//rel:Relationship[contains(@Type,'slideLayout')]").get(0)
                    .getAttribute("Target");
            String layoutPart = PptxPackage.resolveTarget("ppt/slides/slide1.xml", target);
            layoutNameInFile = PackageAssertions
                    .selectNodes(pkg, layoutPart, "//p:cSld").get(0).getAttribute("name");
        }
        assertThat(layoutNameInFile)
                .as("the fixture's layout must be named for this test to mean anything")
                .isNotEmpty();

        try (var pres = new Presentation(deck.toString())) {
            assertThat(pres.getSlides().get(0).getLayoutSlide().getName())
                    .as("layout name reported for a slide related to a layout named '%s'",
                            layoutNameInFile)
                    .isEqualTo(layoutNameInFile);
        }
    }

    // ---------------------------------------------------------------- layout order and type

    private static final String MASTER = "ppt/slideMasters/slideMaster1.xml";
    private static final String MASTER_RELS = "ppt/slideMasters/_rels/slideMaster1.xml.rels";
    private static final String REL_SLIDE_LAYOUT =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private static final Pattern RELATIONSHIP = Pattern.compile("<Relationship\\b[^>]*/>");

    /**
     * The orders the master's relationships are swept through. The relationships file has no
     * order that means anything, and PowerPoint writes it in whatever order it likes, so the
     * layouts must come out the same whichever one it is.
     */
    static Stream<String> relationshipOrders() {
        return Stream.concat(Stream.of("as written", "reversed"),
                IntStream.range(1, 12).mapToObj(k -> "rotated by " + k));
    }

    /**
     * The layouts must be listed in the order the master lists them, in
     * {@code p:sldLayoutIdLst}, not in the order of its relationships file.
     *
     * <p>A reader test: the expected order is read out of the package first, by resolving each
     * {@code p:sldLayoutId/@r:id} and reading the layout's name.</p>
     */
    @ParameterizedTest(name = "master relationships {0}")
    @MethodSource("relationshipOrders")
    void layoutsMustBeListedInTheOrderTheMasterListsThem(String order) throws Exception {
        Path deck = deckWithMasterRelationships(order);
        List<String> expected = layoutNamesInMasterOrder(deck);
        assertThat(expected).as("layouts in the fixture").hasSize(Fixtures.LAYOUTS_IN_FIXTURE);

        try (var pres = new Presentation(deck.toString())) {
            assertThat(namesOf(pres.getLayoutSlides().asICollection()))
                    .as("getLayoutSlides() of a deck whose master lists %s", expected)
                    .containsExactlyElementsOf(expected);
            assertThat(namesOf(pres.getMasters().get(0).getLayoutSlides().asICollection()))
                    .as("the master's layouts, when it lists %s", expected)
                    .containsExactlyElementsOf(expected);
        }
    }

    /**
     * The documented pattern, {@code addEmptySlide(getLayoutSlides().get(0))}, must put the new
     * slide on the master's first layout.
     */
    @Test
    void theFirstLayoutMustBeTheFirstLayoutOfTheMaster() throws Exception {
        Path deck = deckWithMasterRelationships("reversed");
        String firstLayoutPart;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            firstLayoutPart = layoutPartsInMasterOrder(pkg).get(0);
        }

        Path out = tempDir.resolve("first-layout.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.getSlides().addEmptySlide(pres.getLayoutSlides().get(0));
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            assertThat(PackageAssertions.relatedParts(pkg, slides.get(slides.size() - 1),
                    REL_SLIDE_LAYOUT))
                    .as("layout of the slide added on getLayoutSlides().get(0) in %s", out)
                    .containsExactly(firstLayoutPart);
        }
    }

    /**
     * A layout whose type has a {@link SlideLayoutType} must report it, not {@code CUSTOM}.
     *
     * <p>A reader test over every layout of the fixture; {@code objTx}, PowerPoint's "Content
     * with Caption", is the one that used to fall through.</p>
     */
    @Test
    void aLayoutMustReportItsType() throws Exception {
        Path deck = Fixtures.authoredDeck(tempDir, "types.pptx", "First");
        Map<String, String> typeByName = new LinkedHashMap<>();
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            for (String part : layoutPartsInMasterOrder(pkg)) {
                String name = PackageAssertions.selectNodes(pkg, part, "/p:sldLayout/p:cSld")
                        .get(0).getAttribute("name");
                typeByName.put(name, pkg.root(part).getAttribute("type"));
            }
        }
        assertThat(typeByName.values()).as("layout types in the fixture").contains("objTx");

        try (var pres = new Presentation(deck.toString())) {
            for (ILayoutSlide layout : pres.getLayoutSlides().asICollection()) {
                String type = typeByName.get(layout.getName());
                if (type == null || type.isEmpty() || type.equals("cust")) {
                    continue;
                }
                assertThat(layout.getLayoutType())
                        .as("type reported for layout '%s', whose type in the file is %s",
                                layout.getName(), type)
                        .isNotEqualTo(SlideLayoutType.CUSTOM);
                if (type.equals("objTx")) {
                    assertThat(layout.getLayoutType())
                            .as("type reported for layout '%s' (objTx)", layout.getName())
                            .isEqualTo(SlideLayoutType.TITLE_OBJECT_AND_CAPTION);
                }
            }
        }
    }

    private Path deckWithMasterRelationships(String order) throws IOException {
        Path deck = Fixtures.authoredDeck(tempDir, "layouts.pptx", "First", "Second");
        if (order.equals("as written")) {
            return deck;
        }
        return ZipSurgery.copyWithReplacement(deck,
                tempDir.resolve("layouts-" + order.replace(' ', '-') + ".pptx"), MASTER_RELS,
                rels -> reorderRelationships(rels, order));
    }

    private static String reorderRelationships(String rels, String order) {
        Matcher matcher = RELATIONSHIP.matcher(rels);
        List<String> entries = new ArrayList<>();
        int start = -1;
        int end = -1;
        while (matcher.find()) {
            if (start < 0) {
                start = matcher.start();
            }
            end = matcher.end();
            entries.add(matcher.group());
        }
        if (order.equals("reversed")) {
            Collections.reverse(entries);
        } else {
            Collections.rotate(entries, Integer.parseInt(order.substring("rotated by ".length())));
        }
        return rels.substring(0, start) + String.join("", entries) + rels.substring(end);
    }

    private static List<String> layoutPartsInMasterOrder(PptxPackage pkg) throws IOException {
        List<String> parts = new ArrayList<>();
        for (Element id : PackageAssertions.selectNodes(pkg, MASTER,
                "/p:sldMaster/p:sldLayoutIdLst/p:sldLayoutId")) {
            parts.add(PackageAssertions.relatedPartById(pkg, MASTER,
                    id.getAttributeNS(PptxPackage.NS_REL, "id")));
        }
        return parts;
    }

    private static List<String> layoutNamesInMasterOrder(Path deck) throws IOException {
        List<String> names = new ArrayList<>();
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            for (String part : layoutPartsInMasterOrder(pkg)) {
                names.add(PackageAssertions.selectNodes(pkg, part, "/p:sldLayout/p:cSld").get(0)
                        .getAttribute("name"));
            }
        }
        return names;
    }

    private static List<String> namesOf(List<ILayoutSlide> layouts) {
        return layouts.stream().map(ILayoutSlide::getName).toList();
    }
}
