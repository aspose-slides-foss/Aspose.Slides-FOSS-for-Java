package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.ISlide;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Element;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * What refers to a slide must follow the slide: its section entry when the slide is removed,
 * and its notes' reference back to it when the slide is cloned.
 *
 * <p>Sections live in a {@code p14:sectionLst} extension of {@code ppt/presentation.xml} and
 * name their slides by {@code p:sldId/@id}. A removed slide whose id stays in a section leaves
 * the list naming a slide the presentation no longer has.</p>
 *
 * <p>A notes slide relates back to the slide it annotates. A clone's notes that point back at
 * the source slide belong, as far as any reader can tell, to the source.</p>
 */
class SlideDependentsConformanceTest {

    static final String REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    static final String REL_NOTES_SLIDE = REL_BASE + "notesSlide";
    static final String REL_SLIDE = REL_BASE + "slide";
    static final String REL_NOTES_MASTER = REL_BASE + "notesMaster";

    static final String SECTIONS_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
    static final String NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    @TempDir
    Path tempDir;

    // ---------------------------------------------------------------- sections

    /** Removing a slide must remove its id from the section that lists it. */
    @Test
    void aRemovedSlideMustLeaveItsSection() throws Exception {
        Path deck = deckWithSections();
        List<String> ids = slideIds(deck);
        Path out = tempDir.resolve("section-slide-removed.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.getSlides().removeAt(1);
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        Map<String, List<String>> expected = new LinkedHashMap<>();
        expected.put("Intro", List.of(ids.get(0)));
        expected.put("Body", List.of(ids.get(2)));
        assertSections(out, expected);
    }

    /** A subset save must not list the slides it left out in any section. */
    @Test
    void aSubsetSaveMustNotListTheSlidesItLeftOut() throws Exception {
        Path deck = deckWithSections();
        List<String> ids = slideIds(deck);
        Path out = tempDir.resolve("section-subset.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.save(out.toString(), new int[]{0}, SaveFormat.PPTX);
        }

        Map<String, List<String>> expected = new LinkedHashMap<>();
        expected.put("Intro", List.of(ids.get(0)));
        expected.put("Body", List.of());
        assertSections(out, expected);
    }

    // ---------------------------------------------------------------- cloned notes

    /** A clone made in the same session must get notes of its own that point back at it. */
    @Test
    void aClonedSlidesNotesMustPointBackAtTheClone() throws Exception {
        Path out = tempDir.resolve("clone-notes.pptx");
        try (var pres = new Presentation()) {
            ISlide source = pres.getSlides().get(0);
            source.getNotesSlideManager().addNotesSlide().getNotesTextFrame().setText("Source notes");
            pres.getSlides().addClone(source);
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        assertEachSlideHasItsOwnNotes(out, "Source notes", "Source notes");
    }

    /** Cloning a slide of a deck another producer wrote must give the clone its own notes. */
    @Test
    void aClonedLoadedSlidesNotesMustPointBackAtTheClone() throws Exception {
        Path deck = Fixtures.authoredDeckWithNotes(tempDir, "notes.pptx",
                new String[]{"First", "Second"}, new String[]{"First notes", "Second notes"});
        Path out = tempDir.resolve("clone-loaded-notes.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.getSlides().addClone(pres.getSlides().get(0));
            pres.save(out.toString(), SaveFormat.PPTX);
        }
        assertEachSlideHasItsOwnNotes(out, "First notes", "Second notes", "First notes");
    }

    /** Cloning into another presentation must give the clone's notes a notes master there. */
    @Test
    void aSlideClonedIntoAnotherDeckMustHaveNotesThatResolveThere() throws Exception {
        Path deck = Fixtures.authoredDeckWithNotes(tempDir, "source.pptx",
                new String[]{"First"}, new String[]{"Carried notes"});
        Path out = tempDir.resolve("clone-across.pptx");
        try (var source = new Presentation(deck.toString());
             var target = new Presentation()) {
            target.getSlides().addClone(source.getSlides().get(0));
            target.save(out.toString(), SaveFormat.PPTX);
        }
        assertEachSlideHasItsOwnNotes(out, null, "Carried notes");
    }

    // ---------------------------------------------------------------- helpers

    /**
     * Asserts that slide {@code i} has notes with text {@code expected[i]} ({@code null} for
     * none), that each notes part points back at its own slide and at a notes master in the
     * package, and that no two slides share notes.
     */
    private static void assertEachSlideHasItsOwnNotes(Path out, String... expected)
            throws IOException {
        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            assertThat(slides).as("slides in %s", out).hasSize(expected.length);
            List<String> seen = new ArrayList<>();
            for (int i = 0; i < slides.size(); i++) {
                String slide = slides.get(i);
                List<String> notes = PackageAssertions.relatedParts(pkg, slide, REL_NOTES_SLIDE);
                if (expected[i] == null) {
                    assertThat(notes).as("notes of slide %d in %s", i + 1, out).isEmpty();
                    continue;
                }
                assertThat(notes).as("notes of slide %d in %s", i + 1, out).hasSize(1);
                String part = notes.get(0);
                assertThat(seen).as("notes parts already used by earlier slides in %s", out)
                        .doesNotContain(part);
                seen.add(part);
                assertThat(PackageAssertions.contentTypeOverrides(pkg).get("/" + part))
                        .as("content type of %s in %s", part, out)
                        .isEqualTo(NotesMasterConformanceTest.CT_NOTES_SLIDE);
                assertThat(PackageAssertions.relatedParts(pkg, part, REL_SLIDE))
                        .as("slide that %s points back to in %s%n%s", part, out,
                                pkg.text(PptxPackage.relsPartNameFor(part)))
                        .containsExactly(slide);
                List<String> masters = PackageAssertions.relatedParts(pkg, part, REL_NOTES_MASTER);
                assertThat(masters).as("notes master of %s in %s", part, out).hasSize(1);
                assertThat(PackageAssertions.selectNodes(pkg, "ppt/presentation.xml",
                                "/p:presentation/p:notesMasterIdLst/p:notesMasterId").stream()
                        .map(id -> {
                            try {
                                return PackageAssertions.relatedPartById(pkg,
                                        "ppt/presentation.xml",
                                        id.getAttributeNS(PptxPackage.NS_REL, "id"));
                            } catch (IOException e) {
                                throw new java.io.UncheckedIOException(e);
                            }
                        }).toList())
                        .as("notes masters registered in ppt/presentation.xml of %s%n%s", out,
                                pkg.text("ppt/presentation.xml"))
                        .containsExactly(masters.get(0));
                assertThat(NotesMasterConformanceTest.notesBodyText(pkg, part))
                        .as("notes text of slide %d in %s", i + 1, out)
                        .isEqualTo(expected[i]);
            }
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /** Three slides in two sections, Intro = [1] and Body = [2, 3], in PowerPoint's markup. */
    private Path deckWithSections() throws IOException {
        Path plain = Fixtures.authoredDeck(tempDir, "plain.pptx", "One", "Two", "Three");
        List<String> ids = slideIds(plain);
        String sections = "<p:extLst><p:ext uri=\"" + SECTIONS_EXT_URI + "\">"
                + "<p14:sectionLst xmlns:p14=\"" + NS_P14 + "\">"
                + "<p14:section name=\"Intro\" id=\"{8AE94259-FE10-46CC-94FC-F190558C5500}\">"
                + "<p14:sldIdLst><p14:sldId id=\"" + ids.get(0) + "\"/></p14:sldIdLst>"
                + "</p14:section>"
                + "<p14:section name=\"Body\" id=\"{C9CAF6C8-0E1A-4FF7-B237-4A839E46D10A}\">"
                + "<p14:sldIdLst><p14:sldId id=\"" + ids.get(1) + "\"/>"
                + "<p14:sldId id=\"" + ids.get(2) + "\"/></p14:sldIdLst>"
                + "</p14:section></p14:sectionLst></p:ext></p:extLst>";
        Path deck = ZipSurgery.copyWithReplacement(plain, tempDir.resolve("sections.pptx"),
                "ppt/presentation.xml",
                xml -> xml.replace("</p:presentation>", sections + "</p:presentation>"));
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            PackageAssertions.assertChildrenInSchemaOrder(pkg.root("ppt/presentation.xml"),
                    NotesMasterConformanceTest.PRESENTATION_ORDER);
        }
        return deck;
    }

    private static List<String> slideIds(Path deck) throws IOException {
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            return PackageAssertions.selectNodes(pkg, "ppt/presentation.xml",
                    "//p:sldIdLst/p:sldId").stream().map(e -> e.getAttribute("id")).toList();
        }
    }

    /** Asserts the full section list, and that it names exactly the slides of the deck. */
    private static void assertSections(Path out, Map<String, List<String>> expected)
            throws IOException {
        try (PptxPackage pkg = PptxPackage.open(out)) {
            PackageAssertions.assertPackageIsSelfConsistent(pkg);
            String xml = pkg.text("ppt/presentation.xml");
            Map<String, List<String>> actual = new LinkedHashMap<>();
            List<String> listed = new ArrayList<>();
            for (Element section : PackageAssertions.selectNodes(pkg, "ppt/presentation.xml",
                    "//p14:sectionLst/p14:section")) {
                List<String> ids = new ArrayList<>();
                var nodes = section.getElementsByTagNameNS(NS_P14, "sldId");
                for (int i = 0; i < nodes.getLength(); i++) {
                    ids.add(((Element) nodes.item(i)).getAttribute("id"));
                }
                actual.put(section.getAttribute("name"), ids);
                listed.addAll(ids);
            }
            assertThat(actual).as("sections in %s%n%s", out, xml)
                    .containsExactlyEntriesOf(expected);
            assertThat(listed).as("slide ids the sections list in %s%n%s", out, xml)
                    .containsExactlyElementsOf(PackageAssertions.selectNodes(pkg,
                            "ppt/presentation.xml", "//p:sldIdLst/p:sldId").stream()
                            .map(e -> e.getAttribute("id")).toList());
        }
        ThirdPartyReadBack.assertOpens(out);
    }
}
