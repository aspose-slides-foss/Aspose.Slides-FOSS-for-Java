package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Element;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A notes slide must be wired to its slide and to a notes master.
 *
 * <p>ECMA-376 Part 1 §13.3.5 gives a notes slide part an implicit relationship to exactly one
 * notes master part, and the notes master is registered from the presentation in
 * {@code p:notesMasterIdLst} (§13.3.4). Speaker notes used to be written as a bare
 * part with no {@code .rels} at all: no notes master, and not even the relationship back to the
 * slide they annotate. A reader that walks the relationship graph finds notes that belong to
 * nothing, and the notes placeholders have no master to inherit their formatting from.</p>
 *
 * <p>The notes master also needs a theme of its own. A theme part belongs to one master; a
 * notes master related to the slide master's theme is a package PowerPoint refuses to open.</p>
 */
class NotesMasterConformanceTest {

    static final String REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    static final String REL_SLIDE = REL_BASE + "slide";
    static final String REL_NOTES_SLIDE = REL_BASE + "notesSlide";
    static final String REL_NOTES_MASTER = REL_BASE + "notesMaster";
    static final String REL_THEME = REL_BASE + "theme";
    static final String REL_SLIDE_MASTER = REL_BASE + "slideMaster";

    static final String CT_NOTES_SLIDE =
            "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
    static final String CT_NOTES_MASTER =
            "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
    static final String CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";

    /** {@code CT_Presentation} children, in schema sequence order. */
    static final String[] PRESENTATION_ORDER = {
            "p:sldMasterIdLst", "p:notesMasterIdLst", "p:handoutMasterIdLst", "p:sldIdLst",
            "p:sldSz", "p:notesSz", "p:smartTags", "p:embeddedFontLst", "p:custShowLst",
            "p:photoAlbum", "p:custDataLst", "p:kinsoku", "p:defaultTextStyle",
            "p:modifyVerifier", "p:extLst",
    };

    /** {@code CT_NotesMaster} children, in schema sequence order. */
    static final String[] NOTES_MASTER_ORDER = {
            "p:cSld", "p:clrMap", "p:hf", "p:notesStyle", "p:extLst",
    };

    @TempDir
    Path tempDir;

    /** Notes added to a new deck must reach a notes master and their own slide. */
    @Test
    void notesAddedToANewDeckMustBeRelatedToANotesMasterAndToTheirSlide() throws Exception {
        Path out = tempDir.resolve("new-deck-notes.pptx");
        try (var pres = new Presentation()) {
            pres.getSlides().get(0).getNotesSlideManager().addNotesSlide()
                    .getNotesTextFrame().setText("Speaker notes");
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            String slide = PackageAssertions.slidePartNames(pkg).get(0);
            assertNotesWiring(pkg, slide, "Speaker notes");
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /** Notes added to a deck another producer wrote, with no notes master yet, must get one. */
    @Test
    void notesAddedToALoadedDeckWithoutANotesMasterMustGetOne() throws Exception {
        Path deck = Fixtures.authoredDeck(tempDir, "authored.pptx", "First", "Second");
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            assertThat(pkg.entryNames())
                    .as("the fixture must have no notes master for this test to mean anything")
                    .noneMatch(name -> name.startsWith("ppt/notesMasters/"));
        }

        Path out = tempDir.resolve("loaded-notes.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.getSlides().get(1).getNotesSlideManager().addNotesSlide()
                    .getNotesTextFrame().setText("Notes on the second slide");
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            String slide = PackageAssertions.slidePartNames(pkg).get(1);
            assertNotesWiring(pkg, slide, "Notes on the second slide");
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /** A deck that already has a notes master must keep that one, not gain a second. */
    @Test
    void notesAddedToADeckThatHasANotesMasterMustUseIt() throws Exception {
        Path deck = Fixtures.authoredDeckWithNotes(tempDir, "with-master.pptx",
                new String[]{"First", "Second"}, new String[]{"Existing notes", null});
        String existingMaster;
        try (PptxPackage pkg = PptxPackage.open(deck)) {
            String firstNotes = PackageAssertions.relatedParts(pkg,
                    PackageAssertions.slidePartNames(pkg).get(0), REL_NOTES_SLIDE).get(0);
            existingMaster = PackageAssertions.relatedParts(pkg, firstNotes, REL_NOTES_MASTER)
                    .get(0);
        }

        Path out = tempDir.resolve("second-notes.pptx");
        try (var pres = new Presentation(deck.toString())) {
            pres.getSlides().get(1).getNotesSlideManager().addNotesSlide()
                    .getNotesTextFrame().setText("New notes");
            pres.save(out.toString(), SaveFormat.PPTX);
        }

        try (PptxPackage pkg = PptxPackage.open(out)) {
            List<String> slides = PackageAssertions.slidePartNames(pkg);
            String masterOfNew = assertNotesWiring(pkg, slides.get(1), "New notes");
            String masterOfOld = assertNotesWiring(pkg, slides.get(0), "Existing notes");
            assertThat(masterOfNew)
                    .as("notes master of the added notes in %s", out)
                    .isEqualTo(existingMaster)
                    .isEqualTo(masterOfOld);
            assertThat(pkg.entryNames().stream()
                    .filter(name -> name.matches("ppt/notesMasters/notesMaster\\d+\\.xml")))
                    .as("notes master parts in %s", out)
                    .containsExactly(existingMaster);
        }
        ThirdPartyReadBack.assertOpens(out);
    }

    /**
     * Asserts every link a notes slide needs, reading only the package.
     *
     * @return the notes master part the notes slide is related to
     */
    static String assertNotesWiring(PptxPackage pkg, String slidePart, String expectedText)
            throws IOException {
        PackageAssertions.assertPackageIsSelfConsistent(pkg);

        List<String> notesParts = PackageAssertions.relatedParts(pkg, slidePart, REL_NOTES_SLIDE);
        assertThat(notesParts)
                .as("notesSlide relationships of %s in %s", slidePart, pkg.path())
                .hasSize(1);
        String notesPart = notesParts.get(0);
        assertThat(PackageAssertions.contentTypeOverrides(pkg).get("/" + notesPart))
                .as("content type of %s in %s", notesPart, pkg.path())
                .isEqualTo(CT_NOTES_SLIDE);
        assertThat(notesBodyText(pkg, notesPart))
                .as("notes body text of %s in %s", notesPart, pkg.path())
                .isEqualTo(expectedText);

        String notesRels = PptxPackage.relsPartNameFor(notesPart);
        assertThat(pkg.hasPart(notesRels))
                .as("%s must exist in %s; parts: %s", notesRels, pkg.path(), pkg.entryNames())
                .isTrue();
        assertThat(PackageAssertions.relatedParts(pkg, notesPart, REL_SLIDE))
                .as("slide relationships of %s in %s%n%s", notesPart, pkg.path(),
                        pkg.text(notesRels))
                .containsExactly(slidePart);
        List<String> masters = PackageAssertions.relatedParts(pkg, notesPart, REL_NOTES_MASTER);
        assertThat(masters)
                .as("notesMaster relationships of %s in %s%n%s", notesPart, pkg.path(),
                        pkg.text(notesRels))
                .hasSize(1);
        String master = masters.get(0);
        assertThat(PackageAssertions.contentTypeOverrides(pkg).get("/" + master))
                .as("content type of %s in %s", master, pkg.path())
                .isEqualTo(CT_NOTES_MASTER);
        Element masterRoot = pkg.root(master);
        assertThat(masterRoot.getTagName()).as("root of %s", master).isEqualTo("p:notesMaster");
        PackageAssertions.assertChildrenInSchemaOrder(masterRoot, NOTES_MASTER_ORDER);
        assertThat(PackageAssertions.childNames(masterRoot))
                .as("required children of %s", master)
                .contains("p:cSld", "p:clrMap");

        // Registered from the presentation, in p:notesMasterIdLst, in schema position.
        Element presentation = pkg.root("ppt/presentation.xml");
        PackageAssertions.assertChildrenInSchemaOrder(presentation, PRESENTATION_ORDER);
        List<Element> ids = PackageAssertions.selectNodes(pkg, "ppt/presentation.xml",
                "/p:presentation/p:notesMasterIdLst/p:notesMasterId");
        assertThat(ids)
                .as("p:notesMasterId entries in ppt/presentation.xml of %s%n%s",
                        pkg.path(), pkg.text("ppt/presentation.xml"))
                .hasSize(1);
        String relId = ids.get(0).getAttributeNS(PptxPackage.NS_REL, "id");
        assertThat(PackageAssertions.relatedPartById(pkg, "ppt/presentation.xml", relId))
                .as("target of p:notesMasterId r:id=%s in %s", relId, pkg.path())
                .isEqualTo(master);
        assertThat(PackageAssertions.relatedParts(pkg, "ppt/presentation.xml", REL_NOTES_MASTER))
                .as("notesMaster relationships of ppt/presentation.xml in %s", pkg.path())
                .containsExactly(master);

        // The notes master's theme is its own, not a slide master's.
        List<String> themes = PackageAssertions.relatedParts(pkg, master, REL_THEME);
        assertThat(themes).as("theme relationships of %s in %s", master, pkg.path()).hasSize(1);
        String theme = themes.get(0);
        assertThat(PackageAssertions.contentTypeOverrides(pkg).get("/" + theme))
                .as("content type of %s in %s", theme, pkg.path())
                .isEqualTo(CT_THEME);
        for (String slideMaster : PackageAssertions.relatedParts(pkg, "ppt/presentation.xml",
                REL_SLIDE_MASTER)) {
            assertThat(PackageAssertions.relatedParts(pkg, slideMaster, REL_THEME))
                    .as("theme of slide master %s must not be the notes master's theme in %s",
                            slideMaster, pkg.path())
                    .doesNotContain(theme);
        }
        return master;
    }

    /** The concatenated {@code a:t} text of the notes body placeholder of a notes part. */
    static String notesBodyText(PptxPackage pkg, String notesPart) throws IOException {
        StringBuilder text = new StringBuilder();
        for (Element t : PackageAssertions.selectNodes(pkg, notesPart,
                "//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='body']//a:t")) {
            text.append(t.getTextContent());
        }
        return text.toString();
    }
}
