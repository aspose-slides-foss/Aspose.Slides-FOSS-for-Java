package org.aspose.slides.foss.internal.pptx;

import java.nio.charset.StandardCharsets;
import java.util.List;

/**
 * The notes master a package needs before it can carry notes slides.
 *
 * <p>ECMA-376 Part 1 §13.3.5 gives a notes slide exactly one implicit relationship to a notes
 * master, and §13.3.4 registers that master in {@code <p:notesMasterIdLst>} on the
 * presentation. PowerPoint reads notes text back without either, which is what makes the
 * omission easy to ship: nothing shows it until a consumer that checks the relationship graph
 * refuses the package, or a notes placeholder has nothing to inherit its formatting from.</p>
 */
public final class NotesMasterPart {

    /** The part name used for a notes master this class creates. */
    public static final String PART_NAME = "ppt/notesMasters/notesMaster1.xml";

    /** The relationship type from a presentation or a notes slide to a notes master. */
    public static final String REL_TYPE =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";

    private static final String THEME_REL_TYPE =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

    private static final String THEME_DIRECTORY = "ppt/theme/";

    private static final String NOTES_MASTER_XML =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<p:notesMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
            + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
            + "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
            + "<p:cSld><p:spTree>"
            + "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
            + "<p:grpSpPr/></p:spTree></p:cSld>"
            + "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" "
            + "accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" "
            + "accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>"
            + "<p:notesStyle/>"
            + "</p:notesMaster>";

    private NotesMasterPart() {
        // utility class
    }

    /**
     * Returns the notes master part name, creating the part and registering it if the package
     * does not have one yet.
     *
     * @param pkg the package to inspect and, if needed, extend
     * @return the part name of the package's notes master
     */
    public static String ensureInPackage(OpcPackage pkg) {
        var presRels = new RelsHelper(pkg, PresentationPart.PART_NAME);
        for (var rel : presRels.getAllRelationships()) {
            if (REL_TYPE.equals(rel.type()) && !"External".equals(rel.targetMode())) {
                String target = SlidePart.resolveTargetStatic(PresentationPart.PART_NAME,
                        rel.target());
                if (pkg.hasPart(target)) {
                    return target;
                }
                // Points at nothing; it is replaced below rather than left beside the new one.
                presRels.removeRelationship(rel.id());
            }
        }

        if (!pkg.hasPart(PART_NAME)) {
            pkg.setPartBytes(PART_NAME, NOTES_MASTER_XML.getBytes(StandardCharsets.UTF_8));
            var contentTypes = new ContentTypesManager(pkg);
            contentTypes.addOverride(PART_NAME, ContentTypesManager.CONTENT_TYPES.get("notes_master"));
            contentTypes.save();
            relateOwnTheme(pkg);
        }

        String relId = presRels.addRelationship(REL_TYPE,
                SlidePart.computeRelativeTarget(PresentationPart.PART_NAME, PART_NAME));
        presRels.save();
        var presPart = new PresentationPart(pkg);
        presPart.setNotesMasterReference(relId);
        presPart.save();
        return PART_NAME;
    }

    /**
     * Gives the notes master a theme part of its own, copied from the first theme the package
     * already has so the notes use the same fonts and colours.
     *
     * <p>A theme part belongs to exactly one master. A notes master related to the slide
     * master's theme gives a package that validates and that PowerPoint refuses to open — the
     * whole file, not only the notes. Decks PowerPoint writes carry one theme for the slide
     * master and another for the notes master, which is what this reproduces.</p>
     */
    private static void relateOwnTheme(OpcPackage pkg) {
        List<String> themes = pkg.getPartNames().stream()
                .filter(name -> name.startsWith(THEME_DIRECTORY) && name.endsWith(".xml")
                        && name.indexOf('/', THEME_DIRECTORY.length()) < 0)
                .sorted()
                .toList();
        if (themes.isEmpty()) {
            return;
        }
        String themePart = null;
        for (int index = 1; themePart == null; index++) {
            String candidate = THEME_DIRECTORY + "theme" + index + ".xml";
            if (!pkg.hasPart(candidate)) {
                themePart = candidate;
            }
        }
        pkg.setPartBytes(themePart, pkg.getPartBytes(themes.get(0)).clone());
        var contentTypes = new ContentTypesManager(pkg);
        contentTypes.addOverride(themePart, ContentTypesManager.CONTENT_TYPES.get("theme"));
        contentTypes.save();

        var rels = new RelsHelper(pkg, PART_NAME);
        rels.addRelationship(THEME_REL_TYPE, SlidePart.computeRelativeTarget(PART_NAME, themePart));
        rels.save();
    }
}
