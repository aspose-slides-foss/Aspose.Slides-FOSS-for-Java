package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.opc.Relationship;
import org.aspose.slides.foss.internal.opc.RelationshipsManager;
import org.aspose.slides.foss.internal.pptx.NotesSlidePart;
import org.aspose.slides.foss.internal.pptx.OpcPackage;
import org.aspose.slides.foss.internal.pptx.SlidePart;

import java.util.List;
import java.util.Optional;

/**
 * Manages the notes slide for a given slide.
 */
public final class NotesSlideManager implements INotesSlideManager {

    private final Slide slide;
    private final OpcPackage pkg;
    private NotesSlide notesSlide;

    /**
     * Creates a NotesSlideManager for the given slide.
     *
     * @param slide the parent slide
     * @param pkg   the OPC package
     */
    NotesSlideManager(Slide slide, OpcPackage pkg) {
        this.slide = slide;
        this.pkg = pkg;
        // Check if notes already exist
        String notesUri = getNotesPartUri();
        if (pkg.hasPart(notesUri)) {
            var notesPart = new NotesSlidePart(pkg, notesUri);
            this.notesSlide = new NotesSlide(slide, notesPart);
        }
    }

    /** The parent slide for relationship-based initialization (may be {@code null}). */
    private ISlide slideRef;
    /** OPC package for relationship-based initialization (may be {@code null}). */
    private org.aspose.slides.foss.internal.opc.OpcPackage opcPackage;
    /** Slide part name for relationship resolution (may be {@code null}). */
    private String slidePartName;

    /**
     * Internal initialization using OPC relationship resolution.
     *
     * <p>This is an alternative initialization path that resolves the notes slide
     * through the slide's OPC relationships rather than by convention-based URI.</p>
     *
     * @param slide         the parent slide
     * @param opcPackage    the OPC package (from {@code internal.opc})
     * @param slidePartName the slide part name (e.g. {@code "ppt/slides/slide1.xml"})
     */
    public void initInternal(ISlide slide, org.aspose.slides.foss.internal.opc.OpcPackage opcPackage,
                             String slidePartName) {
        this.slideRef = slide;
        this.opcPackage = opcPackage;
        this.slidePartName = slidePartName;
        this.notesSlide = null;
    }

    /**
     * Resolves the notes slide part name from the slide's OPC relationships.
     *
     * <p>Looks up the {@code notes_slide} relationship type in the slide part's
     * {@code .rels} file and resolves the relative target to an absolute part name.</p>
     *
     * @return the notes slide part name, or empty if no notes relationship exists
     */
    public Optional<String> getNotesPartName() {
        if (opcPackage == null || slidePartName == null) {
            return Optional.empty();
        }
        var relsManager = new RelationshipsManager(opcPackage, slidePartName);
        List<Relationship> rels = relsManager.getRelationshipsByType(
                RelationshipsManager.REL_TYPES.get("notes_slide"));
        if (!rels.isEmpty()) {
            String resolved = SlidePart.resolveTargetStatic(slidePartName, rels.getFirst().target());
            return Optional.of(resolved);
        }
        return Optional.empty();
    }

    /**
     * Loads and caches a {@link NotesSlide} from the given part name.
     *
     * <p>Creates a {@link NotesSlidePart} for the part, then initializes a new
     * {@link NotesSlide} via its {@link NotesSlide#initInternal initInternal} method.</p>
     *
     * @param partName the OPC part name of the notes slide
     * @return the loaded notes slide
     */
    public INotesSlide loadNotesSlide(String partName) {
        var notesPart = new NotesSlidePart(pkg, partName);
        var ns = new NotesSlide();
        ISlide parentSlide = slideRef != null ? slideRef : slide;
        IPresentation presentationRef = parentSlide != null ? parentSlide.getPresentation() : null;
        ns.initInternal(
                presentationRef,
                opcPackage,
                partName,
                notesPart,
                parentSlide
        );
        this.notesSlide = ns;
        return ns;
    }

    @Override
    public INotesSlide getNotesSlide() {
        if (notesSlide != null) {
            return notesSlide;
        }
        // Try relationship-based resolution if initInternal was called
        Optional<String> partName = getNotesPartName();
        if (partName.isPresent()) {
            return loadNotesSlide(partName.get());
        }
        return null;
    }

    @Override
    public INotesSlide addNotesSlide() {
        if (notesSlide != null) {
            return notesSlide;
        }
        String notesUri = getNotesPartUri();
        var notesPart = new NotesSlidePart(pkg, notesUri);
        int slideNumber = slide.getIndex() + 1;
        notesPart.ensurePartExists(slideNumber);

        // Add content type override
        addContentTypeOverride(slideNumber);

        // Add relationship from slide to notes
        addSlideNotesRelationship(slideNumber);

        notesSlide = new NotesSlide(slide, notesPart);
        return notesSlide;
    }

    @Override
    public void removeNotesSlide() {
        if (notesSlide == null) {
            return;
        }
        // Removing only the bytes would leave the slide's notesSlide relationship and the
        // content-type Override pointing at a part that is no longer there, which strict
        // readers and PowerPoint reject.
        pkg.removePartCascading(getNotesPartUri(), slide.getSlidePartUri());
        notesSlide = null;
    }

    private String getNotesPartUri() {
        return "ppt/notesSlides/notesSlide" + (slide.getIndex() + 1) + ".xml";
    }

    private void addContentTypeOverride(int slideNumber) {
        // Parse [Content_Types].xml and add override for notes slide
        var doc = pkg.parseXml("[Content_Types].xml");
        if (doc == null) return;

        var root = doc.getDocumentElement();
        String partName = "/ppt/notesSlides/notesSlide" + slideNumber + ".xml";

        // Check if already exists
        var overrides = root.getElementsByTagName("Override");
        for (int i = 0; i < overrides.getLength(); i++) {
            var el = (org.w3c.dom.Element) overrides.item(i);
            if (partName.equals(el.getAttribute("PartName"))) {
                return;
            }
        }

        var override = doc.createElement("Override");
        override.setAttribute("PartName", partName);
        override.setAttribute("ContentType",
                "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml");
        root.appendChild(override);
        pkg.serializeXml("[Content_Types].xml", doc);
    }

    private void addSlideNotesRelationship(int slideNumber) {
        String relsUri = "ppt/slides/_rels/slide" + slideNumber + ".xml.rels";
        String type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
        String relNs = "http://schemas.openxmlformats.org/package/2006/relationships";

        var doc = pkg.parseXml(relsUri);
        if (doc == null) {
            doc = OpcPackage.newDocument();
            var root = doc.createElementNS(relNs, "Relationships");
            doc.appendChild(root);
            var rel = doc.createElementNS(relNs, "Relationship");
            rel.setAttribute("Id", "rId_notes");
            rel.setAttribute("Type", type);
            rel.setAttribute("Target", "../notesSlides/notesSlide" + slideNumber + ".xml");
            root.appendChild(rel);
            pkg.serializeXml(relsUri, doc);
            return;
        }

        var root = doc.getDocumentElement();
        var rels = root.getElementsByTagName("Relationship");
        for (int i = 0; i < rels.getLength(); i++) {
            var rel = (org.w3c.dom.Element) rels.item(i);
            if (type.equals(rel.getAttribute("Type"))) return;
        }
        var rel = doc.createElementNS(relNs, "Relationship");
        rel.setAttribute("Id", "rId_notes");
        rel.setAttribute("Type", type);
        rel.setAttribute("Target", "../notesSlides/notesSlide" + slideNumber + ".xml");
        root.appendChild(rel);
        pkg.serializeXml(relsUri, doc);
    }
}
