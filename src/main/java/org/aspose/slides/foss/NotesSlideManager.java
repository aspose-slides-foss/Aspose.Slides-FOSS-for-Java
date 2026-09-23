package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.opc.Relationship;
import org.aspose.slides.foss.internal.opc.RelationshipsManager;
import org.aspose.slides.foss.internal.pptx.NotesSlidePart;
import org.aspose.slides.foss.internal.pptx.OpcPackage;
import org.aspose.slides.foss.internal.pptx.RelsHelper;
import org.aspose.slides.foss.internal.pptx.SlidePart;

import java.util.List;
import java.util.Optional;

/**
 * Manages the notes slide for a given slide.
 */
public final class NotesSlideManager implements INotesSlideManager {

    private static final String NOTES_SLIDE_REL_TYPE =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

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

    /**
     * {@inheritDoc}
     *
     * <p>The notes slide is written with the relationships a reader needs to place it: from
     * the slide to the notes slide, back from the notes slide to the slide, and from the notes
     * slide to the presentation's notes master, which is created if the presentation has
     * none.</p>
     */
    @Override
    public INotesSlide addNotesSlide() {
        if (notesSlide != null) {
            return notesSlide;
        }
        String slidePartUri = slide.getSlidePartUri();
        var notesPart = NotesSlidePart.createEmpty(pkg, slidePartUri);

        var slideRels = new RelsHelper(pkg, slidePartUri);
        slideRels.addRelationship(NOTES_SLIDE_REL_TYPE,
                SlidePart.computeRelativeTarget(slidePartUri, notesPart.getPartName()));
        slideRels.save();

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
        pkg.removePartCascading(notesSlide.getNotesPart().getPartName(), slide.getSlidePartUri());
        notesSlide = null;
    }

    private String getNotesPartUri() {
        return "ppt/notesSlides/notesSlide" + (slide.getIndex() + 1) + ".xml";
    }
}
