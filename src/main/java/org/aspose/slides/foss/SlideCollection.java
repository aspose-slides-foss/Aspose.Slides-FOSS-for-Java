package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.OpcPackage;
import org.aspose.slides.foss.internal.pptx.PresentationPart;
import org.aspose.slides.foss.internal.pptx.RelsHelper;
import org.aspose.slides.foss.internal.pptx.SlidePart;
import org.aspose.slides.foss.internal.pptx.SlideReference;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Represents a collection of slides in a presentation.
 */
public final class SlideCollection implements ISlideCollection {

    private static final Pattern SLIDE_FILE_PATTERN =
            Pattern.compile("ppt/slides/slide(\\d+)\\.xml$");

    private static final String REL_TYPE_SLIDE =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    private final List<Slide> slides = new ArrayList<>();
    private Presentation presentation;
    private OpcPackage opcPackage;
    private Function<String, LayoutSlide> layoutResolver;

    /**
     * Sets the owning presentation (called by {@link Presentation} during construction).
     *
     * @param presentation the owning presentation
     */
    void setPresentation(Presentation presentation) {
        this.presentation = presentation;
    }

    /**
     * Internal initialization for the slide collection.
     *
     * <p>Loads slides from OPC relationships in the presentation part, preserving
     * the ordering defined in {@code presentation.xml}.</p>
     *
     * @param presentation     the parent Presentation object
     * @param opcPackage       the OPC package
     * @param presentationPart the PresentationPart managing presentation.xml
     * @param layoutResolver   callable that resolves a layout part name to a LayoutSlide,
     *                         or {@code null}
     */
    public void initInternal(IPresentation presentation, OpcPackage opcPackage,
                             PresentationPart presentationPart,
                             Function<String, LayoutSlide> layoutResolver) {
        this.presentation = (Presentation) presentation;
        this.opcPackage = opcPackage;
        this.layoutResolver = layoutResolver;
        slides.clear();

        var presRels = new RelsHelper(opcPackage, PresentationPart.PART_NAME);

        for (SlideReference slideRef : presentationPart.slideReferences()) {
            RelsHelper.RelEntry rel = null;
            for (RelsHelper.RelEntry entry : presRels.getAllRelationships()) {
                if (entry.id().equals(slideRef.getRelId())) {
                    rel = entry;
                    break;
                }
            }
            if (rel == null) {
                continue;
            }

            // Resolve target path relative to ppt/
            String target = rel.target();
            String partName;
            if (target.startsWith("/")) {
                partName = target.replaceFirst("^/+", "");
            } else {
                partName = "ppt/" + target;
            }

            // Extract slide file number from part name to create Slide with correct index
            Matcher m = SLIDE_FILE_PATTERN.matcher(partName);
            if (m.matches()) {
                int fileNum = Integer.parseInt(m.group(1));
                Slide slide = new Slide(this.presentation, fileNum - 1);
                slides.add(slide);
            }
        }
    }

    @Override
    public ISlide get(int index) {
        return slides.get(index);
    }

    @Override
    public int size() {
        return slides.size();
    }

    @Override
    public List<ISlide> asICollection() {
        return Collections.unmodifiableList(new ArrayList<>(slides));
    }

    @Override
    public Iterable<ISlide> asIEnumerable() {
        return asICollection();
    }

    @Override
    public ISlide addClone(ISlide sourceSlide) {
        return cloneSlideInternal(sourceSlide, -1, null, null, false);
    }

    @Override
    public ISlide addClone(ISlide sourceSlide, ILayoutSlide destLayout) {
        return cloneSlideInternal(sourceSlide, -1, destLayout, null, false);
    }

    @Override
    public ISlide addClone(ISlide sourceSlide, IMasterSlide destMaster, boolean allowCloneMissingLayout) {
        return cloneSlideInternal(sourceSlide, -1, null, destMaster, allowCloneMissingLayout);
    }

    @Override
    public ISlide insertClone(int index, ISlide sourceSlide) {
        return cloneSlideInternal(sourceSlide, index, null, null, false);
    }

    @Override
    public ISlide insertClone(int index, ISlide sourceSlide, ILayoutSlide destLayout) {
        return cloneSlideInternal(sourceSlide, index, destLayout, null, false);
    }

    @Override
    public ISlide insertClone(int index, ISlide sourceSlide, IMasterSlide destMaster, boolean allowCloneMissingLayout) {
        return cloneSlideInternal(sourceSlide, index, null, destMaster, allowCloneMissingLayout);
    }

    @Override
    public ISlide[] toArray() {
        return slides.toArray(new ISlide[0]);
    }

    @Override
    public ISlide[] toArray(int startIndex, int count) {
        return slides.subList(startIndex, startIndex + count).toArray(new ISlide[0]);
    }

    @Override
    public ISlide addEmptySlide(ILayoutSlide layout) {
        return addEmptySlideInternal(layout, -1);
    }

    @Override
    public ISlide insertEmptySlide(int index, ILayoutSlide layout) {
        return addEmptySlideInternal(layout, index);
    }

    @Override
    public void remove(ISlide value) {
        int index = indexOf(value);
        if (index < 0) {
            return;
        }
        removeAt(index);
    }

    @Override
    public void removeAt(int index) {
        Slide slide = slides.get(index);
        unregisterSlide(slide);
        slides.remove(index);
    }

    /**
     * Deletes a slide from the package: its {@code p:sldId}, the presentation relationship
     * that reached it, its content-type {@code Override}, its part, its part's own
     * {@code .rels}, and anything the slide was the last reference to.
     *
     * @param slide the slide being removed from this collection
     */
    private void unregisterSlide(Slide slide) {
        String partName = slide.getSlidePartUri();
        if (partName == null) {
            return;
        }
        OpcPackage pkg = getEffectivePackage();
        var presRels = new RelsHelper(pkg, PresentationPart.PART_NAME);

        String relId = null;
        for (RelsHelper.RelEntry entry : presRels.getAllRelationships()) {
            if (REL_TYPE_SLIDE.equals(entry.type())
                    && partName.equals(SlidePart.resolveTargetStatic(
                            PresentationPart.PART_NAME, entry.target()))) {
                relId = entry.id();
                break;
            }
        }

        if (relId != null) {
            PresentationPart presPart = getEffectivePresentationPart();
            if (presPart.removeSlideReferenceByRelId(relId)) {
                presPart.save();
            }
        }

        pkg.removePartCascading(partName, PresentationPart.PART_NAME);
    }

    @Override
    public int indexOf(ISlide slide) {
        if (slide instanceof Slide s) {
            for (int i = 0; i < slides.size(); i++) {
                if (slides.get(i).getSlideId() == s.getSlideId()) {
                    return i;
                }
            }
            return -1;
        }
        return slides.indexOf(slide);
    }

    @Override
    public Iterator<ISlide> iterator() {
        return new ArrayList<ISlide>(slides).iterator();
    }

    /**
     * Adds a slide to the collection (package-private, used during loading).
     *
     * @param slide the slide to add
     */
    void add(Slide slide) {
        slides.add(slide);
    }

    /**
     * Returns the internal list of slides (for serialization).
     *
     * @return the backing list
     */
    List<Slide> getInternalList() {
        return slides;
    }

    // ── Internal methods ──────────────────────────────────────────────

    /**
     * Internal implementation for adding or inserting an empty slide.
     *
     * <p>Creates the slide XML part in the OPC package, establishes the layout
     * relationship, registers the slide in {@code presentation.xml}, and adds
     * a {@link Slide} object to this collection.</p>
     *
     * @param layout the layout slide to use for the new slide
     * @param index  position to insert at; {@code -1} means append at end
     * @return the newly created slide
     */
    public ISlide addEmptySlideInternal(ILayoutSlide layout, int index) {
        OpcPackage pkg = getEffectivePackage();
        PresentationPart presPart = getEffectivePresentationPart();

        // Determine the layout part name. A slide must relate to a layout; the layout objects
        // the presentation exposes are not always backed by a part, so fall back to a real one.
        String layoutPartName = resolveLayoutPartName(layout);
        if (layoutPartName == null) {
            layoutPartName = getFirstLayoutPartName().orElseThrow(() -> new IllegalStateException(
                    "Cannot add a slide: the presentation has no slide layout part to relate it to"));
        }

        // Determine the next available slide file number
        int nextNum = getNextSlideFileNumber();
        String partName = "ppt/slides/slide" + nextNum + ".xml";

        // Create the empty slide XML + rels + content type
        SlidePart.createEmpty(pkg, partName, layoutPartName);

        // Add relationship from presentation to the new slide
        var presRels = new RelsHelper(pkg, PresentationPart.PART_NAME);
        String relativeTarget = "slides/slide" + nextNum + ".xml";
        String relId = presRels.addRelationship(REL_TYPE_SLIDE, relativeTarget);
        presRels.save();

        // Add slide reference to presentation.xml
        presPart.addSlideReference(relId, null, index);
        presPart.save();

        // Create Slide object
        Slide slide = new Slide(presentation, nextNum - 1);
        slide.setLayoutSlide(layout);

        // Add to internal list
        if (index < 0 || index >= slides.size()) {
            slides.add(slide);
        } else {
            slides.add(index, slide);
        }

        return slide;
    }

    /**
     * Finds the next available slide file number by scanning existing parts.
     *
     * @return the smallest positive integer {@code n} such that
     *         {@code ppt/slides/slideN.xml} does not exist in the package
     */
    public int getNextSlideFileNumber() {
        OpcPackage pkg = getEffectivePackage();
        Set<Integer> existingNums = new HashSet<>();
        for (String partName : pkg.getPartNames()) {
            Matcher m = SLIDE_FILE_PATTERN.matcher(partName);
            if (m.matches()) {
                existingNums.add(Integer.parseInt(m.group(1)));
            }
        }
        int num = 1;
        while (existingNums.contains(num)) {
            num++;
        }
        return num;
    }

    /**
     * Internal implementation for cloning a slide.
     *
     * <p>Copies the slide XML and its related resources into the destination
     * package, re-pointing the layout relationship as needed. When cloning
     * across presentations the master slide chain is cloned automatically.</p>
     *
     * @param sourceSlide              the slide to clone
     * @param index                    position to insert at; {@code -1} means append at end
     * @param destLayout               optional destination layout slide
     * @param destMaster               optional destination master slide
     * @param allowCloneMissingLayout  if {@code true} and layout not found, use
     *                                 the first available layout as a fallback
     * @return the cloned slide
     */
    public ISlide cloneSlideInternal(ISlide sourceSlide, int index,
                                     ILayoutSlide destLayout, IMasterSlide destMaster,
                                     boolean allowCloneMissingLayout) {
        OpcPackage pkg = getEffectivePackage();

        // Get source package and part name. The clone is made from the source slide's bytes,
        // so any edit still held in its DOM has to reach the package first.
        Slide sourceSlideImpl = (Slide) sourceSlide;
        sourceSlideImpl.flush();
        OpcPackage sourcePackage = ((Presentation) sourceSlideImpl.getPresentation()).getPackage();
        String sourcePartName = sourceSlideImpl.getSlidePartUri();

        // Determine the next available slide number
        int nextNum = getNextSlideFileNumber();
        String destPartName = "ppt/slides/slide" + nextNum + ".xml";

        // Determine the layout to use
        String destLayoutPartName = null;
        if (destLayout != null) {
            destLayoutPartName = resolveLayoutPartName(destLayout);
        } else if (destMaster != null) {
            destLayoutPartName = findMatchingLayout(sourceSlide, destMaster, allowCloneMissingLayout)
                    .orElse(null);
        } else {
            ILayoutSlide sourceLayout = sourceSlide.getLayoutSlide();
            if (sourcePackage == pkg) {
                // Same presentation — use the same layout
                destLayoutPartName = resolveLayoutPartName(sourceLayout);
            } else {
                // Different presentation — clone the master slide chain
                destLayoutPartName = cloneMasterChainForSlide(sourceSlide, sourceLayout);
            }
        }

        // Fall back to first available layout if nothing resolved
        if (destLayoutPartName == null) {
            destLayoutPartName = getFirstLayoutPartName().orElse(null);
        }

        // Clone the slide part
        if (destLayoutPartName != null) {
            SlidePart.cloneFrom(sourcePackage, sourcePartName, pkg, destPartName, destLayoutPartName);
        } else {
            // No layout available — copy raw bytes
            byte[] content = sourcePackage.getPartBytes(sourcePartName);
            if (content != null) {
                pkg.setPartBytes(destPartName, content);
            }
        }

        // Add relationship from presentation to the new slide
        var presRels = new RelsHelper(pkg, PresentationPart.PART_NAME);
        String relativeTarget = "slides/slide" + nextNum + ".xml";
        String relId = presRels.addRelationship(REL_TYPE_SLIDE, relativeTarget);
        presRels.save();

        // Add slide reference to presentation.xml. Parsed only now: cloning notes can add a
        // notes master, which writes presentation.xml, and an earlier parse would undo that.
        PresentationPart presPart = getEffectivePresentationPart();
        presPart.addSlideReference(relId, null, index);
        presPart.save();

        // Create Slide object
        Slide slide = new Slide(presentation, nextNum - 1);

        // Add to internal list
        if (index < 0 || index >= slides.size()) {
            slides.add(slide);
        } else {
            slides.add(index, slide);
        }

        return slide;
    }

    /**
     * Finds a matching layout in the destination master for the source slide.
     *
     * <p>Matching is attempted first by layout type, then by name. If no match
     * is found, the first layout of the master is returned as a fallback.</p>
     *
     * @param sourceSlide the source slide whose layout to match
     * @param destMaster  the destination master slide
     * @param allowClone  if {@code true}, permits fallback to the first layout
     * @return the matching layout part name, or {@code null} if none available
     */
    public Optional<String> findMatchingLayout(ISlide sourceSlide, IMasterSlide destMaster,
                                               boolean allowClone) {
        ILayoutSlide sourceLayout = sourceSlide.getLayoutSlide();
        if (sourceLayout == null) {
            return firstLayoutPartNameOf(destMaster);
        }

        SlideLayoutType sourceType = sourceLayout.getLayoutType();
        String sourceName = sourceLayout.getName();

        // Match by layout type
        ILayoutSlideCollection layouts = destMaster.getLayoutSlides();
        if (layouts != null) {
            for (int i = 0; i < layouts.size(); i++) {
                ILayoutSlide layout = layouts.get(i);
                if (layout.getLayoutType() == sourceType) {
                    return Optional.ofNullable(resolveLayoutPartName(layout));
                }
            }

            // Match by name
            for (int i = 0; i < layouts.size(); i++) {
                ILayoutSlide layout = layouts.get(i);
                if (layout.getName() != null && layout.getName().equals(sourceName)) {
                    return Optional.ofNullable(resolveLayoutPartName(layout));
                }
            }
        }

        // Fallback to first layout
        return firstLayoutPartNameOf(destMaster);
    }

    /**
     * Finds a layout in the destination presentation matching the source layout type.
     *
     * @param sourceLayout the source layout to match
     * @return the matching layout part name, or empty if none found
     */
    public Optional<String> findLayoutByType(ILayoutSlide sourceLayout) {
        SlideLayoutType sourceType;
        try {
            sourceType = sourceLayout.getLayoutType();
        } catch (UnsupportedOperationException e) {
            return findLayoutFromLayoutSlides(sourceLayout);
        }

        // Search presentation's layout slides
        try {
            IGlobalLayoutSlideCollection globalLayouts = presentation.getLayoutSlides();
            if (globalLayouts != null) {
                for (int i = 0; i < globalLayouts.size(); i++) {
                    ILayoutSlide layout = globalLayouts.get(i);
                    try {
                        if (layout.getLayoutType() == sourceType) {
                            return Optional.ofNullable(resolveLayoutPartName(layout));
                        }
                    } catch (UnsupportedOperationException ignored) {
                        // skip
                    }
                }
            }
        } catch (UnsupportedOperationException ignored) {
            // skip
        }

        return getFirstLayoutPartName();
    }

    /**
     * Finds a matching layout by name using the presentation's layout slides.
     *
     * @param sourceLayout the source layout whose name to match
     * @return the matching layout part name, or empty if none found
     */
    public Optional<String> findLayoutFromLayoutSlides(ILayoutSlide sourceLayout) {
        try {
            String sourceName = sourceLayout.getName();
            IGlobalLayoutSlideCollection globalLayouts = presentation.getLayoutSlides();
            if (globalLayouts != null) {
                for (int i = 0; i < globalLayouts.size(); i++) {
                    ILayoutSlide layout = globalLayouts.get(i);
                    try {
                        if (layout.getName() != null && layout.getName().equals(sourceName)) {
                            return Optional.ofNullable(resolveLayoutPartName(layout));
                        }
                    } catch (UnsupportedOperationException ignored) {
                        // skip
                    }
                }
            }
        } catch (UnsupportedOperationException ignored) {
            // skip
        }
        return getFirstLayoutPartName();
    }

    /**
     * Returns the first available layout part name in the destination presentation.
     *
     * <p>Checks the presentation's layout slide collection first, then falls back
     * to scanning the package for layout files.</p>
     *
     * @return the first layout part name, or empty if none found
     */
    public Optional<String> getFirstLayoutPartName() {
        // Try presentation's layout slides
        try {
            IGlobalLayoutSlideCollection globalLayouts = presentation.getLayoutSlides();
            if (globalLayouts != null && !globalLayouts.isEmpty()) {
                ILayoutSlide first = globalLayouts.get(0);
                String partName = resolveLayoutPartName(first);
                if (partName != null) {
                    return Optional.of(partName);
                }
            }
        } catch (UnsupportedOperationException ignored) {
            // skip
        }

        // Scan the package for layout files
        OpcPackage pkg = getEffectivePackage();
        for (String partName : pkg.getPartNames()) {
            if (partName.startsWith("ppt/slideLayouts/") && partName.endsWith(".xml")) {
                return Optional.of(partName);
            }
        }

        return Optional.empty();
    }

    /**
     * Clones the master slide chain for a slide being cloned from another presentation.
     *
     * <p>This clones the source layout's master slide (including its layouts)
     * into the destination presentation, then returns the cloned layout's part
     * name that matches the source layout.</p>
     *
     * @param sourceSlide  the source slide being cloned
     * @param sourceLayout the source slide's layout
     * @return the cloned layout's part name, or {@code null} if cloning fails
     */
    public String cloneMasterChainForSlide(ISlide sourceSlide, ILayoutSlide sourceLayout) {
        if (sourceLayout == null) {
            return getFirstLayoutPartName().orElse(null);
        }

        // Get the source master
        IMasterSlide sourceMaster;
        try {
            sourceMaster = sourceLayout.getMasterSlide();
        } catch (UnsupportedOperationException e) {
            return getFirstLayoutPartName().orElse(null);
        }

        if (sourceMaster == null) {
            return getFirstLayoutPartName().orElse(null);
        }

        // Clone the master into the destination presentation
        IMasterSlide clonedMaster;
        try {
            clonedMaster = presentation.getMasters().addClone(sourceMaster);
        } catch (Exception e) {
            return getFirstLayoutPartName().orElse(null);
        }

        // Find the cloned layout matching the source layout
        SlideLayoutType sourceType = null;
        String sourceName = null;
        try {
            sourceType = sourceLayout.getLayoutType();
        } catch (UnsupportedOperationException ignored) {
            // skip
        }
        try {
            sourceName = sourceLayout.getName();
        } catch (UnsupportedOperationException ignored) {
            // skip
        }

        ILayoutSlideCollection clonedLayouts = clonedMaster.getLayoutSlides();
        if (clonedLayouts != null) {
            for (int i = 0; i < clonedLayouts.size(); i++) {
                ILayoutSlide clonedLayout = clonedLayouts.get(i);
                // Match by type
                if (sourceType != null) {
                    try {
                        if (clonedLayout.getLayoutType() == sourceType) {
                            return resolveLayoutPartName(clonedLayout);
                        }
                    } catch (UnsupportedOperationException ignored) {
                        // skip
                    }
                }
                // Match by name
                if (sourceName != null) {
                    try {
                        if (sourceName.equals(clonedLayout.getName())) {
                            return resolveLayoutPartName(clonedLayout);
                        }
                    } catch (UnsupportedOperationException ignored) {
                        // skip
                    }
                }
            }

            // Fallback: first layout from cloned master
            if (!clonedLayouts.isEmpty()) {
                return resolveLayoutPartName(clonedLayouts.get(0));
            }
        }

        return getFirstLayoutPartName().orElse(null);
    }

    // ── Private helpers ────────────────────────────────────────────────

    private OpcPackage getEffectivePackage() {
        if (opcPackage != null) {
            return opcPackage;
        }
        if (presentation != null) {
            return presentation.getPackage();
        }
        throw new IllegalStateException("SlideCollection not initialized");
    }

    /**
     * Returns a presentation part parsed from the package as it stands now.
     *
     * <p>Deliberately not cached: {@code presentation.xml} is also written by the document
     * properties and slide-numbering code, and a manager holding a DOM parsed before one of
     * those writes would silently undo it on its next save.</p>
     */
    private PresentationPart getEffectivePresentationPart() {
        return new PresentationPart(getEffectivePackage());
    }

    private static String resolveLayoutPartName(ILayoutSlide layout) {
        if (layout instanceof LayoutSlide ls) {
            String name = ls.getPartName();
            if (name != null) {
                return name;
            }
        }
        return null;
    }

    private static Optional<String> firstLayoutPartNameOf(IMasterSlide master) {
        ILayoutSlideCollection layouts = master.getLayoutSlides();
        if (layouts != null && !layouts.isEmpty()) {
            return Optional.ofNullable(resolveLayoutPartName(layouts.get(0)));
        }
        return Optional.empty();
    }

}
