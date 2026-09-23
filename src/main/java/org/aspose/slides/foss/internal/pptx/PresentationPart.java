package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * Manages the {@code ppt/presentation.xml} part.
 *
 * <p>This class provides methods to:</p>
 * <ul>
 *   <li>Parse presentation structure (slides, masters, sizes)</li>
 *   <li>Add/remove slide references</li>
 *   <li>Get/set presentation properties</li>
 *   <li>Serialize back to XML</li>
 * </ul>
 */
public final class PresentationPart {

    /** The part URI inside the OPC package. */
    public static final String PART_NAME = "ppt/presentation.xml";

    private static final String NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static final String NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private final OpcPackage pkg;
    private Document doc;
    private Element root;
    private final List<SlideReference> slideRefs = new ArrayList<>();
    private final List<MasterReference> masterRefs = new ArrayList<>();

    /**
     * Creates a presentation part manager and loads the XML from the package.
     *
     * @param pkg the OPC package containing the presentation
     * @throws IllegalArgumentException if {@code ppt/presentation.xml} is not found
     */
    public PresentationPart(OpcPackage pkg) {
        this.pkg = pkg;
        load();
    }

    /**
     * Loads and parses {@code presentation.xml} from the package.
     *
     * @throws IllegalArgumentException if the part is not found
     */
    public void load() {
        doc = pkg.parseXml(PART_NAME);
        if (doc == null) {
            throw new IllegalArgumentException(PART_NAME + " not found in package");
        }
        root = doc.getDocumentElement();
        parseSlides();
        parseMasters();
    }

    /**
     * Parses slide references from the presentation XML.
     */
    public void parseSlides() {
        slideRefs.clear();
        NodeList lists = root.getElementsByTagNameNS(NS_P, "sldIdLst");
        if (lists.getLength() > 0) {
            Element sldIdLst = (Element) lists.item(0);
            NodeList items = sldIdLst.getElementsByTagNameNS(NS_P, "sldId");
            for (int i = 0; i < items.getLength(); i++) {
                Element el = (Element) items.item(i);
                int slideId = Integer.parseInt(el.getAttribute("id"));
                String relId = el.getAttributeNS(NS_R, "id");
                slideRefs.add(new SlideReference(slideId, relId));
            }
        }
    }

    /**
     * Parses master slide references from the presentation XML.
     */
    public void parseMasters() {
        masterRefs.clear();
        NodeList lists = root.getElementsByTagNameNS(NS_P, "sldMasterIdLst");
        if (lists.getLength() > 0) {
            Element sldMasterIdLst = (Element) lists.item(0);
            NodeList items = sldMasterIdLst.getElementsByTagNameNS(NS_P, "sldMasterId");
            for (int i = 0; i < items.getLength(); i++) {
                Element el = (Element) items.item(i);
                long masterId = Long.parseLong(el.getAttribute("id"));
                String relId = el.getAttributeNS(NS_R, "id");
                masterRefs.add(new MasterReference(masterId, relId));
            }
        }
    }

    /**
     * Returns a copy of the master slide reference list.
     *
     * @return list of master references
     */
    public List<MasterReference> masterReferences() {
        return List.copyOf(masterRefs);
    }

    /**
     * Returns the root element of the presentation XML document.
     *
     * @return the root DOM element, or {@code null} if not loaded
     */
    public Element elementTree() {
        return root;
    }

    /**
     * Returns a copy of the slide reference list in order.
     *
     * @return list of slide references
     */
    public List<SlideReference> slideReferences() {
        return List.copyOf(slideRefs);
    }

    /**
     * Returns the number of slides.
     *
     * @return the slide count
     */
    public int slideCount() {
        return slideRefs.size();
    }

    /**
     * Finds a slide reference by its unique slide ID.
     *
     * @param slideId the unique slide ID
     * @return the matching reference, or empty if not found
     */
    public Optional<SlideReference> getSlideRefById(int slideId) {
        for (SlideReference ref : slideRefs) {
            if (ref.getSlideId() == slideId) {
                return Optional.of(ref);
            }
        }
        return Optional.empty();
    }

    /**
     * Finds a slide reference by its relationship ID.
     *
     * @param relId the relationship ID (e.g. {@code "rId2"})
     * @return the matching reference, or empty if not found
     */
    public Optional<SlideReference> getSlideRefByRelId(String relId) {
        for (SlideReference ref : slideRefs) {
            if (ref.getRelId().equals(relId)) {
                return Optional.of(ref);
            }
        }
        return Optional.empty();
    }

    /**
     * Generates the next available slide ID.
     *
     * @return the next slide ID (starts at 256 per PPTX convention)
     */
    public int getNextSlideId() {
        if (slideRefs.isEmpty()) {
            return 256;
        }
        int max = 0;
        for (SlideReference ref : slideRefs) {
            if (ref.getSlideId() > max) {
                max = ref.getSlideId();
            }
        }
        return max + 1;
    }

    /**
     * Generates the next available master/layout slide ID.
     *
     * <p>In PPTX, master IDs ({@code sldMasterIdLst}) and layout IDs ({@code sldLayoutIdLst})
     * share the same ID space and must all be unique. This method scans both to find the
     * next available ID.</p>
     *
     * @return the next master/layout ID
     */
    public long getNextMasterId() {
        long maxId = 2147483647L; // one below PPTX convention start

        for (MasterReference ref : masterRefs) {
            if (ref.getMasterId() > maxId) {
                maxId = ref.getMasterId();
            }
        }

        // Check layout IDs across all master slides in the package
        for (String partName : pkg.getPartNames()) {
            if (partName.startsWith("ppt/slideMasters/") && partName.endsWith(".xml")) {
                Document masterDoc = pkg.parseXml(partName);
                if (masterDoc != null) {
                    NodeList layouts = masterDoc.getElementsByTagNameNS(NS_P, "sldLayoutId");
                    for (int i = 0; i < layouts.getLength(); i++) {
                        Element el = (Element) layouts.item(i);
                        String idStr = el.getAttribute("id");
                        if (!idStr.isEmpty()) {
                            long layoutId = Long.parseLong(idStr);
                            if (layoutId > maxId) {
                                maxId = layoutId;
                            }
                        }
                    }
                }
            }
        }

        return maxId + 1;
    }

    /**
     * Adds a new master slide reference to the presentation.
     *
     * @param relId    the relationship ID for the master slide
     * @param masterId the specific master ID, or {@code null} to auto-generate
     * @return the created master reference
     */
    public MasterReference addMasterReference(String relId, Long masterId) {
        if (masterId == null) {
            masterId = getNextMasterId();
        }

        MasterReference ref = new MasterReference(masterId, relId);
        masterRefs.add(ref);

        // Find or create sldMasterIdLst
        NodeList lists = root.getElementsByTagNameNS(NS_P, "sldMasterIdLst");
        Element sldMasterIdLst;
        if (lists.getLength() > 0) {
            sldMasterIdLst = (Element) lists.item(0);
        } else {
            sldMasterIdLst = doc.createElementNS(NS_P, "p:sldMasterIdLst");
            // Insert before sldIdLst if present
            NodeList sldIdLists = root.getElementsByTagNameNS(NS_P, "sldIdLst");
            if (sldIdLists.getLength() > 0) {
                root.insertBefore(sldMasterIdLst, sldIdLists.item(0));
            } else {
                root.appendChild(sldMasterIdLst);
            }
        }

        Element elem = doc.createElementNS(NS_P, "p:sldMasterId");
        elem.setAttribute("id", String.valueOf(masterId));
        elem.setAttributeNS(NS_R, "r:id", relId);
        sldMasterIdLst.appendChild(elem);

        return ref;
    }

    /**
     * Registers the notes master in {@code <p:notesMasterIdLst>}, replacing any entry there.
     *
     * <p>A presentation has at most one notes master. The list is created, if it is missing,
     * at the position {@code CT_Presentation} requires — after {@code sldMasterIdLst} and
     * before {@code sldIdLst} — since a list appended after the slide list is a schema
     * error, not a formatting preference.</p>
     *
     * @param relId the id of the presentation's relationship to the notes master
     */
    public void setNotesMasterReference(String relId) {
        NodeList lists = root.getElementsByTagNameNS(NS_P, "notesMasterIdLst");
        Element idLst;
        if (lists.getLength() > 0) {
            idLst = (Element) lists.item(0);
            while (idLst.getFirstChild() != null) {
                idLst.removeChild(idLst.getFirstChild());
            }
        } else {
            idLst = SchemaOrder.insert(root, doc.createElementNS(NS_P, "p:notesMasterIdLst"));
        }
        Element entry = doc.createElementNS(NS_P, "p:notesMasterId");
        entry.setAttributeNS(NS_R, "r:id", relId);
        idLst.appendChild(entry);
    }

    /**
     * Adds a new slide reference to the presentation.
     *
     * @param relId   the relationship ID for the slide
     * @param slideId the specific slide ID, or {@code null} to auto-generate
     * @param index   position to insert at; {@code -1} means append at end
     * @return the created slide reference
     */
    public SlideReference addSlideReference(String relId, Integer slideId, int index) {
        if (slideId == null) {
            slideId = getNextSlideId();
        }

        SlideReference ref = new SlideReference(slideId, relId);

        // Add to internal list
        if (index < 0 || index >= slideRefs.size()) {
            slideRefs.add(ref);
        } else {
            slideRefs.add(index, ref);
        }

        // Find or create sldIdLst
        NodeList lists = root.getElementsByTagNameNS(NS_P, "sldIdLst");
        Element sldIdLst;
        if (lists.getLength() > 0) {
            sldIdLst = (Element) lists.item(0);
        } else {
            sldIdLst = doc.createElementNS(NS_P, "p:sldIdLst");
            root.appendChild(sldIdLst);
        }

        Element elem = doc.createElementNS(NS_P, "p:sldId");
        elem.setAttribute("id", String.valueOf(slideId));
        elem.setAttributeNS(NS_R, "r:id", relId);

        // Insert at position in XML
        NodeList existing = sldIdLst.getElementsByTagNameNS(NS_P, "sldId");
        if (index < 0 || index >= existing.getLength()) {
            sldIdLst.appendChild(elem);
        } else {
            sldIdLst.insertBefore(elem, existing.item(index));
        }

        return ref;
    }

    /**
     * Removes a slide reference by its slide ID.
     *
     * @param slideId the unique slide ID to remove
     * @return {@code true} if removed, {@code false} if not found
     */
    public boolean removeSlideReference(int slideId) {
        // Remove from internal list
        boolean found = false;
        for (int i = 0; i < slideRefs.size(); i++) {
            if (slideRefs.get(i).getSlideId() == slideId) {
                slideRefs.remove(i);
                found = true;
                break;
            }
        }
        if (!found) {
            return false;
        }

        // Remove from XML
        NodeList lists = root.getElementsByTagNameNS(NS_P, "sldIdLst");
        if (lists.getLength() > 0) {
            Element sldIdLst = (Element) lists.item(0);
            NodeList items = sldIdLst.getElementsByTagNameNS(NS_P, "sldId");
            for (int i = 0; i < items.getLength(); i++) {
                Element el = (Element) items.item(i);
                if (el.getAttribute("id").equals(String.valueOf(slideId))) {
                    sldIdLst.removeChild(el);
                    break;
                }
            }
        }

        return true;
    }

    /**
     * Removes a slide reference by the relationship it points through.
     *
     * @param relId the relationship ID of the slide to unregister
     * @return {@code true} if removed, {@code false} if no reference used that relationship
     */
    public boolean removeSlideReferenceByRelId(String relId) {
        Optional<SlideReference> ref = getSlideRefByRelId(relId);
        return ref.isPresent() && removeSlideReference(ref.get().getSlideId());
    }

    /**
     * Returns the slide size in EMUs (English Metric Units).
     *
     * @return an array {@code [width, height]} in EMUs
     */
    public int[] getSlideSize() {
        NodeList nodes = root.getElementsByTagNameNS(NS_P, "sldSz");
        if (nodes.getLength() > 0) {
            Element el = (Element) nodes.item(0);
            int cx = parseIntAttr(el, "cx", 9_144_000);
            int cy = parseIntAttr(el, "cy", 6_858_000);
            return new int[]{cx, cy};
        }
        return new int[]{9_144_000, 6_858_000}; // default widescreen
    }

    /**
     * Returns the notes slide size in EMUs.
     *
     * @return an array {@code [width, height]} in EMUs
     */
    public int[] getNotesSize() {
        NodeList nodes = root.getElementsByTagNameNS(NS_P, "notesSz");
        if (nodes.getLength() > 0) {
            Element el = (Element) nodes.item(0);
            int cx = parseIntAttr(el, "cx", 6_858_000);
            int cy = parseIntAttr(el, "cy", 9_144_000);
            return new int[]{cx, cy};
        }
        return new int[]{6_858_000, 9_144_000}; // default portrait
    }

    /**
     * Sets the notes slide size in EMUs.
     *
     * @param cx width in EMUs
     * @param cy height in EMUs
     */
    public void setNotesSize(int cx, int cy) {
        NodeList nodes = root.getElementsByTagNameNS(NS_P, "notesSz");
        if (nodes.getLength() > 0) {
            Element el = (Element) nodes.item(0);
            el.setAttribute("cx", String.valueOf(cx));
            el.setAttribute("cy", String.valueOf(cy));
        } else {
            Element notesSz = doc.createElementNS(NS_P, "p:notesSz");
            notesSz.setAttribute("cx", String.valueOf(cx));
            notesSz.setAttribute("cy", String.valueOf(cy));
            // Insert after sldSz if present
            NodeList sldSzNodes = root.getElementsByTagNameNS(NS_P, "sldSz");
            if (sldSzNodes.getLength() > 0) {
                Element sldSz = (Element) sldSzNodes.item(0);
                root.insertBefore(notesSz, sldSz.getNextSibling());
            } else {
                root.appendChild(notesSz);
            }
        }
    }

    /**
     * Returns the first slide number for numbering.
     *
     * @return the first slide number (defaults to 1)
     */
    public int getFirstSlideNumber() {
        String val = root.getAttribute("firstSlideNum");
        if (val == null || val.isEmpty()) {
            return 1;
        }
        return Integer.parseInt(val);
    }

    /**
     * Sets the first slide number for numbering.
     *
     * @param number the first slide number
     */
    public void setFirstSlideNumber(int number) {
        root.setAttribute("firstSlideNum", String.valueOf(number));
    }

    /**
     * Saves the {@code presentation.xml} back to the package.
     */
    public void save() {
        pkg.serializeXml(PART_NAME, doc);
    }

    private static int parseIntAttr(Element el, String attr, int defaultValue) {
        String val = el.getAttribute(attr);
        if (val == null || val.isEmpty()) {
            return defaultValue;
        }
        return Integer.parseInt(val);
    }
}
