package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Objects;
import java.util.Set;
import org.aspose.slides.foss.internal.xml.SecureXml;

/**
 * Manages a notes slide XML part ({@code ppt/notesSlides/notesSlideN.xml}).
 *
 * <p>Provides access to notes text and placeholder management within a
 * notes slide. The XML is loaded once on construction and mutated in memory;
 * call {@link #save()} to persist changes back to the package.</p>
 */
public final class NotesSlidePart {

    private static final String NS_P = PptxConstants.NAMESPACES.get("p");
    private static final String NS_A = PptxConstants.NAMESPACES.get("a");
    private static final String NS_R = PptxConstants.NAMESPACES.get("r");

    /** Text-bearing placeholder types that receive an empty {@code <p:txBody>} on creation. */
    private static final Set<String> TEXT_PH_TYPES = Set.of("dt", "ftr", "hdr");

    private static final String REL_NS =
            "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String REL_TYPE_SLIDE =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    private static final String REL_TYPE_NOTES_MASTER =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
    private static final String CONTENT_TYPE_NOTES_SLIDE =
            "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
    private static final String CT_NS =
            "http://schemas.openxmlformats.org/package/2006/content-types";

    private final OpcPackage pkg;
    private final String partName;
    private Document doc;
    private Element root;

    /**
     * Creates a {@code NotesSlidePart} for an existing part in the package.
     *
     * @param pkg      the OPC package containing the notes slide
     * @param partName the part path (e.g. {@code "ppt/notesSlides/notesSlide1.xml"})
     * @throws IllegalArgumentException if the part does not exist in the package
     */
    public NotesSlidePart(OpcPackage pkg, String partName) {
        this.pkg = Objects.requireNonNull(pkg);
        this.partName = Objects.requireNonNull(partName);
        load();
    }

    /**
     * Loads and parses the notes slide XML from the package.
     *
     * <p>If the part does not yet exist, {@code doc} and {@code root} remain {@code null}.
     * Use {@link #ensurePartExists(int)} to create a minimal part first.</p>
     */
    public void load() {
        doc = pkg.parseXml(partName);
        if (doc != null) {
            root = doc.getDocumentElement();
        }
    }

    /**
     * Returns the part name of this notes slide.
     *
     * @return the part path within the package
     */
    public String getPartName() {
        return partName;
    }

    /**
     * Returns the slide name from {@code <p:cSld name="...">}.
     *
     * @return the name attribute value, or empty string if not set
     */
    public String getName() {
        NodeList list = root.getElementsByTagNameNS(NS_P, "cSld");
        if (list.getLength() > 0) {
            return ((Element) list.item(0)).getAttribute("name");
        }
        return "";
    }

    /**
     * Sets the slide name on {@code <p:cSld name="...">}.
     *
     * @param value the name to set
     */
    public void setName(String value) {
        NodeList list = root.getElementsByTagNameNS(NS_P, "cSld");
        if (list.getLength() > 0) {
            ((Element) list.item(0)).setAttribute("name", value);
        }
    }

    /**
     * Returns the {@code <p:spTree>} element from the notes slide.
     *
     * @return the shape tree element, or {@code null} if not found
     */
    public Element getSpTree() {
        NodeList list = root.getElementsByTagNameNS(NS_P, "spTree");
        return list.getLength() > 0 ? (Element) list.item(0) : null;
    }

    /**
     * Finds the first placeholder shape with the given type.
     *
     * @param phType the placeholder type string (e.g. {@code "body"}, {@code "ftr"})
     * @return the {@code <p:sp>} element, or {@code null} if not found
     */
    public Element findPlaceholder(String phType) {
        Element spTree = getSpTree();
        if (spTree == null) {
            return null;
        }

        NodeList shapes = spTree.getElementsByTagNameNS(NS_P, "sp");
        for (int i = 0; i < shapes.getLength(); i++) {
            Element sp = (Element) shapes.item(i);
            NodeList phNodes = sp.getElementsByTagNameNS(NS_P, "ph");
            for (int j = 0; j < phNodes.getLength(); j++) {
                Element ph = (Element) phNodes.item(j);
                if (phType.equals(ph.getAttribute("type"))) {
                    return sp;
                }
            }
        }
        return null;
    }

    /**
     * Returns the {@code <p:txBody>} element of the notes body placeholder.
     *
     * @return the text body element, or {@code null} if not found
     */
    public Element getNotesTxbody() {
        Element bodySp = findPlaceholder("body");
        if (bodySp == null) {
            return null;
        }
        NodeList list = bodySp.getElementsByTagNameNS(NS_P, "txBody");
        return list.getLength() > 0 ? (Element) list.item(0) : null;
    }

    /**
     * Checks whether a placeholder of the given type exists.
     *
     * @param phType the placeholder type string
     * @return {@code true} if the placeholder exists
     */
    public boolean hasPlaceholder(String phType) {
        return findPlaceholder(phType) != null;
    }

    /**
     * Removes the placeholder shape of the given type.
     *
     * @param phType the placeholder type string to remove
     */
    public void removePlaceholder(String phType) {
        Element sp = findPlaceholder(phType);
        if (sp != null) {
            Element spTree = getSpTree();
            if (spTree != null) {
                spTree.removeChild(sp);
                save();
            }
        }
    }

    /**
     * Adds a minimal placeholder shape of the given type if not already present.
     *
     * @param phType the placeholder type string to add
     */
    public void addPlaceholder(String phType) {
        if (hasPlaceholder(phType)) {
            return;
        }

        Element spTree = getSpTree();
        if (spTree == null) {
            return;
        }

        int maxId = 1;
        NodeList shapes = spTree.getElementsByTagNameNS(NS_P, "sp");
        for (int i = 0; i < shapes.getLength(); i++) {
            Element sp = (Element) shapes.item(i);
            NodeList cNvPrs = sp.getElementsByTagNameNS(NS_P, "cNvPr");
            for (int j = 0; j < cNvPrs.getLength(); j++) {
                String idStr = ((Element) cNvPrs.item(j)).getAttribute("id");
                if (!idStr.isEmpty()) {
                    try {
                        maxId = Math.max(maxId, Integer.parseInt(idStr));
                    } catch (NumberFormatException ignored) {
                        // Non-numeric value; use default
                    }
                }
            }
        }

        int shapeId = maxId + 1;
        Element spElem = buildPlaceholderShape(phType, shapeId);
        spTree.appendChild(spElem);
        save();
    }

    /**
     * Builds a minimal placeholder shape element for notes slides.
     *
     * @param phType  the placeholder type string
     * @param shapeId the shape ID to assign
     * @return a new {@code <p:sp>} element
     */
    public Element buildPlaceholderShape(String phType, int shapeId) {
        Element sp = doc.createElementNS(NS_P, "p:sp");

        Element nvSpPr = doc.createElementNS(NS_P, "p:nvSpPr");
        Element cNvPr = doc.createElementNS(NS_P, "p:cNvPr");
        cNvPr.setAttribute("id", String.valueOf(shapeId));
        cNvPr.setAttribute("name", phType + " Placeholder " + shapeId);
        nvSpPr.appendChild(cNvPr);

        Element cNvSpPr = doc.createElementNS(NS_P, "p:cNvSpPr");
        Element spLocks = doc.createElementNS(NS_A, "a:spLocks");
        spLocks.setAttribute("noGrp", "1");
        cNvSpPr.appendChild(spLocks);
        nvSpPr.appendChild(cNvSpPr);

        Element nvPr = doc.createElementNS(NS_P, "p:nvPr");
        Element ph = doc.createElementNS(NS_P, "p:ph");
        ph.setAttribute("type", phType);
        nvPr.appendChild(ph);
        nvSpPr.appendChild(nvPr);

        sp.appendChild(nvSpPr);

        sp.appendChild(doc.createElementNS(NS_P, "p:spPr"));

        if (TEXT_PH_TYPES.contains(phType)) {
            Element txBody = doc.createElementNS(NS_P, "p:txBody");
            txBody.appendChild(doc.createElementNS(NS_A, "a:bodyPr"));
            txBody.appendChild(doc.createElementNS(NS_A, "a:lstStyle"));
            Element aP = doc.createElementNS(NS_A, "a:p");
            aP.appendChild(doc.createElementNS(NS_A, "a:endParaRPr"));
            txBody.appendChild(aP);
            sp.appendChild(txBody);
        }

        return sp;
    }

    /**
     * Sets text content for a placeholder shape, adding it if it does not exist.
     *
     * @param phType the placeholder type string
     * @param text   the text to set
     */
    public void setPlaceholderText(String phType, String text) {
        if (!hasPlaceholder(phType)) {
            addPlaceholder(phType);
        }

        Element sp = findPlaceholder(phType);
        if (sp == null) {
            return;
        }

        Element txBody = getFirstChildElement(sp, NS_P, "txBody");
        if (txBody == null) {
            txBody = doc.createElementNS(NS_P, "p:txBody");
            txBody.appendChild(doc.createElementNS(NS_A, "a:bodyPr"));
            txBody.appendChild(doc.createElementNS(NS_A, "a:lstStyle"));
            sp.appendChild(txBody);
        }

        // Remove existing paragraphs
        NodeList existingParas = txBody.getElementsByTagNameNS(NS_A, "p");
        while (existingParas.getLength() > 0) {
            txBody.removeChild(existingParas.item(0));
        }

        Element aP = doc.createElementNS(NS_A, "a:p");
        if (text != null && !text.isEmpty()) {
            Element aR = doc.createElementNS(NS_A, "a:r");
            Element aT = doc.createElementNS(NS_A, "a:t");
            aT.setTextContent(text);
            aR.appendChild(aT);
            aP.appendChild(aR);
        } else {
            aP.appendChild(doc.createElementNS(NS_A, "a:endParaRPr"));
        }
        txBody.appendChild(aP);
        save();
    }

    /**
     * Saves the notes slide XML back to the package.
     */
    public void save() {
        pkg.serializeXml(partName, doc);
    }

    /**
     * Resolves a relative target path to an absolute part name.
     *
     * @param target the relative or absolute target path
     * @return the resolved absolute part name
     */
    public String resolveTarget(String target) {
        if (target.startsWith("/")) {
            return target.substring(1);
        }
        String baseDir = partName.contains("/")
                ? partName.substring(0, partName.lastIndexOf('/'))
                : "";
        String[] segments = (baseDir + "/" + target).split("/");
        var resolved = new ArrayList<String>();
        for (String segment : segments) {
            if ("..".equals(segment)) {
                if (!resolved.isEmpty()) {
                    resolved.removeLast();
                }
            } else if (!segment.isEmpty() && !".".equals(segment)) {
                resolved.add(segment);
            }
        }
        return String.join("/", resolved);
    }

    // ── Static factory and lifecycle methods ─────────────────────────────

    /**
     * Creates a new empty notes slide in the package for a given slide.
     *
     * @param pkg           the OPC package
     * @param slidePartName the part name of the owning slide
     * @return the newly created {@code NotesSlidePart}
     */
    public static NotesSlidePart createEmpty(OpcPackage pkg, String slidePartName) {
        int nextNum = 1;
        String partName;
        while (true) {
            String candidate = "ppt/notesSlides/notesSlide" + nextNum + ".xml";
            if (!pkg.hasPart(candidate)) {
                partName = candidate;
                break;
            }
            nextNum++;
        }

        byte[] notesXml = buildNotesXml();
        pkg.setPartBytes(partName, notesXml);

        // Create the notes slide's own relationships
        String relsUri = getRelsPartName(partName);
        Document relsDoc = OpcPackage.newDocument();
        relsDoc.setXmlStandalone(true);
        Element relsRoot = relsDoc.createElementNS(REL_NS, "Relationships");
        relsDoc.appendChild(relsRoot);

        // Relationship: notes slide → parent slide
        String slideRelative = computeRelativeTarget(partName, slidePartName);
        Element rel1 = relsDoc.createElementNS(REL_NS, "Relationship");
        rel1.setAttribute("Id", "rId1");
        rel1.setAttribute("Type", REL_TYPE_SLIDE);
        rel1.setAttribute("Target", slideRelative);
        relsRoot.appendChild(rel1);

        // Relationship: notes slide → notes master (if present)
        String notesMasterPartName = findNotesMaster(pkg);
        if (notesMasterPartName != null) {
            String masterRelative = computeRelativeTarget(partName, notesMasterPartName);
            Element rel2 = relsDoc.createElementNS(REL_NS, "Relationship");
            rel2.setAttribute("Id", "rId2");
            rel2.setAttribute("Type", REL_TYPE_NOTES_MASTER);
            rel2.setAttribute("Target", masterRelative);
            relsRoot.appendChild(rel2);
        }

        pkg.serializeXml(relsUri, relsDoc);

        // Register content type
        addContentTypeOverride(pkg, partName, CONTENT_TYPE_NOTES_SLIDE);

        return new NotesSlidePart(pkg, partName);
    }

    /**
     * Finds the notes master part name in the package.
     *
     * @param pkg the OPC package
     * @return the part name, or {@code null} if not found
     */
    public static String findNotesMaster(OpcPackage pkg) {
        for (String part : pkg.getPartNames()) {
            if (part.startsWith("ppt/notesMasters/") && part.endsWith(".xml")) {
                return part;
            }
        }
        return null;
    }

    /**
     * Computes a relative path from one part to another.
     *
     * @param fromPart the source part path
     * @param toPart   the destination part path
     * @return the relative path
     */
    public static String computeRelativeTarget(String fromPart, String toPart) {
        String fromDir = fromPart.contains("/")
                ? fromPart.substring(0, fromPart.lastIndexOf('/'))
                : "";
        String toDir = toPart.contains("/")
                ? toPart.substring(0, toPart.lastIndexOf('/'))
                : "";
        String toFile = toPart.substring(toPart.lastIndexOf('/') + 1);

        if (fromDir.equals(toDir)) {
            return toFile;
        }

        String[] fromParts = fromDir.isEmpty() ? new String[0] : fromDir.split("/");
        String[] toParts = toDir.isEmpty() ? new String[0] : toDir.split("/");

        int commonLen = 0;
        for (int i = 0; i < Math.min(fromParts.length, toParts.length); i++) {
            if (fromParts[i].equals(toParts[i])) {
                commonLen = i + 1;
            } else {
                break;
            }
        }

        int upCount = fromParts.length - commonLen;
        var downSegments = new ArrayList<String>();
        for (int i = commonLen; i < toParts.length; i++) {
            downSegments.add(toParts[i]);
        }
        String downPath = String.join("/", downSegments);

        var sb = new StringBuilder();
        sb.append("../".repeat(upCount));
        if (!downPath.isEmpty()) {
            sb.append(downPath).append('/');
        }
        sb.append(toFile);
        return sb.toString();
    }

    /**
     * Builds a minimal notes slide XML with slide image and body placeholders.
     *
     * @return UTF-8 encoded XML bytes
     */
    public static byte[] buildNotesXml() {
        Document xmlDoc = OpcPackage.newDocument();
        xmlDoc.setXmlStandalone(true);

        Element notes = xmlDoc.createElementNS(NS_P, "p:notes");
        notes.setAttribute("xmlns:a", NS_A);
        notes.setAttribute("xmlns:r", NS_R);
        xmlDoc.appendChild(notes);

        // <p:cSld>
        Element cSld = xmlDoc.createElementNS(NS_P, "p:cSld");
        Element spTree = xmlDoc.createElementNS(NS_P, "p:spTree");

        // Group shape header (required)
        Element nvGrpSpPr = xmlDoc.createElementNS(NS_P, "p:nvGrpSpPr");
        Element cNvPrGrp = xmlDoc.createElementNS(NS_P, "p:cNvPr");
        cNvPrGrp.setAttribute("id", "1");
        cNvPrGrp.setAttribute("name", "");
        nvGrpSpPr.appendChild(cNvPrGrp);
        nvGrpSpPr.appendChild(xmlDoc.createElementNS(NS_P, "p:cNvGrpSpPr"));
        nvGrpSpPr.appendChild(xmlDoc.createElementNS(NS_P, "p:nvPr"));
        spTree.appendChild(nvGrpSpPr);

        Element grpSpPr = xmlDoc.createElementNS(NS_P, "p:grpSpPr");
        Element xfrm = xmlDoc.createElementNS(NS_A, "a:xfrm");
        Element off = xmlDoc.createElementNS(NS_A, "a:off");
        off.setAttribute("x", "0");
        off.setAttribute("y", "0");
        xfrm.appendChild(off);
        Element ext = xmlDoc.createElementNS(NS_A, "a:ext");
        ext.setAttribute("cx", "0");
        ext.setAttribute("cy", "0");
        xfrm.appendChild(ext);
        Element chOff = xmlDoc.createElementNS(NS_A, "a:chOff");
        chOff.setAttribute("x", "0");
        chOff.setAttribute("y", "0");
        xfrm.appendChild(chOff);
        Element chExt = xmlDoc.createElementNS(NS_A, "a:chExt");
        chExt.setAttribute("cx", "0");
        chExt.setAttribute("cy", "0");
        xfrm.appendChild(chExt);
        grpSpPr.appendChild(xfrm);
        spTree.appendChild(grpSpPr);

        // Slide image placeholder (type="sldImg")
        Element sp1 = xmlDoc.createElementNS(NS_P, "p:sp");
        Element nvSpPr1 = xmlDoc.createElementNS(NS_P, "p:nvSpPr");
        Element cNvPr1 = xmlDoc.createElementNS(NS_P, "p:cNvPr");
        cNvPr1.setAttribute("id", "2");
        cNvPr1.setAttribute("name", "Slide Image Placeholder 1");
        nvSpPr1.appendChild(cNvPr1);
        Element cNvSpPr1 = xmlDoc.createElementNS(NS_P, "p:cNvSpPr");
        Element spLocks1 = xmlDoc.createElementNS(NS_A, "a:spLocks");
        spLocks1.setAttribute("noGrp", "1");
        cNvSpPr1.appendChild(spLocks1);
        nvSpPr1.appendChild(cNvSpPr1);
        Element nvPr1 = xmlDoc.createElementNS(NS_P, "p:nvPr");
        Element ph1 = xmlDoc.createElementNS(NS_P, "p:ph");
        ph1.setAttribute("type", "sldImg");
        nvPr1.appendChild(ph1);
        nvSpPr1.appendChild(nvPr1);
        sp1.appendChild(nvSpPr1);
        sp1.appendChild(xmlDoc.createElementNS(NS_P, "p:spPr"));
        spTree.appendChild(sp1);

        // Notes body placeholder (type="body", idx="1")
        Element sp2 = xmlDoc.createElementNS(NS_P, "p:sp");
        Element nvSpPr2 = xmlDoc.createElementNS(NS_P, "p:nvSpPr");
        Element cNvPr2 = xmlDoc.createElementNS(NS_P, "p:cNvPr");
        cNvPr2.setAttribute("id", "3");
        cNvPr2.setAttribute("name", "Notes Placeholder 2");
        nvSpPr2.appendChild(cNvPr2);
        Element cNvSpPr2 = xmlDoc.createElementNS(NS_P, "p:cNvSpPr");
        Element spLocks2 = xmlDoc.createElementNS(NS_A, "a:spLocks");
        spLocks2.setAttribute("noGrp", "1");
        cNvSpPr2.appendChild(spLocks2);
        nvSpPr2.appendChild(cNvSpPr2);
        Element nvPr2 = xmlDoc.createElementNS(NS_P, "p:nvPr");
        Element ph2 = xmlDoc.createElementNS(NS_P, "p:ph");
        ph2.setAttribute("type", "body");
        ph2.setAttribute("idx", "1");
        nvPr2.appendChild(ph2);
        nvSpPr2.appendChild(nvPr2);
        sp2.appendChild(nvSpPr2);
        sp2.appendChild(xmlDoc.createElementNS(NS_P, "p:spPr"));
        Element txBody = xmlDoc.createElementNS(NS_P, "p:txBody");
        txBody.appendChild(xmlDoc.createElementNS(NS_A, "a:bodyPr"));
        txBody.appendChild(xmlDoc.createElementNS(NS_A, "a:lstStyle"));
        Element aP = xmlDoc.createElementNS(NS_A, "a:p");
        aP.appendChild(xmlDoc.createElementNS(NS_A, "a:endParaRPr"));
        txBody.appendChild(aP);
        sp2.appendChild(txBody);
        spTree.appendChild(sp2);

        cSld.appendChild(spTree);
        notes.appendChild(cSld);

        // Color map override
        Element clrMapOvr = xmlDoc.createElementNS(NS_P, "p:clrMapOvr");
        clrMapOvr.appendChild(xmlDoc.createElementNS(NS_A, "a:masterClrMapping"));
        notes.appendChild(clrMapOvr);

        try {
            var transformer = SecureXml.transformerFactory().newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
            transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            var baos = new ByteArrayOutputStream();
            transformer.transform(new DOMSource(xmlDoc), new StreamResult(baos));
            return baos.toByteArray();
        } catch (Exception e) {
            throw new IllegalStateException("Failed to serialize notes XML", e);
        }
    }

    /**
     * Deletes a notes slide and its associated files from the package.
     *
     * @param pkg      the OPC package
     * @param partName the notes slide part name to delete
     */
    public static void delete(OpcPackage pkg, String partName) {
        pkg.removePart(partName);

        String relsPartName = getRelsPartName(partName);
        pkg.removePart(relsPartName);

        removeContentTypeOverride(pkg, partName);
    }

    // ── Convenience methods (used by higher-level API) ─────────────────

    /**
     * Ensures the notes slide XML part exists in the package with a minimal structure.
     *
     * @param slideNumber the 1-based slide number this notes slide is associated with
     */
    public void ensurePartExists(int slideNumber) {
        if (pkg.hasPart(partName)) {
            return;
        }
        byte[] xml = buildNotesXml();
        pkg.setPartBytes(partName, xml);
        load();
    }

    /**
     * Returns the text from the body (notes) placeholder.
     *
     * @return the notes text, or empty string if not found
     */
    public String getNotesText() {
        Element spTree = getSpTree();
        if (spTree == null) {
            return "";
        }
        Element shape = findPlaceholder("body");
        if (shape == null) {
            return "";
        }
        Element txBody = getFirstChildElement(shape, NS_P, "txBody");
        if (txBody == null) {
            return "";
        }
        var sb = new StringBuilder();
        NodeList runs = txBody.getElementsByTagNameNS(NS_A, "t");
        for (int i = 0; i < runs.getLength(); i++) {
            sb.append(runs.item(i).getTextContent());
        }
        return sb.toString();
    }

    /**
     * Sets the text of the body (notes) placeholder.
     *
     * @param text the notes text
     */
    public void setNotesText(String text) {
        setPlaceholderText("body", text);
    }

    // ── Private helpers ──────────────────────────────────────────────────

    private static Element getFirstChildElement(Element parent, String ns, String localName) {
        NodeList list = parent.getElementsByTagNameNS(ns, localName);
        return list.getLength() > 0 ? (Element) list.item(0) : null;
    }

    private static String getRelsPartName(String partName) {
        int lastSlash = partName.lastIndexOf('/');
        if (lastSlash >= 0) {
            return partName.substring(0, lastSlash + 1) + "_rels/"
                    + partName.substring(lastSlash + 1) + ".rels";
        }
        return "_rels/" + partName + ".rels";
    }

    private static void addContentTypeOverride(OpcPackage pkg, String partName,
                                                String contentType) {
        Document ctDoc = pkg.parseXml("[Content_Types].xml");
        if (ctDoc == null) {
            return;
        }
        Element ctRoot = ctDoc.getDocumentElement();
        String partNameWithSlash = "/" + partName;
        NodeList overrides = ctRoot.getElementsByTagName("Override");
        for (int i = 0; i < overrides.getLength(); i++) {
            Element ov = (Element) overrides.item(i);
            if (partNameWithSlash.equals(ov.getAttribute("PartName"))) {
                ov.setAttribute("ContentType", contentType);
                pkg.serializeXml("[Content_Types].xml", ctDoc);
                return;
            }
        }

        Element override = ctDoc.createElementNS(CT_NS, "Override");
        override.setAttribute("PartName", partNameWithSlash);
        override.setAttribute("ContentType", contentType);
        ctRoot.appendChild(override);
        pkg.serializeXml("[Content_Types].xml", ctDoc);
    }

    private static void removeContentTypeOverride(OpcPackage pkg, String partName) {
        Document ctDoc = pkg.parseXml("[Content_Types].xml");
        if (ctDoc == null) {
            return;
        }
        Element ctRoot = ctDoc.getDocumentElement();
        String partNameWithSlash = "/" + partName;
        NodeList overrides = ctRoot.getElementsByTagName("Override");
        for (int i = overrides.getLength() - 1; i >= 0; i--) {
            Element ov = (Element) overrides.item(i);
            if (partNameWithSlash.equals(ov.getAttribute("PartName"))) {
                ctRoot.removeChild(ov);
            }
        }
        pkg.serializeXml("[Content_Types].xml", ctDoc);
    }
}
