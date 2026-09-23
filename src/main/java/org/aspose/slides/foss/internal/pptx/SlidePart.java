package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.aspose.slides.foss.internal.xml.SecureXml;

/**
 * Static utility methods for cloning slide parts and updating relationships.
 *
 * <p>These methods are used by {@code MasterSlideCollection} and similar classes
 * when cloning master slides, layout slides, and their related resources between
 * OPC packages.</p>
 */
public final class SlidePart {

    private static final String NS_R =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String REL_NS =
            "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String REL_TYPE_SLIDE_LAYOUT =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private static final String REL_TYPE_NOTES_SLIDE =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
    private static final String REL_TYPE_SLIDE =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    private SlidePart() {
        // utility class
    }

    /**
     * Creates an empty slide part in the OPC package with a relationship to the given layout.
     *
     * @param pkg             the OPC package
     * @param partName        the part name for the new slide (e.g. {@code "ppt/slides/slide2.xml"})
     * @param layoutPartName  the part name of the layout slide to reference
     */
    public static void createEmpty(OpcPackage pkg, String partName, String layoutPartName) {
        String slideXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
                + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                + "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
                + "<p:cSld><p:spTree>"
                + "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
                + "<p:grpSpPr/></p:spTree></p:cSld>"
                + "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";
        pkg.setPartBytes(partName, slideXml.getBytes(StandardCharsets.UTF_8));

        // Create rels pointing to the layout
        var rels = new RelsHelper(pkg, partName);
        String relativeTarget = computeRelativeTarget(partName, layoutPartName);
        rels.addRelationship(REL_TYPE_SLIDE_LAYOUT, relativeTarget);
        rels.save();

        // Register content type
        var ctManager = new ContentTypesManager(pkg);
        ctManager.addOverride(partName, ContentTypesManager.CONTENT_TYPES.get("slide"));
        ctManager.save();
    }

    /**
     * Clones a slide part from a source package to a destination package,
     * re-pointing the layout relationship to the given destination layout.
     *
     * @param sourcePackage       the source OPC package
     * @param sourcePartName      the source slide part name
     * @param destPackage         the destination OPC package
     * @param destPartName        the destination slide part name
     * @param destLayoutPartName  the layout part name to reference in the destination
     */
    public static void cloneFrom(OpcPackage sourcePackage, String sourcePartName,
                                 OpcPackage destPackage, String destPartName,
                                 String destLayoutPartName) {
        byte[] sourceContent = sourcePackage.getPartBytes(sourcePartName);
        if (sourceContent == null) {
            throw new IllegalArgumentException("Slide not found: " + sourcePartName);
        }

        Document destDoc = parseXml(sourceContent);
        Element destRoot = destDoc.getDocumentElement();

        var sourceRels = new RelsHelper(sourcePackage, sourcePartName);
        var destRels = new RelsHelper(destPackage, destPartName);
        Map<String, String> ridMapping = new HashMap<>();

        for (RelsHelper.RelEntry rel : sourceRels.getAllRelationships()) {
            if (REL_TYPE_SLIDE_LAYOUT.equals(rel.type())) {
                // Re-point to destination layout
                String relativeTarget = computeRelativeTarget(destPartName, destLayoutPartName);
                String newRid = destRels.addRelationship(rel.type(), relativeTarget);
                ridMapping.put(rel.id(), newRid);
            } else if ("External".equals(rel.targetMode())) {
                String newRid = destRels.addRelationship(rel.type(), rel.target(), "External");
                ridMapping.put(rel.id(), newRid);
            } else if (REL_TYPE_NOTES_SLIDE.equals(rel.type())) {
                String sourceNotes = resolveTargetStatic(sourcePartName, rel.target());
                String destNotes = cloneNotesSlide(
                        sourcePackage, sourceNotes, destPackage, destPartName);
                if (destNotes != null) {
                    String newRid = destRels.addRelationship(rel.type(),
                            computeRelativeTarget(destPartName, destNotes));
                    ridMapping.put(rel.id(), newRid);
                }
            } else {
                String sourceTarget = resolveTargetStatic(sourcePartName, rel.target());
                String destTarget = cloneRelatedPart(
                        sourcePackage, sourceTarget, destPackage, destPartName, rel.type());
                String relativeTarget = computeRelativeTarget(destPartName, destTarget);
                String newRid = destRels.addRelationship(rel.type(), relativeTarget);
                ridMapping.put(rel.id(), newRid);
            }
        }

        updateRidReferences(destRoot, ridMapping);
        destPackage.setPartBytes(destPartName, serializeXml(destRoot));
        destRels.save();

        var ctManager = new ContentTypesManager(destPackage);
        ctManager.addOverride(destPartName, ContentTypesManager.CONTENT_TYPES.get("slide"));
        ctManager.save();
    }

    /**
     * Resolves a relative target path against a source part name to produce an absolute part path.
     *
     * @param sourcePartName the source part path (e.g. {@code "ppt/slideMasters/slideMaster1.xml"})
     * @param relativeTarget the relative target (e.g. {@code "../theme/theme1.xml"})
     * @return the resolved absolute part path (e.g. {@code "ppt/theme/theme1.xml"})
     */
    public static String resolveTargetStatic(String sourcePartName, String relativeTarget) {
        if (relativeTarget.startsWith("/")) {
            return relativeTarget.substring(1);
        }
        String baseDir = sourcePartName.contains("/")
                ? sourcePartName.substring(0, sourcePartName.lastIndexOf('/'))
                : "";
        String combined = baseDir.isEmpty() ? relativeTarget : baseDir + "/" + relativeTarget;
        String[] segments = combined.split("/");
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

    /**
     * Computes a relative path from one part to another within the OPC package.
     *
     * @param fromPart the source part path
     * @param toPart   the destination part path
     * @return the relative path from source to destination
     */
    public static String computeRelativeTarget(String fromPart, String toPart) {
        return NotesSlidePart.computeRelativeTarget(fromPart, toPart);
    }

    /**
     * Clones a related part (theme, image, etc.) from one package to another.
     *
     * <p>The destination part name is computed based on the relationship type and
     * existing parts in the destination package.</p>
     *
     * @param sourcePackage  the source OPC package
     * @param sourceTarget   the absolute path of the source part
     * @param destPackage    the destination OPC package
     * @param destPartName   the part name of the referring part in the destination
     * @param relType        the relationship type URI
     * @return the absolute path of the cloned part in the destination package
     */
    public static String cloneRelatedPart(OpcPackage sourcePackage, String sourceTarget,
                                          OpcPackage destPackage, String destPartName,
                                          String relType) {
        byte[] content = sourcePackage.getPartBytes(sourceTarget);
        if (content == null) {
            return sourceTarget;
        }

        // Determine destination path - use same path if not taken, otherwise find next number
        String destTarget = sourceTarget;
        if (destPackage.hasPart(destTarget)) {
            destTarget = findNextAvailableName(destPackage, sourceTarget);
        }

        destPackage.setPartBytes(destTarget, content);

        // Copy the .rels file for the cloned part if one exists
        String sourceRels = getRelsPartName(sourceTarget);
        byte[] relsContent = sourcePackage.getPartBytes(sourceRels);
        if (relsContent != null) {
            String destRels = getRelsPartName(destTarget);
            destPackage.setPartBytes(destRels, relsContent);
        }

        return destTarget;
    }

    /**
     * Copies a notes slide for a cloned slide.
     *
     * <p>The notes slide's relationships are copied with their ids, so that every
     * {@code r:id} in the notes still resolves, but two of them are re-pointed rather than
     * copied: the one back to the slide, which must name the clone — a clone's notes that
     * point at the source slide belong, to any reader, to the source — and the one to the
     * notes master, which must name the destination package's notes master. The copy gets
     * its own content-type override.</p>
     *
     * @param sourcePackage     the package holding the source notes slide
     * @param sourceNotes       the part name of the source notes slide
     * @param destPackage       the package receiving the clone
     * @param destSlidePartName the part name of the cloned slide
     * @return the part name of the copied notes slide, or {@code null} if the source part is
     *         missing
     */
    private static String cloneNotesSlide(OpcPackage sourcePackage, String sourceNotes,
                                          OpcPackage destPackage, String destSlidePartName) {
        byte[] content = sourcePackage.getPartBytes(sourceNotes);
        if (content == null) {
            return null;
        }
        String destNotes = NotesSlidePart.newPartName(destPackage, destSlidePartName);
        destPackage.setPartBytes(destNotes, content.clone());
        String destRelsName = getRelsPartName(destNotes);
        destPackage.removePart(destRelsName);

        boolean hasSlide = false;
        boolean hasMaster = false;
        byte[] sourceRels = sourcePackage.getPartBytes(getRelsPartName(sourceNotes));
        if (sourceRels != null) {
            Document relsDoc = parseXml(sourceRels);
            NodeList relElements = relsDoc.getDocumentElement()
                    .getElementsByTagNameNS(REL_NS, "Relationship");
            for (int i = 0; i < relElements.getLength(); i++) {
                Element rel = (Element) relElements.item(i);
                String type = rel.getAttribute("Type");
                if (REL_TYPE_SLIDE.equals(type)) {
                    rel.setAttribute("Target", computeRelativeTarget(destNotes, destSlidePartName));
                    hasSlide = true;
                } else if (NotesMasterPart.REL_TYPE.equals(type)) {
                    rel.setAttribute("Target", computeRelativeTarget(destNotes,
                            NotesMasterPart.ensureInPackage(destPackage)));
                    hasMaster = true;
                }
            }
            destPackage.serializeXml(destRelsName, relsDoc);
        }
        if (!hasSlide || !hasMaster) {
            var rels = new RelsHelper(destPackage, destNotes);
            if (!hasSlide) {
                rels.addRelationship(REL_TYPE_SLIDE,
                        computeRelativeTarget(destNotes, destSlidePartName));
            }
            if (!hasMaster) {
                rels.addRelationship(NotesMasterPart.REL_TYPE, computeRelativeTarget(destNotes,
                        NotesMasterPart.ensureInPackage(destPackage)));
            }
            rels.save();
        }

        var ctManager = new ContentTypesManager(destPackage);
        ctManager.addOverride(destNotes, ContentTypesManager.CONTENT_TYPES.get("notes_slide"));
        ctManager.save();
        return destNotes;
    }

    /**
     * Updates all {@code r:id} attribute references in an XML element tree according
     * to the given mapping.
     *
     * @param root       the root element to update
     * @param ridMapping a map from old relationship IDs to new relationship IDs
     */
    public static void updateRidReferences(Element root, Map<String, String> ridMapping) {
        if (ridMapping.isEmpty()) {
            return;
        }
        updateRidReferencesRecursive(root, ridMapping);
    }

    // ── Private helpers ──────────────────────────────────────────────────

    private static void updateRidReferencesRecursive(Element element, Map<String, String> ridMapping) {
        // Check r:id attribute
        String rId = element.getAttributeNS(NS_R, "id");
        if (!rId.isEmpty() && ridMapping.containsKey(rId)) {
            element.setAttributeNS(NS_R, "r:id", ridMapping.get(rId));
        }

        // Check r:embed attribute
        String rEmbed = element.getAttributeNS(NS_R, "embed");
        if (!rEmbed.isEmpty() && ridMapping.containsKey(rEmbed)) {
            element.setAttributeNS(NS_R, "r:embed", ridMapping.get(rEmbed));
        }

        // Check r:link attribute
        String rLink = element.getAttributeNS(NS_R, "r:link");
        if (!rLink.isEmpty() && ridMapping.containsKey(rLink)) {
            element.setAttributeNS(NS_R, "r:link", ridMapping.get(rLink));
        }

        // Also check non-namespaced attributes that might contain rIds
        NamedNodeMap attrs = element.getAttributes();
        for (int i = 0; i < attrs.getLength(); i++) {
            Node attr = attrs.item(i);
            if ("r:id".equals(attr.getNodeName()) || "r:embed".equals(attr.getNodeName())
                    || "r:link".equals(attr.getNodeName())) {
                String val = attr.getNodeValue();
                if (ridMapping.containsKey(val)) {
                    attr.setNodeValue(ridMapping.get(val));
                }
            }
        }

        // Recurse into children
        NodeList children = element.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if (child instanceof Element childElem) {
                updateRidReferencesRecursive(childElem, ridMapping);
            }
        }
    }

    private static String getRelsPartName(String partName) {
        int lastSlash = partName.lastIndexOf('/');
        if (lastSlash >= 0) {
            return partName.substring(0, lastSlash + 1) + "_rels/"
                    + partName.substring(lastSlash + 1) + ".rels";
        }
        return "_rels/" + partName + ".rels";
    }

    private static String findNextAvailableName(OpcPackage pkg, String originalName) {
        int dotIdx = originalName.lastIndexOf('.');
        if (dotIdx < 0) {
            dotIdx = originalName.length();
        }
        String base = originalName.substring(0, dotIdx);
        String ext = originalName.substring(dotIdx);

        // Strip trailing digits from base to find the pattern
        int numStart = base.length();
        while (numStart > 0 && Character.isDigit(base.charAt(numStart - 1))) {
            numStart--;
        }
        String prefix = base.substring(0, numStart);

        int num = 1;
        while (true) {
            String candidate = prefix + num + ext;
            if (!pkg.hasPart(candidate)) {
                return candidate;
            }
            num++;
        }
    }

    /**
     * Serializes an XML element to UTF-8 bytes with XML declaration.
     *
     * @param root the root element to serialize
     * @return the serialized XML as bytes
     */
    public static byte[] serializeXml(Element root) {
        try {
            var transformer = SecureXml.transformerFactory().newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
            transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            var baos = new ByteArrayOutputStream();
            transformer.transform(new DOMSource(root.getOwnerDocument()), new StreamResult(baos));
            return baos.toByteArray();
        } catch (Exception e) {
            throw new IllegalStateException("Failed to serialize XML", e);
        }
    }

    /**
     * Parses XML bytes into a DOM Document.
     *
     * @param xmlBytes the XML content as bytes
     * @return the parsed Document
     */
    public static Document parseXml(byte[] xmlBytes) {
        try {
            var factory = SecureXml.documentBuilderFactory();
            factory.setNamespaceAware(true);
            return factory.newDocumentBuilder()
                    .parse(new ByteArrayInputStream(xmlBytes));
        } catch (Exception e) {
            throw new IllegalStateException("Failed to parse XML", e);
        }
    }
}
