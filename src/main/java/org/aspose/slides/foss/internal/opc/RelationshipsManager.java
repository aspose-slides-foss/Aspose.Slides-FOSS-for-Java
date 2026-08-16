package org.aspose.slides.foss.internal.opc;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.StringWriter;
import java.io.UncheckedIOException;
import java.nio.charset.StandardCharsets;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;
import org.aspose.slides.foss.internal.xml.SecureXml;

/**
 * Manages relationships ({@code .rels}) files in an OPC package.
 *
 * <p>Each part can have an associated {@code .rels} file that defines its
 * relationships to other parts. This class provides methods to:
 * <ul>
 *   <li>Load relationships for a part</li>
 *   <li>Add/remove relationships</li>
 *   <li>Generate unique relationship IDs</li>
 *   <li>Serialize back to XML</li>
 * </ul>
 */
public final class RelationshipsManager {

    /** OPC relationships XML namespace. */
    public static final String RELS_NAMESPACE =
            "http://schemas.openxmlformats.org/package/2006/relationships";

    /** Convenience namespace prefix string for element matching. */
    public static final String RELS_NS = RELS_NAMESPACE;

    /** Common OPC relationship type URIs keyed by short name. */
    public static final Map<String, String> REL_TYPES = Map.ofEntries(
            Map.entry("office_document",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            Map.entry("slide",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"),
            Map.entry("slide_layout",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"),
            Map.entry("slide_master",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"),
            Map.entry("notes_slide",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"),
            Map.entry("notes_master",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"),
            Map.entry("handout_master",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"),
            Map.entry("theme",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
            Map.entry("core_properties",
                    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"),
            Map.entry("extended_properties",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"),
            Map.entry("thumbnail",
                    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"),
            Map.entry("image",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
            Map.entry("hyperlink",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
            Map.entry("chart",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
            Map.entry("oleObject",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"),
            Map.entry("package",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"),
            Map.entry("audio",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"),
            Map.entry("video",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"),
            Map.entry("comments",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"),
            Map.entry("commentAuthors",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors")
    );

    private final OpcPackage opcPackage;
    private final String sourcePart;
    private final String relsPartName;
    private Document document;
    private Element root;
    private final Map<String, Relationship> relationships = new LinkedHashMap<>();

    /**
     * Initialize the relationships manager for a specific part.
     *
     * @param opcPackage the OPC package
     * @param sourcePart the part path whose relationships to manage;
     *                   empty string for root relationships ({@code _rels/.rels})
     */
    public RelationshipsManager(OpcPackage opcPackage, String sourcePart) {
        this.opcPackage = opcPackage;
        this.sourcePart = sourcePart;
        this.relsPartName = getRelsPartName(sourcePart);
        load();
    }

    /**
     * Initialize the relationships manager for the root relationships.
     *
     * @param opcPackage the OPC package
     */
    public RelationshipsManager(OpcPackage opcPackage) {
        this(opcPackage, "");
    }

    /**
     * Get the {@code .rels} file path for a given source part.
     *
     * <p>Examples:
     * <ul>
     *   <li>{@code ""} &rarr; {@code "_rels/.rels"}</li>
     *   <li>{@code "ppt/presentation.xml"} &rarr; {@code "ppt/_rels/presentation.xml.rels"}</li>
     *   <li>{@code "ppt/slides/slide1.xml"} &rarr; {@code "ppt/slides/_rels/slide1.xml.rels"}</li>
     * </ul>
     *
     * @param sourcePart the source part path
     * @return the corresponding {@code .rels} part name
     */
    public static String getRelsPartName(String sourcePart) {
        if (sourcePart == null || sourcePart.isEmpty()) {
            return "_rels/.rels";
        }

        int lastSlash = sourcePart.lastIndexOf('/');
        if (lastSlash >= 0) {
            String directory = sourcePart.substring(0, lastSlash);
            String filename = sourcePart.substring(lastSlash + 1);
            return directory + "/_rels/" + filename + ".rels";
        } else {
            return "_rels/" + sourcePart + ".rels";
        }
    }

    /**
     * Load and parse the {@code .rels} file from the package.
     */
    void load() {
        Optional<byte[]> content = opcPackage.getPart(relsPartName);
        if (content.isPresent()) {
            try {
                DocumentBuilderFactory factory = SecureXml.documentBuilderFactory();
                factory.setNamespaceAware(true);
                DocumentBuilder builder = factory.newDocumentBuilder();
                document = builder.parse(new ByteArrayInputStream(content.get()));
                root = document.getDocumentElement();
                parseRelationships();
            } catch (ParserConfigurationException | SAXException | IOException e) {
                throw new UncheckedIOException("Failed to parse .rels file: " + relsPartName,
                        e instanceof IOException ioe ? ioe : new IOException(e));
            }
        } else {
            try {
                DocumentBuilderFactory factory = SecureXml.documentBuilderFactory();
                factory.setNamespaceAware(true);
                DocumentBuilder builder = factory.newDocumentBuilder();
                document = builder.newDocument();
                document.setXmlStandalone(true);
                root = document.createElementNS(RELS_NAMESPACE, "Relationships");
                document.appendChild(root);
            } catch (ParserConfigurationException e) {
                throw new IllegalStateException("Failed to create XML document", e);
            }
        }
    }

    /**
     * Parse relationships from the loaded XML.
     */
    void parseRelationships() {
        relationships.clear();
        NodeList relElements = root.getElementsByTagNameNS(RELS_NAMESPACE, "Relationship");
        for (int i = 0; i < relElements.getLength(); i++) {
            Element relElem = (Element) relElements.item(i);
            String id = relElem.getAttribute("Id");
            String type = relElem.getAttribute("Type");
            String target = relElem.getAttribute("Target");
            String targetMode = relElem.hasAttribute("TargetMode")
                    ? relElem.getAttribute("TargetMode") : null;

            Relationship rel = new Relationship(
                    id, type, target, Optional.ofNullable(targetMode));
            relationships.put(id, rel);
        }
    }

    /**
     * Get a relationship by ID.
     *
     * @param relId the relationship ID (e.g., "rId1")
     * @return the relationship, or empty if not found
     */
    public Optional<Relationship> getRelationship(String relId) {
        return Optional.ofNullable(relationships.get(relId));
    }

    /**
     * Get all relationships of a specific type.
     *
     * @param relType the relationship type URI
     * @return list of matching relationships
     */
    public List<Relationship> getRelationshipsByType(String relType) {
        return relationships.values().stream()
                .filter(r -> r.type().equals(relType))
                .toList();
    }

    /**
     * Get all relationships.
     *
     * @return list of all relationships
     */
    public List<Relationship> getAllRelationships() {
        return List.copyOf(relationships.values());
    }

    /**
     * Generate a unique relationship ID.
     *
     * @return a new unique ID of the form "rIdN"
     */
    String generateRelId() {
        Set<String> existingIds = relationships.keySet();
        int counter = 1;
        while (true) {
            String relId = "rId" + counter;
            if (!existingIds.contains(relId)) {
                return relId;
            }
            counter++;
        }
    }

    /**
     * Add a new relationship.
     *
     * @param relType    the relationship type URI
     * @param target     the target part path (relative to source)
     * @param relId      optional specific ID; if {@code null}, auto-generated
     * @param targetMode optional target mode ("External" for external links)
     * @return the relationship ID
     */
    public String addRelationship(String relType, String target,
                                  String relId, String targetMode) {
        if (relId == null) {
            relId = generateRelId();
        }

        Relationship rel = new Relationship(
                relId, relType, target, Optional.ofNullable(targetMode));
        relationships.put(relId, rel);

        Element relElem = document.createElementNS(RELS_NAMESPACE, "Relationship");
        relElem.setAttribute("Id", relId);
        relElem.setAttribute("Type", relType);
        relElem.setAttribute("Target", target);
        if (targetMode != null) {
            relElem.setAttribute("TargetMode", targetMode);
        }
        root.appendChild(relElem);

        return relId;
    }

    /**
     * Add a new relationship with auto-generated ID and no target mode.
     *
     * @param relType the relationship type URI
     * @param target  the target part path (relative to source)
     * @return the relationship ID
     */
    public String addRelationship(String relType, String target) {
        return addRelationship(relType, target, null, null);
    }

    /**
     * Remove a relationship by ID.
     *
     * @param relId the relationship ID to remove
     * @return {@code true} if removed, {@code false} if not found
     */
    public boolean removeRelationship(String relId) {
        if (!relationships.containsKey(relId)) {
            return false;
        }

        relationships.remove(relId);

        NodeList relElements = root.getElementsByTagNameNS(RELS_NAMESPACE, "Relationship");
        for (int i = 0; i < relElements.getLength(); i++) {
            Element relElem = (Element) relElements.item(i);
            if (relId.equals(relElem.getAttribute("Id"))) {
                root.removeChild(relElem);
                return true;
            }
        }

        return true;
    }

    /**
     * Save the relationships back to the package.
     */
    public void save() {
        if (relationships.isEmpty()) {
            opcPackage.deletePart(relsPartName);
            return;
        }

        try {
            TransformerFactory tf = SecureXml.transformerFactory();
            Transformer transformer = tf.newTransformer();
            transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
            transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            transformer.transform(new DOMSource(document), new StreamResult(baos));
            opcPackage.setPart(relsPartName, baos.toByteArray());
        } catch (TransformerException e) {
            throw new IllegalStateException("Failed to serialize .rels XML", e);
        }
    }

    /**
     * Get the {@code .rels} file part name.
     *
     * @return the part name for the relationships file
     */
    public String getPartName() {
        return relsPartName;
    }
}
