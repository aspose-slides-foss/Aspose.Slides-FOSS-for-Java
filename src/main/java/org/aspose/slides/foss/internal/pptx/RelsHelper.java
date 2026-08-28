package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Lightweight helper for reading and writing OPC relationship files
 * using {@link OpcPackage}.
 *
 * <p>Unlike {@link org.aspose.slides.foss.internal.opc.RelationshipsManager}, this helper
 * operates on the {@code internal.pptx.OpcPackage} which stores parts as byte arrays
 * with XML parse/serialize support.</p>
 */
public final class RelsHelper {

    private static final String REL_NS =
            "http://schemas.openxmlformats.org/package/2006/relationships";

    private final OpcPackage pkg;
    private final String sourcePart;
    private final String relsPartName;
    private Document doc;
    private Element root;
    private final Map<String, RelEntry> relationships = new LinkedHashMap<>();
    private int nextId = 1;

    /**
     * Creates a relationships helper for the given part.
     *
     * @param pkg        the OPC package
     * @param sourcePart the part whose relationships to manage
     */
    public RelsHelper(OpcPackage pkg, String sourcePart) {
        this.pkg = pkg;
        this.sourcePart = sourcePart;
        this.relsPartName = getRelsPartName(sourcePart);
        load();
    }

    /**
     * Returns the {@code .rels} part name for a given source part.
     *
     * @param sourcePart the source part path
     * @return the corresponding .rels path
     */
    public static String getRelsPartName(String sourcePart) {
        if (sourcePart == null || sourcePart.isEmpty()) {
            return "_rels/.rels";
        }
        int lastSlash = sourcePart.lastIndexOf('/');
        if (lastSlash >= 0) {
            return sourcePart.substring(0, lastSlash) + "/_rels/"
                    + sourcePart.substring(lastSlash + 1) + ".rels";
        }
        return "_rels/" + sourcePart + ".rels";
    }

    private void load() {
        doc = pkg.parseXml(relsPartName);
        if (doc != null) {
            root = doc.getDocumentElement();
            parseRelationships();
        } else {
            doc = OpcPackage.newDocument();
            doc.setXmlStandalone(true);
            root = doc.createElementNS(REL_NS, "Relationships");
            doc.appendChild(root);
        }
    }

    private void parseRelationships() {
        relationships.clear();
        NodeList relElements = root.getElementsByTagNameNS(REL_NS, "Relationship");
        if (relElements.getLength() == 0) {
            // Try without namespace (some files may not use NS prefix)
            relElements = root.getElementsByTagName("Relationship");
        }
        for (int i = 0; i < relElements.getLength(); i++) {
            Element el = (Element) relElements.item(i);
            String id = el.getAttribute("Id");
            String type = el.getAttribute("Type");
            String target = el.getAttribute("Target");
            String targetMode = el.hasAttribute("TargetMode")
                    ? el.getAttribute("TargetMode") : null;
            relationships.put(id, new RelEntry(id, type, target, targetMode));

            // Track max rId number
            if (id.startsWith("rId")) {
                try {
                    int num = Integer.parseInt(id.substring(3));
                    if (num >= nextId) {
                        nextId = num + 1;
                    }
                } catch (NumberFormatException ignored) {
                    // Non-numeric value; use default
                }
            }
        }
    }

    /**
     * Returns all relationships.
     *
     * @return list of all relationship entries
     */
    public List<RelEntry> getAllRelationships() {
        return List.copyOf(relationships.values());
    }

    /**
     * Adds a new relationship with auto-generated ID.
     *
     * @param type   the relationship type URI
     * @param target the target path
     * @return the generated relationship ID
     */
    public String addRelationship(String type, String target) {
        return addRelationship(type, target, null);
    }

    /**
     * Adds a new relationship with optional target mode.
     *
     * @param type       the relationship type URI
     * @param target     the target path
     * @param targetMode optional target mode (e.g. "External"), or {@code null}
     * @return the generated relationship ID
     */
    public String addRelationship(String type, String target, String targetMode) {
        String id = "rId" + nextId++;
        relationships.put(id, new RelEntry(id, type, target, targetMode));

        Element relElem = doc.createElementNS(REL_NS, "Relationship");
        relElem.setAttribute("Id", id);
        relElem.setAttribute("Type", type);
        relElem.setAttribute("Target", target);
        if (targetMode != null) {
            relElem.setAttribute("TargetMode", targetMode);
        }
        root.appendChild(relElem);

        return id;
    }

    /**
     * Removes a relationship by ID.
     *
     * @param id the relationship ID to remove
     * @return {@code true} if a relationship with that ID was present and has been removed
     */
    public boolean removeRelationship(String id) {
        if (relationships.remove(id) == null) {
            return false;
        }
        NodeList relElements = root.getElementsByTagName("*");
        for (int i = 0; i < relElements.getLength(); i++) {
            Element el = (Element) relElements.item(i);
            if ("Relationship".equals(el.getLocalName()) && id.equals(el.getAttribute("Id"))) {
                root.removeChild(el);
                break;
            }
        }
        return true;
    }

    /**
     * Removes every internal relationship whose target resolves to the given part.
     *
     * @param partUri the part being pointed at, as an absolute package path
     * @return {@code true} if at least one relationship was removed
     */
    public boolean removeRelationshipsTo(String partUri) {
        boolean removed = false;
        for (RelEntry entry : getAllRelationships()) {
            if ("External".equals(entry.targetMode())) {
                continue;
            }
            if (partUri.equals(SlidePart.resolveTargetStatic(sourcePart, entry.target()))) {
                removed |= removeRelationship(entry.id());
            }
        }
        return removed;
    }

    /**
     * Returns the parts inside the package that this part's relationships point at.
     *
     * @return absolute package paths of the internal relationship targets, in declaration order
     */
    public List<String> internalTargets() {
        var targets = new ArrayList<String>();
        for (RelEntry entry : relationships.values()) {
            if ("External".equals(entry.targetMode())) {
                continue;
            }
            targets.add(SlidePart.resolveTargetStatic(sourcePart, entry.target()));
        }
        return targets;
    }

    /**
     * Returns the part that owns a {@code .rels} part.
     *
     * @param relsPartName the {@code .rels} path
     * @return the owning part path; the empty string for the package-level {@code _rels/.rels}
     */
    public static String getSourcePartName(String relsPartName) {
        int marker = relsPartName.lastIndexOf("_rels/");
        if (marker < 0 || !relsPartName.endsWith(".rels")) {
            return relsPartName;
        }
        String dir = relsPartName.substring(0, marker);
        String file = relsPartName.substring(marker + "_rels/".length());
        return dir + file.substring(0, file.length() - ".rels".length());
    }

    /**
     * Saves the relationships back to the package.
     */
    public void save() {
        if (relationships.isEmpty()) {
            pkg.removePart(relsPartName);
            return;
        }
        pkg.serializeXml(relsPartName, doc);
    }

    /**
     * A single relationship entry.
     *
     * @param id         the relationship ID
     * @param type       the relationship type URI
     * @param target     the target path
     * @param targetMode the target mode, or {@code null} for internal
     */
    public record RelEntry(String id, String type, String target, String targetMode) {
    }
}
