package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Manages a master slide XML part ({@code ppt/slideMasters/slideMasterN.xml}).
 *
 * <p>Read-only for now; provides access to master slide properties
 * and layout slide relationships.</p>
 */
public final class MasterSlidePart {

    private static final String NS_P =
            "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static final String NS_RELS =
            "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String NS_R =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String REL_TYPE_SLIDE_LAYOUT =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";

    private final OpcPackage pkg;
    private final String partName;
    private Document document;
    private Element root;

    /**
     * Creates a {@code MasterSlidePart} backed by the given OPC package and part name.
     *
     * <p>The master slide XML is eagerly loaded and parsed during construction.</p>
     *
     * @param pkg      the OPC package
     * @param partName the part path (e.g. {@code ppt/slideMasters/slideMaster1.xml})
     * @throws IllegalArgumentException if the part is not found in the package
     */
    public MasterSlidePart(OpcPackage pkg, String partName) {
        this.pkg = Objects.requireNonNull(pkg);
        this.partName = Objects.requireNonNull(partName);
        load();
    }

    /**
     * Load and parse the master slide XML from the package.
     *
     * @throws IllegalArgumentException if the part is not found in the package
     */
    void load() {
        document = pkg.parseXml(partName);
        if (document == null) {
            throw new IllegalArgumentException("Master slide part not found: " + partName);
        }
        root = document.getDocumentElement();
    }

    /**
     * Returns the part name of this master slide.
     *
     * @return the part name (e.g. {@code ppt/slideMasters/slideMaster1.xml})
     */
    public String getPartName() {
        return partName;
    }

    /**
     * Returns the master slide name from {@code <p:cSld name="...">}.
     *
     * @return the master slide name, or an empty string if not specified
     */
    public String getName() {
        NodeList csldNodes = root.getElementsByTagNameNS(NS_P, "cSld");
        if (csldNodes.getLength() > 0) {
            Element csld = (Element) csldNodes.item(0);
            String name = csld.getAttribute("name");
            return name != null ? name : "";
        }
        return "";
    }

    /**
     * Sets the master slide name on {@code <p:cSld name="...">}.
     *
     * @param value the master slide name to set
     */
    public void setName(String value) {
        NodeList csldNodes = root.getElementsByTagNameNS(NS_P, "cSld");
        if (csldNodes.getLength() > 0) {
            Element csld = (Element) csldNodes.item(0);
            csld.setAttribute("name", value);
        }
    }

    /**
     * Returns the layout slide part names of this master, in the order of its
     * {@code p:sldLayoutIdLst}.
     *
     * @return list of resolved layout slide part names
     */
    public List<String> getLayoutPartNames() {
        String relsUri = computeRelsUri(partName);
        Document relsDoc = pkg.parseXml(relsUri);
        if (relsDoc == null) {
            return List.of();
        }
        NodeList rels = relsDoc.getElementsByTagNameNS(NS_RELS, "Relationship");
        Map<String, String> layoutById = new LinkedHashMap<>();
        for (int i = 0; i < rels.getLength(); i++) {
            Element rel = (Element) rels.item(i);
            if (REL_TYPE_SLIDE_LAYOUT.equals(rel.getAttribute("Type"))) {
                String target = rel.getAttribute("Target");
                layoutById.put(rel.getAttribute("Id"), resolveTarget(target));
            }
        }

        // The master's layouts are the entries of p:sldLayoutIdLst, in that order. The order
        // of the .rels file means nothing and PowerPoint writes it in any order, so following
        // it made getLayoutSlides().get(0) an arbitrary layout on real files. A layout that is
        // related but not listed is kept, after the listed ones, as before.
        List<String> result = new ArrayList<>();
        NodeList ids = root.getElementsByTagNameNS(NS_P, "sldLayoutId");
        for (int i = 0; i < ids.getLength(); i++) {
            String layout = layoutById.get(((Element) ids.item(i)).getAttributeNS(NS_R, "id"));
            if (layout != null && !result.contains(layout)) {
                result.add(layout);
            }
        }
        for (String layout : layoutById.values()) {
            if (!result.contains(layout)) {
                result.add(layout);
            }
        }
        return result;
    }

    /**
     * Resolves a relative target path to an absolute part name,
     * normalizing {@code ..} segments.
     *
     * @param target the target path (relative or absolute)
     * @return the resolved absolute part name
     */
    String resolveTarget(String target) {
        if (target.startsWith("/")) {
            return target.replaceFirst("^/+", "");
        }
        String baseDir = partName.contains("/")
                ? partName.substring(0, partName.lastIndexOf('/'))
                : "";
        String combined = baseDir + "/" + target;
        String[] segments = combined.split("/");
        List<String> resolved = new ArrayList<>();
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
     * Saves the master slide XML back to the package.
     */
    public void save() {
        pkg.serializeXml(partName, document);
    }

    /**
     * Computes the {@code .rels} URI for a given part URI.
     *
     * <p>For example, {@code ppt/slideMasters/slideMaster1.xml} becomes
     * {@code ppt/slideMasters/_rels/slideMaster1.xml.rels}.</p>
     */
    private static String computeRelsUri(String uri) {
        int lastSlash = uri.lastIndexOf('/');
        if (lastSlash < 0) {
            return "_rels/" + uri + ".rels";
        }
        String dir = uri.substring(0, lastSlash);
        String fileName = uri.substring(lastSlash + 1);
        return dir + "/_rels/" + fileName + ".rels";
    }
}
