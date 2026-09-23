package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Manages a layout slide XML part ({@code ppt/slideLayouts/slideLayoutN.xml}).
 *
 * <p>Read-only for now; provides access to layout properties such as
 * layout type, layout name, and the associated master slide part.</p>
 */
public final class LayoutSlidePart {

    private static final String NS_P =
            "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static final String NS_RELS =
            "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String REL_TYPE_SLIDE_MASTER =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

    /** OOXML layout type values mapped to SlideLayoutType enum value strings. */
    private static final Map<String, String> LAYOUT_TYPE_MAP = Map.ofEntries(
            Map.entry("blank", "Blank"),
            Map.entry("chart", "Chart"),
            Map.entry("chartAndTx", "ChartAndText"),
            Map.entry("clipArtAndTx", "ClipArtAndText"),
            Map.entry("clipArtAndVertTx", "ClipArtAndVerticalText"),
            Map.entry("cust", "Custom"),
            Map.entry("dgm", "Diagram"),
            Map.entry("fourObj", "FourObjects"),
            Map.entry("mediaAndTx", "MediaAndText"),
            Map.entry("obj", "Object"),
            Map.entry("objAndTx", "ObjectAndText"),
            Map.entry("objAndTwoObj", "ObjectAndTwoObject"),
            Map.entry("objOnly", "ObjectOnly"),
            Map.entry("objOverTx", "ObjectOverText"),
            Map.entry("objTx", "TitleObjectAndCaption"),
            Map.entry("picTx", "PictureAndCaption"),
            Map.entry("secHead", "SectionHeader"),
            Map.entry("tbl", "Table"),
            Map.entry("title", "Title"),
            Map.entry("titleOnly", "TitleOnly"),
            Map.entry("twoColTx", "TwoColumnText"),
            Map.entry("twoObj", "TwoObjects"),
            Map.entry("twoObjAndObj", "TwoObjectsAndObject"),
            Map.entry("twoObjAndTx", "TwoObjectsAndText"),
            Map.entry("twoObjOverTx", "TwoObjectsOverText"),
            Map.entry("twoTxTwoObj", "TwoTextAndTwoObjects"),
            Map.entry("tx", "Text"),
            Map.entry("txAndChart", "TextAndChart"),
            Map.entry("txAndClipArt", "TextAndClipArt"),
            Map.entry("txAndMedia", "TextAndMedia"),
            Map.entry("txAndObj", "TextAndObject"),
            Map.entry("txAndTwoObj", "TextAndTwoObjects"),
            Map.entry("txOverObj", "TextOverObject"),
            Map.entry("vertTitleAndTx", "VerticalTitleAndText"),
            Map.entry("vertTitleAndTxOverChart", "VerticalTitleAndTextOverChart"),
            Map.entry("vertTx", "VerticalText")
    );

    private final OpcPackage pkg;
    private final String partName;
    private Document document;
    private Element root;

    /**
     * Creates a {@code LayoutSlidePart} backed by the given OPC package and part name.
     *
     * <p>The layout XML is eagerly loaded and parsed during construction.</p>
     *
     * @param pkg      the OPC package
     * @param partName the part path (e.g. {@code ppt/slideLayouts/slideLayout1.xml})
     * @throws IllegalArgumentException if the part is not found in the package
     */
    public LayoutSlidePart(OpcPackage pkg, String partName) {
        this.pkg = Objects.requireNonNull(pkg);
        this.partName = Objects.requireNonNull(partName);
        load();
    }

    /**
     * Load and parse the layout slide XML from the package.
     *
     * @throws IllegalArgumentException if the part is not found in the package
     */
    void load() {
        document = pkg.parseXml(partName);
        if (document == null) {
            throw new IllegalArgumentException("Layout slide part not found: " + partName);
        }
        root = document.getDocumentElement();
    }

    /**
     * Returns the part name of this layout slide.
     *
     * @return the part name (e.g. {@code ppt/slideLayouts/slideLayout1.xml})
     */
    public String getPartName() {
        return partName;
    }

    /**
     * Returns the layout name from {@code <p:cSld name="...">}.
     *
     * @return the layout name, or an empty string if not specified
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
     * Sets the layout name on {@code <p:cSld name="...">}.
     *
     * @param value the layout name to set
     */
    public void setName(String value) {
        NodeList csldNodes = root.getElementsByTagNameNS(NS_P, "cSld");
        if (csldNodes.getLength() > 0) {
            Element csld = (Element) csldNodes.item(0);
            csld.setAttribute("name", value);
        }
    }

    /**
     * Returns the raw layout type string from the {@code type} attribute
     * of the {@code <p:sldLayout>} element.
     *
     * @return the raw layout type string, defaulting to {@code "cust"} if not specified
     */
    public String getLayoutTypeRaw() {
        String type = root.getAttribute("type");
        return (type == null || type.isEmpty()) ? "cust" : type;
    }

    /**
     * Returns the {@code SlideLayoutType} enum value string for this layout.
     *
     * <p>Maps the raw OOXML type attribute to a friendly enum name via
     * the layout type map. Unknown types default to {@code "Custom"}.</p>
     *
     * @return the layout type enum value string
     */
    public String getLayoutTypeValue() {
        return LAYOUT_TYPE_MAP.getOrDefault(getLayoutTypeRaw(), "Custom");
    }

    /**
     * Resolves the master slide part name from this layout's relationships.
     *
     * @return the master slide part name, or {@code null} if no relationship is found
     */
    public String getMasterPartName() {
        String relsUri = computeRelsUri(partName);
        Document relsDoc = pkg.parseXml(relsUri);
        if (relsDoc == null) {
            return null;
        }
        NodeList rels = relsDoc.getElementsByTagNameNS(NS_RELS, "Relationship");
        for (int i = 0; i < rels.getLength(); i++) {
            Element rel = (Element) rels.item(i);
            if (REL_TYPE_SLIDE_MASTER.equals(rel.getAttribute("Type"))) {
                String target = rel.getAttribute("Target");
                return resolveTarget(target);
            }
        }
        return null;
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
            // Absolute path — strip leading slash(es)
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
     * Saves the layout slide XML back to the package.
     */
    public void save() {
        pkg.serializeXml(partName, document);
    }

    /**
     * Computes the {@code .rels} URI for a given part URI.
     *
     * <p>For example, {@code ppt/slideLayouts/slideLayout1.xml} becomes
     * {@code ppt/slideLayouts/_rels/slideLayout1.xml.rels}.</p>
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
