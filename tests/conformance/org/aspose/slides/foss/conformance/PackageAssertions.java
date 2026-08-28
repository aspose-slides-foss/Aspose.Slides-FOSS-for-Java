package org.aspose.slides.foss.conformance;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.namespace.NamespaceContext;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Assertions about the structure of a produced OPC package.
 *
 * <p>Every method here reads the ZIP and the XML inside it. None of them calls the library that
 * wrote the file. See {@code tests/conformance/README.md} for why that distinction is the whole
 * point of this package.</p>
 */
public final class PackageAssertions {

    private static final Map<String, String> PREFIXES = new LinkedHashMap<>();

    static {
        PREFIXES.put("ct", PptxPackage.NS_CONTENT_TYPES);
        PREFIXES.put("rel", PptxPackage.NS_RELS_PART);
        PREFIXES.put("r", PptxPackage.NS_REL);
        PREFIXES.put("p", PptxPackage.NS_P);
        PREFIXES.put("a", PptxPackage.NS_A);
        PREFIXES.put("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
        PREFIXES.put("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");
        PREFIXES.put("p188", "http://schemas.microsoft.com/office/powerpoint/2018/8/main");
        PREFIXES.put("ep", "http://schemas.openxmlformats.org/officeDocument/2006/"
                + "extended-properties");
        PREFIXES.put("dc", "http://purl.org/dc/elements/1.1/");
        PREFIXES.put("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    }

    private PackageAssertions() {
    }

    // ---------------------------------------------------------------- package-wide rules

    /**
     * Runs every package-wide rule: relationship references resolve, relationship targets exist,
     * every part has a content type, and no content-type override names a part that is missing.
     *
     * @param pkg the package
     * @throws IOException if a part cannot be read
     */
    public static void assertPackageIsSelfConsistent(PptxPackage pkg) throws IOException {
        assertEveryRelationshipReferenceResolves(pkg);
        assertEveryInternalRelationshipTargetExists(pkg);
        assertEveryPartHasAContentType(pkg);
        assertNoContentTypeOverrideNamesAMissingPart(pkg);
    }

    /**
     * Asserts that every {@code r:id}, {@code r:embed} and {@code r:link} in every XML part is
     * declared as a {@code Relationship Id} in that part's own {@code .rels}.
     *
     * @param pkg the package
     * @throws IOException if a part cannot be read
     */
    public static void assertEveryRelationshipReferenceResolves(PptxPackage pkg) throws IOException {
        List<String> unresolved = new ArrayList<>();
        for (String partName : pkg.xmlPartNames()) {
            if (partName.endsWith(".rels")) {
                continue;
            }
            Set<String> referenced = relationshipReferencesIn(pkg.root(partName));
            if (referenced.isEmpty()) {
                continue;
            }
            Set<String> declared = relationshipIdsOf(pkg, partName);
            for (String id : referenced) {
                if (!declared.contains(id)) {
                    unresolved.add(partName + " references relationship id '" + id
                            + "' which is not declared in "
                            + PptxPackage.relsPartNameFor(partName)
                            + " (declared: " + declared + ")");
                }
            }
        }
        assertThat(unresolved)
                .as("dangling relationship references in %s", pkg.path())
                .isEmpty();
    }

    /**
     * Asserts that every internal relationship target names a part that is present in the package.
     *
     * @param pkg the package
     * @throws IOException if a part cannot be read
     */
    public static void assertEveryInternalRelationshipTargetExists(PptxPackage pkg)
            throws IOException {
        List<String> missing = new ArrayList<>();
        for (String relsName : pkg.xmlPartNames()) {
            if (!relsName.endsWith(".rels")) {
                continue;
            }
            String owner = ownerOfRelsPart(relsName);
            NodeList rels = pkg.root(relsName)
                    .getElementsByTagNameNS(PptxPackage.NS_RELS_PART, "Relationship");
            for (int i = 0; i < rels.getLength(); i++) {
                Element rel = (Element) rels.item(i);
                if ("External".equals(rel.getAttribute("TargetMode"))) {
                    continue;
                }
                String target = rel.getAttribute("Target");
                String resolved = PptxPackage.resolveTarget(owner, target);
                if (!pkg.hasPart(resolved)) {
                    missing.add(relsName + " declares " + rel.getAttribute("Id")
                            + " -> " + target + " which resolves to '" + resolved
                            + "', a part that is not in the package");
                }
            }
        }
        assertThat(missing)
                .as("relationships pointing at parts that do not exist in %s", pkg.path())
                .isEmpty();
    }

    /**
     * Asserts that every part resolves a content type through an {@code Override} or a
     * {@code Default} in {@code [Content_Types].xml}.
     *
     * @param pkg the package
     * @throws IOException if a part cannot be read
     */
    public static void assertEveryPartHasAContentType(PptxPackage pkg) throws IOException {
        Map<String, String> defaults = contentTypeDefaults(pkg);
        Map<String, String> overrides = contentTypeOverrides(pkg);
        List<String> untyped = new ArrayList<>();
        for (String partName : pkg.entryNames()) {
            if (partName.equals(PptxPackage.CONTENT_TYPES_PART)) {
                continue;
            }
            if (overrides.containsKey("/" + partName)) {
                continue;
            }
            String extension = extensionOf(partName);
            if (extension != null && defaults.containsKey(extension)) {
                continue;
            }
            untyped.add(partName);
        }
        assertThat(untyped)
                .as("parts with no content type in %s (defaults: %s)", pkg.path(), defaults.keySet())
                .isEmpty();
    }

    /**
     * Asserts that no {@code Override} in {@code [Content_Types].xml} names a part that has been
     * removed from the package.
     *
     * @param pkg the package
     * @throws IOException if a part cannot be read
     */
    public static void assertNoContentTypeOverrideNamesAMissingPart(PptxPackage pkg)
            throws IOException {
        List<String> dangling = new ArrayList<>();
        for (String partName : contentTypeOverrides(pkg).keySet()) {
            if (!pkg.hasPart(partName.substring(1))) {
                dangling.add(partName);
            }
        }
        assertThat(dangling)
                .as("content-type overrides naming parts that are not in %s", pkg.path())
                .isEmpty();
    }

    // ---------------------------------------------------------------- slides

    /**
     * Counts the slides the package actually declares, by reading
     * {@code ppt/presentation.xml}'s {@code p:sldIdLst}.
     *
     * <p>Never count {@code ppt/slides/slideN.xml} entries instead: a slide part that is not
     * referenced from {@code sldIdLst} is not a slide, and counting files is exactly how a
     * whole class of packaging defect stays invisible.</p>
     *
     * @param pkg the package
     * @return the number of {@code p:sldId} children of {@code p:sldIdLst}
     * @throws IOException if the presentation part cannot be read
     */
    public static int registeredSlideCount(PptxPackage pkg) throws IOException {
        return selectNodes(pkg, "ppt/presentation.xml", "//p:sldIdLst/p:sldId").size();
    }

    /**
     * Asserts, four independent ways, that the package declares exactly {@code expected} slides:
     * the {@code p:sldIdLst} entries, the presentation-level slide relationships, the
     * content-type overrides and the slide parts themselves must all agree.
     *
     * @param pkg      the package
     * @param expected the expected slide count
     * @throws IOException if a part cannot be read
     */
    public static void assertRegisteredSlideCount(PptxPackage pkg, int expected) throws IOException {
        assertThat(registeredSlideCount(pkg))
                .as("p:sldId entries in ppt/presentation.xml of %s", pkg.path())
                .isEqualTo(expected);

        int relationships = 0;
        NodeList rels = pkg.root("ppt/_rels/presentation.xml.rels")
                .getElementsByTagNameNS(PptxPackage.NS_RELS_PART, "Relationship");
        for (int i = 0; i < rels.getLength(); i++) {
            if (((Element) rels.item(i)).getAttribute("Type").endsWith("/relationships/slide")) {
                relationships++;
            }
        }
        assertThat(relationships)
                .as("slide relationships in ppt/_rels/presentation.xml.rels of %s", pkg.path())
                .isEqualTo(expected);

        long overrides = contentTypeOverrides(pkg).entrySet().stream()
                .filter(e -> e.getValue().equals(
                        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"))
                .count();
        assertThat(overrides)
                .as("slide content-type overrides in %s", pkg.path())
                .isEqualTo(expected);

        long parts = pkg.entryNames().stream()
                .filter(n -> n.matches("ppt/slides/slide\\d+\\.xml"))
                .count();
        assertThat(parts)
                .as("ppt/slides/slideN.xml parts in %s", pkg.path())
                .isEqualTo(expected);
    }

    // ---------------------------------------------------------------- XML shape

    /**
     * Evaluates an XPath over a part and returns the matching elements.
     *
     * <p>Bind namespaces with the prefixes {@code p}, {@code a}, {@code r}, {@code ct},
     * {@code rel}, {@code p14}, {@code p15}, {@code p188}, {@code ep}, {@code dc} and
     * {@code cp}.</p>
     *
     * @param pkg      the package
     * @param partName the part to search
     * @param xpath    the expression
     * @return the matching elements, in document order
     * @throws IOException if the part cannot be read
     */
    public static List<Element> selectNodes(PptxPackage pkg, String partName, String xpath)
            throws IOException {
        Document doc = pkg.xml(partName);
        XPath engine = XPathFactory.newInstance().newXPath();
        engine.setNamespaceContext(namespaceContext());
        NodeList nodes;
        try {
            nodes = (NodeList) engine.evaluate(xpath, doc, XPathConstants.NODESET);
        } catch (XPathExpressionException e) {
            throw new IllegalArgumentException("bad XPath: " + xpath, e);
        }
        List<Element> elements = new ArrayList<>();
        for (int i = 0; i < nodes.getLength(); i++) {
            if (nodes.item(i) instanceof Element element) {
                elements.add(element);
            }
        }
        return elements;
    }

    /**
     * Asserts that a part contains exactly one element matching {@code xpath}, and that it carries
     * the given attributes with the given values.
     *
     * @param pkg      the package
     * @param partName the part to search
     * @param xpath    the expression selecting the element
     * @param attrs    attribute name to expected value; names may be prefixed, e.g. {@code r:embed}
     * @return the matched element
     * @throws IOException if the part cannot be read
     */
    public static Element assertElementWithAttributes(PptxPackage pkg, String partName,
                                                      String xpath, Map<String, String> attrs)
            throws IOException {
        List<Element> found = selectNodes(pkg, partName, xpath);
        assertThat(found)
                .as("elements matching %s in %s of %s%n%s",
                        xpath, partName, pkg.path(), pkg.text(partName))
                .hasSize(1);
        Element element = found.get(0);
        for (Map.Entry<String, String> expected : attrs.entrySet()) {
            String name = expected.getKey();
            String actual;
            int colon = name.indexOf(':');
            if (colon < 0) {
                actual = element.hasAttribute(name) ? element.getAttribute(name) : null;
            } else {
                String uri = PREFIXES.get(name.substring(0, colon));
                String local = name.substring(colon + 1);
                actual = element.hasAttributeNS(uri, local)
                        ? element.getAttributeNS(uri, local) : null;
            }
            assertThat(actual)
                    .as("attribute %s of %s in %s of %s", name, xpath, partName, pkg.path())
                    .isEqualTo(expected.getValue());
        }
        return element;
    }

    /**
     * @param parent an element
     * @return the qualified names of its element children, in document order
     */
    public static List<String> childNames(Element parent) {
        List<String> names = new ArrayList<>();
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element child) {
                names.add(child.getTagName());
            }
        }
        return names;
    }

    /**
     * Asserts that the element children of {@code parent} appear in the relative order the schema
     * requires, and that no child appears which the schema does not allow there.
     *
     * @param parent      the element whose children are checked
     * @param schemaOrder the qualified child names, in schema sequence order
     */
    public static void assertChildrenInSchemaOrder(Element parent, String... schemaOrder) {
        List<String> order = List.of(schemaOrder);
        List<String> actual = childNames(parent);
        for (String name : actual) {
            assertThat(order)
                    .as("child <%s> of <%s> is not in the schema sequence for that element",
                            name, parent.getTagName())
                    .contains(name);
        }
        int previous = -1;
        for (String name : actual) {
            int index = order.indexOf(name);
            assertThat(index)
                    .as("children of <%s> are %s, which is not the schema order %s",
                            parent.getTagName(), actual, order)
                    .isGreaterThanOrEqualTo(previous);
            previous = index;
        }
    }

    // ---------------------------------------------------------------- content types

    /**
     * @param pkg the package
     * @return extension (lower case, no dot) to content type, from {@code Default} declarations
     * @throws IOException if the content-types part cannot be read
     */
    public static Map<String, String> contentTypeDefaults(PptxPackage pkg) throws IOException {
        Map<String, String> defaults = new LinkedHashMap<>();
        NodeList nodes = pkg.root(PptxPackage.CONTENT_TYPES_PART)
                .getElementsByTagNameNS(PptxPackage.NS_CONTENT_TYPES, "Default");
        for (int i = 0; i < nodes.getLength(); i++) {
            Element element = (Element) nodes.item(i);
            defaults.put(element.getAttribute("Extension").toLowerCase(Locale.ROOT),
                    element.getAttribute("ContentType"));
        }
        return defaults;
    }

    /**
     * @param pkg the package
     * @return part name (with leading slash) to content type, from {@code Override} declarations
     * @throws IOException if the content-types part cannot be read
     */
    public static Map<String, String> contentTypeOverrides(PptxPackage pkg) throws IOException {
        Map<String, String> overrides = new LinkedHashMap<>();
        NodeList nodes = pkg.root(PptxPackage.CONTENT_TYPES_PART)
                .getElementsByTagNameNS(PptxPackage.NS_CONTENT_TYPES, "Override");
        for (int i = 0; i < nodes.getLength(); i++) {
            Element element = (Element) nodes.item(i);
            overrides.put(element.getAttribute("PartName"), element.getAttribute("ContentType"));
        }
        return overrides;
    }

    /**
     * @param pkg the package
     * @return the content type declared for {@code /ppt/presentation.xml}
     * @throws IOException if the content-types part cannot be read
     */
    public static String mainPartContentType(PptxPackage pkg) throws IOException {
        return contentTypeOverrides(pkg).get("/ppt/presentation.xml");
    }

    // ---------------------------------------------------------------- internals

    private static Set<String> relationshipReferencesIn(Element root) {
        Set<String> ids = new LinkedHashSet<>();
        collectRelationshipReferences(root, ids);
        return ids;
    }

    private static void collectRelationshipReferences(Node node, Set<String> ids) {
        if (node instanceof Element element) {
            NamedNodeMap attributes = element.getAttributes();
            for (int i = 0; i < attributes.getLength(); i++) {
                Attr attr = (Attr) attributes.item(i);
                if (PptxPackage.NS_REL.equals(attr.getNamespaceURI())
                        && !attr.getValue().isEmpty()) {
                    ids.add(attr.getValue());
                }
            }
        }
        NodeList children = node.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            collectRelationshipReferences(children.item(i), ids);
        }
    }

    private static Set<String> relationshipIdsOf(PptxPackage pkg, String partName)
            throws IOException {
        Set<String> ids = new LinkedHashSet<>();
        String relsName = PptxPackage.relsPartNameFor(partName);
        if (!pkg.hasPart(relsName)) {
            return ids;
        }
        NodeList rels = pkg.root(relsName)
                .getElementsByTagNameNS(PptxPackage.NS_RELS_PART, "Relationship");
        for (int i = 0; i < rels.getLength(); i++) {
            ids.add(((Element) rels.item(i)).getAttribute("Id"));
        }
        return ids;
    }

    private static String ownerOfRelsPart(String relsName) {
        int marker = relsName.lastIndexOf("_rels/");
        String dir = relsName.substring(0, marker);
        String file = relsName.substring(marker + "_rels/".length());
        String owner = file.substring(0, file.length() - ".rels".length());
        return dir + owner;
    }

    private static String extensionOf(String partName) {
        int dot = partName.lastIndexOf('.');
        int slash = partName.lastIndexOf('/');
        if (dot < 0 || dot < slash) {
            return null;
        }
        return partName.substring(dot + 1).toLowerCase(Locale.ROOT);
    }

    private static NamespaceContext namespaceContext() {
        return new NamespaceContext() {
            @Override
            public String getNamespaceURI(String prefix) {
                return PREFIXES.getOrDefault(prefix, "");
            }

            @Override
            public String getPrefix(String namespaceURI) {
                for (Map.Entry<String, String> e : PREFIXES.entrySet()) {
                    if (e.getValue().equals(namespaceURI)) {
                        return e.getKey();
                    }
                }
                return null;
            }

            @Override
            public Iterator<String> getPrefixes(String namespaceURI) {
                String prefix = getPrefix(namespaceURI);
                return prefix == null ? List.<String>of().iterator() : List.of(prefix).iterator();
            }
        };
    }
}
