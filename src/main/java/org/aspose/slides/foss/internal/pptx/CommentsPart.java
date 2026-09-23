package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Manages a slide comments XML part ({@code ppt/comments/slideN.xml}).
 *
 * <p>One file exists per slide that has comments. This class provides methods
 * for loading, creating, and manipulating comment elements within that part.</p>
 */
public final class CommentsPart {

    private static final String NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static final String REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String REL_TYPE_COMMENTS =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    private static final String CONTENT_TYPE_COMMENTS =
            "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
    private static final String CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

    private final OpcPackage pkg;
    private final String partName;
    private Document doc;
    private Element root;

    /**
     * Creates a CommentsPart for the given part in the package.
     *
     * @param pkg      the OPC package
     * @param partName the part path (e.g. {@code "ppt/comments/slide1.xml"})
     * @throws IllegalArgumentException if the part does not exist
     */
    public CommentsPart(OpcPackage pkg, String partName) {
        this.pkg = Objects.requireNonNull(pkg);
        this.partName = Objects.requireNonNull(partName);
        load();
    }

    /**
     * Returns the part name (URI) of this comments part.
     *
     * @return the part name
     */
    public String getPartName() {
        return partName;
    }

    // ---- Query methods ----

    /**
     * Returns all comments in this part.
     *
     * @return list of comment data wrappers
     */
    public List<CommentData> getComments() {
        var result = new ArrayList<CommentData>();
        NodeList nodes = root.getElementsByTagNameNS(NS_P, "cm");
        for (int i = 0; i < nodes.getLength(); i++) {
            result.add(new CommentData((Element) nodes.item(i)));
        }
        return result;
    }

    /**
     * Returns all comments by the given author.
     *
     * @param authorId the author ID to filter by
     * @return list of matching comments
     */
    public List<CommentData> getCommentsByAuthor(int authorId) {
        var result = new ArrayList<CommentData>();
        NodeList nodes = root.getElementsByTagNameNS(NS_P, "cm");
        for (int i = 0; i < nodes.getLength(); i++) {
            Element el = (Element) nodes.item(i);
            String aid = el.getAttribute("authorId");
            if (!aid.isEmpty() && Integer.parseInt(aid) == authorId) {
                result.add(new CommentData(el));
            }
        }
        return result;
    }

    /**
     * Finds a comment by author ID and index.
     *
     * @param authorId the author ID
     * @param idx      the comment index
     * @return the matching comment, or empty
     */
    public Optional<CommentData> findCommentByIdx(int authorId, int idx) {
        NodeList nodes = root.getElementsByTagNameNS(NS_P, "cm");
        for (int i = 0; i < nodes.getLength(); i++) {
            Element el = (Element) nodes.item(i);
            String aid = el.getAttribute("authorId");
            String idxStr = el.getAttribute("idx");
            if (!aid.isEmpty() && Integer.parseInt(aid) == authorId
                    && !idxStr.isEmpty() && Integer.parseInt(idxStr) == idx) {
                return Optional.of(new CommentData(el));
            }
        }
        return Optional.empty();
    }

    /**
     * Finds a comment by index across all authors (for parentCmId lookup).
     *
     * @param idx the comment index
     * @return the matching comment, or empty
     */
    public Optional<CommentData> findCommentByIdxAll(int idx) {
        NodeList nodes = root.getElementsByTagNameNS(NS_P, "cm");
        for (int i = 0; i < nodes.getLength(); i++) {
            Element el = (Element) nodes.item(i);
            String idxStr = el.getAttribute("idx");
            if (!idxStr.isEmpty() && Integer.parseInt(idxStr) == idx) {
                return Optional.of(new CommentData(el));
            }
        }
        return Optional.empty();
    }

    // ---- Mutation methods ----

    /**
     * Appends a new comment element and returns its data wrapper.
     *
     * @param authorId  the author ID
     * @param idx       the comment index
     * @param text      the comment text
     * @param posX      the X position in cm
     * @param posY      the Y position in cm
     * @param dtStr     the OOXML datetime string
     * @param parentIdx the parent comment index for replies, or {@code null}
     * @return the created comment data
     */
    public CommentData addComment(int authorId, int idx, String text, double posX,
                                  double posY, String dtStr, Integer parentIdx) {
        Element elem = doc.createElementNS(NS_P, "p:cm");
        elem.setAttribute("authorId", String.valueOf(authorId));
        elem.setAttribute("dt", dtStr);
        elem.setAttribute("idx", String.valueOf(idx));
        if (parentIdx != null) {
            elem.setAttribute("parentCmId", String.valueOf(parentIdx));
        }

        Element pos = doc.createElementNS(NS_P, "p:pos");
        pos.setAttribute("x", String.valueOf(CommentData.toPositionUnits(posX)));
        pos.setAttribute("y", String.valueOf(CommentData.toPositionUnits(posY)));
        elem.appendChild(pos);

        Element textElem = doc.createElementNS(NS_P, "p:text");
        textElem.setTextContent(text);
        elem.appendChild(textElem);

        root.appendChild(elem);
        return new CommentData(elem);
    }

    /**
     * Inserts a comment at the given index among existing comments.
     *
     * @param index     the position to insert at
     * @param authorId  the author ID
     * @param idx       the comment index
     * @param text      the comment text
     * @param posX      the X position in cm
     * @param posY      the Y position in cm
     * @param dtStr     the OOXML datetime string
     * @param parentIdx the parent comment index for replies, or {@code null}
     * @return the created comment data
     */
    public CommentData insertComment(int index, int authorId, int idx, String text,
                                     double posX, double posY, String dtStr, Integer parentIdx) {
        Element elem = doc.createElementNS(NS_P, "p:cm");
        elem.setAttribute("authorId", String.valueOf(authorId));
        elem.setAttribute("dt", dtStr);
        elem.setAttribute("idx", String.valueOf(idx));
        if (parentIdx != null) {
            elem.setAttribute("parentCmId", String.valueOf(parentIdx));
        }

        Element pos = doc.createElementNS(NS_P, "p:pos");
        pos.setAttribute("x", String.valueOf(CommentData.toPositionUnits(posX)));
        pos.setAttribute("y", String.valueOf(CommentData.toPositionUnits(posY)));
        elem.appendChild(pos);

        Element textElem = doc.createElementNS(NS_P, "p:text");
        textElem.setTextContent(text);
        elem.appendChild(textElem);

        List<Element> allCm = getCmElements();
        if (index >= allCm.size()) {
            root.appendChild(elem);
        } else {
            root.insertBefore(elem, allCm.get(index));
        }
        return new CommentData(elem);
    }

    /**
     * Removes a comment by author ID and index.
     *
     * @param authorId the author ID
     * @param idx      the comment index
     */
    public void removeComment(int authorId, int idx) {
        NodeList nodes = root.getElementsByTagNameNS(NS_P, "cm");
        for (int i = 0; i < nodes.getLength(); i++) {
            Element el = (Element) nodes.item(i);
            String aid = el.getAttribute("authorId");
            String idxStr = el.getAttribute("idx");
            if (!aid.isEmpty() && Integer.parseInt(aid) == authorId
                    && !idxStr.isEmpty() && Integer.parseInt(idxStr) == idx) {
                root.removeChild(el);
                return;
            }
        }
    }

    /**
     * Removes a specific comment element from the document.
     *
     * @param elem the DOM element to remove
     */
    public void removeCommentElem(Element elem) {
        try {
            root.removeChild(elem);
        } catch (org.w3c.dom.DOMException e) {
            // Element not a child of root — ignore
        }
    }

    /**
     * Removes the comment at the given positional index.
     *
     * @param index the zero-based index
     */
    public void removeCommentsAt(int index) {
        List<Element> allCm = getCmElements();
        if (index >= 0 && index < allCm.size()) {
            root.removeChild(allCm.get(index));
        }
    }

    /**
     * Removes all comments from this part.
     */
    public void clear() {
        List<Element> allCm = getCmElements();
        for (Element el : allCm) {
            root.removeChild(el);
        }
    }

    /**
     * Returns the number of comments in this part.
     *
     * @return the comment count
     */
    public int count() {
        return root.getElementsByTagNameNS(NS_P, "cm").getLength();
    }

    /**
     * Returns whether this part contains no comments.
     *
     * @return {@code true} if empty
     */
    public boolean isEmpty() {
        return count() == 0;
    }

    // ---- Persistence ----

    /**
     * Saves the comments XML back to the package.
     */
    public void save() {
        pkg.serializeXml(partName, doc);
    }

    // ---- Static factory and lifecycle methods ----

    /**
     * Creates a new empty comments part for a slide and registers all relationships.
     *
     * @param pkg           the OPC package
     * @param slidePartName the slide part name (e.g. {@code "ppt/slides/slide1.xml"})
     * @return the created CommentsPart
     */
    public static CommentsPart createForSlide(OpcPackage pkg, String slidePartName) {
        // Find a unique part name
        int num = 1;
        String partName;
        while (true) {
            String candidate = "ppt/comments/slide" + num + ".xml";
            if (!pkg.hasPart(candidate)) {
                partName = candidate;
                break;
            }
            num++;
        }

        // Build minimal XML: <p:cmLst xmlns:p="..."/>
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<p:cmLst xmlns:p=\"" + NS_P + "\"/>";
        pkg.setPartBytes(partName, xml.getBytes(StandardCharsets.UTF_8));

        // Add content type override
        addContentTypeOverride(pkg, partName, CONTENT_TYPE_COMMENTS);

        // Add relationship from slide to comments part
        String relTarget = computeRelativeTarget(slidePartName, partName);
        addSlideRelationship(pkg, slidePartName, REL_TYPE_COMMENTS, relTarget);

        return new CommentsPart(pkg, partName);
    }

    /**
     * Loads the comments part for a slide, if it exists.
     *
     * @param pkg           the OPC package
     * @param slidePartName the slide part name
     * @return the comments part, or empty if no comments exist for this slide
     */
    public static Optional<CommentsPart> loadForSlide(OpcPackage pkg, String slidePartName) {
        String relsUri = getRelsPartName(slidePartName);
        Document relsDoc = pkg.parseXml(relsUri);
        if (relsDoc == null) {
            return Optional.empty();
        }

        Element relsRoot = relsDoc.getDocumentElement();
        NodeList rels = relsRoot.getElementsByTagName("Relationship");
        for (int i = 0; i < rels.getLength(); i++) {
            Element rel = (Element) rels.item(i);
            if (REL_TYPE_COMMENTS.equals(rel.getAttribute("Type"))) {
                String target = rel.getAttribute("Target");
                String resolved = resolveTarget(slidePartName, target);
                if (pkg.hasPart(resolved)) {
                    return Optional.of(new CommentsPart(pkg, resolved));
                }
                return Optional.empty();
            }
        }
        return Optional.empty();
    }

    /**
     * Deletes the comments part for a slide, removing the part, content type, and relationship.
     *
     * @param pkg           the OPC package
     * @param slidePartName the slide part name
     */
    public static void delete(OpcPackage pkg, String slidePartName) {
        String relsUri = getRelsPartName(slidePartName);
        Document relsDoc = pkg.parseXml(relsUri);
        if (relsDoc == null) {
            return;
        }

        Element relsRoot = relsDoc.getDocumentElement();
        NodeList rels = relsRoot.getElementsByTagName("Relationship");
        boolean modified = false;
        // Iterate backwards to safely remove nodes
        for (int i = rels.getLength() - 1; i >= 0; i--) {
            Element rel = (Element) rels.item(i);
            if (REL_TYPE_COMMENTS.equals(rel.getAttribute("Type"))) {
                String target = rel.getAttribute("Target");
                String resolved = resolveTarget(slidePartName, target);

                // Remove the part itself
                pkg.removePart(resolved);

                // Remove content type override
                removeContentTypeOverride(pkg, resolved);

                // Remove the relationship element
                relsRoot.removeChild(rel);
                modified = true;
            }
        }
        if (modified) {
            pkg.serializeXml(relsUri, relsDoc);
        }
    }

    // ---- Path resolution utilities ----

    /**
     * Resolves a relative target path to an absolute part name.
     *
     * @param fromPart the source part path
     * @param target   the relative or absolute target
     * @return the resolved absolute part name
     */
    public static String resolveTarget(String fromPart, String target) {
        if (target.startsWith("/")) {
            return target.substring(1);
        }
        String baseDir = fromPart.contains("/")
                ? fromPart.substring(0, fromPart.lastIndexOf('/'))
                : "";
        String[] parts = (baseDir + "/" + target).split("/");
        var resolved = new ArrayList<String>();
        for (String part : parts) {
            if ("..".equals(part)) {
                if (!resolved.isEmpty()) {
                    resolved.removeLast();
                }
            } else if (!part.isEmpty() && !".".equals(part)) {
                resolved.add(part);
            }
        }
        return String.join("/", resolved);
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
        var downParts = new ArrayList<String>();
        for (int i = commonLen; i < toParts.length; i++) {
            downParts.add(toParts[i]);
        }
        String downPath = String.join("/", downParts);

        var sb = new StringBuilder();
        sb.append("../".repeat(upCount));
        if (!downPath.isEmpty()) {
            sb.append(downPath).append('/');
        }
        sb.append(toFile);
        return sb.toString();
    }

    // ---- Date/time utilities ----

    /**
     * Converts a {@link LocalDateTime} or {@link LocalDate} to an OOXML datetime string.
     *
     * @param dt the date/time value
     * @return the formatted string
     */
    public static String dtToStr(LocalDateTime dt) {
        if (dt == null) {
            return "";
        }
        int ms = dt.getNano() / 1_000_000;
        return dt.format(DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss"))
                + ".%03d".formatted(ms);
    }

    /**
     * Converts a {@link LocalDate} to an OOXML datetime string.
     *
     * @param date the date value
     * @return the formatted string
     */
    public static String dtToStr(LocalDate date) {
        if (date == null) {
            return "";
        }
        return date.format(DateTimeFormatter.ofPattern("yyyy-MM-dd")) + "T00:00:00.000";
    }

    /**
     * Parses an OOXML datetime string to a {@link LocalDateTime}.
     *
     * @param s the datetime string
     * @return the parsed value, or empty if the string is null/blank or unparseable
     */
    public static Optional<LocalDateTime> strToDt(String s) {
        if (s == null || s.isBlank()) {
            return Optional.empty();
        }
        // Try formats: with fractional seconds, without, date-only
        String[] formats = {
                "yyyy-MM-dd'T'HH:mm:ss.SSSSSS",
                "yyyy-MM-dd'T'HH:mm:ss.SSS",
                "yyyy-MM-dd'T'HH:mm:ss",
                "yyyy-MM-dd"
        };
        for (String fmt : formats) {
            try {
                var formatter = DateTimeFormatter.ofPattern(fmt);
                if (fmt.equals("yyyy-MM-dd")) {
                    LocalDate date = LocalDate.parse(s, formatter);
                    return Optional.of(date.atStartOfDay());
                }
                return Optional.of(LocalDateTime.parse(s, formatter));
            } catch (DateTimeParseException e) {
                // Try next format
            }
        }
        return Optional.empty();
    }

    // ---- Private helpers ----

    private void load() {
        doc = pkg.parseXml(partName);
        if (doc == null) {
            throw new IllegalArgumentException("Comments part not found: " + partName);
        }
        root = doc.getDocumentElement();
    }

    private List<Element> getCmElements() {
        var result = new ArrayList<Element>();
        NodeList nodes = root.getElementsByTagNameNS(NS_P, "cm");
        for (int i = 0; i < nodes.getLength(); i++) {
            result.add((Element) nodes.item(i));
        }
        return result;
    }

    private static String getRelsPartName(String partName) {
        int lastSlash = partName.lastIndexOf('/');
        if (lastSlash >= 0) {
            return partName.substring(0, lastSlash + 1) + "_rels/"
                    + partName.substring(lastSlash + 1) + ".rels";
        }
        return "_rels/" + partName + ".rels";
    }

    private static void addSlideRelationship(OpcPackage pkg, String slidePartName,
                                             String type, String target) {
        String relsUri = getRelsPartName(slidePartName);
        Document relsDoc = pkg.parseXml(relsUri);
        Element relsRoot;
        if (relsDoc == null) {
            relsDoc = OpcPackage.newDocument();
            relsRoot = relsDoc.createElementNS(REL_NS, "Relationships");
            relsDoc.appendChild(relsRoot);
        } else {
            relsRoot = relsDoc.getDocumentElement();
            // Check if relationship already exists
            NodeList rels = relsRoot.getElementsByTagName("Relationship");
            for (int i = 0; i < rels.getLength(); i++) {
                Element rel = (Element) rels.item(i);
                if (type.equals(rel.getAttribute("Type"))) {
                    return;
                }
            }
        }

        Element rel = relsDoc.createElementNS(REL_NS, "Relationship");
        rel.setAttribute("Id", "rId_comments");
        rel.setAttribute("Type", type);
        rel.setAttribute("Target", target);
        relsRoot.appendChild(rel);
        pkg.serializeXml(relsUri, relsDoc);
    }

    private static void addContentTypeOverride(OpcPackage pkg, String partName, String contentType) {
        Document ctDoc = pkg.parseXml("[Content_Types].xml");
        if (ctDoc == null) {
            return;
        }
        Element ctRoot = ctDoc.getDocumentElement();
        // Check for existing override
        NodeList overrides = ctRoot.getElementsByTagName("Override");
        String partNameWithSlash = "/" + partName;
        for (int i = 0; i < overrides.getLength(); i++) {
            Element ov = (Element) overrides.item(i);
            if (partNameWithSlash.equals(ov.getAttribute("PartName"))) {
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
        NodeList overrides = ctRoot.getElementsByTagName("Override");
        String partNameWithSlash = "/" + partName;
        for (int i = overrides.getLength() - 1; i >= 0; i--) {
            Element ov = (Element) overrides.item(i);
            if (partNameWithSlash.equals(ov.getAttribute("PartName"))) {
                ctRoot.removeChild(ov);
            }
        }
        pkg.serializeXml("[Content_Types].xml", ctDoc);
    }
}
