package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.UUID;

/**
 * Reads and writes the modern threaded-comment parts.
 *
 * <p>A reply is not an attribute of the classic {@code <p:cm>} element:
 * {@code CT_Comment} has no place to record one, so a library that models
 * replies has to write the separate parts PowerPoint 2018 introduced —
 * {@code ppt/threadedComments/threadedCommentN.xml}, whose {@code p188:cm}
 * entries carry {@code @id} and {@code @parentId}, and {@code ppt/authors.xml},
 * which gives each author the GUID those entries refer to.</p>
 *
 * <p>The classic {@code ppt/comments/} parts are still written beside these, so
 * that a reader that does not understand threading still shows the comments.</p>
 *
 * <p>Identifiers are derived from the content they name — the author from name
 * and initials, a comment from its slide and its position in that slide's
 * comment list — rather than drawn at random, so that saving the same
 * presentation twice produces the same bytes.</p>
 */
public final class ThreadedCommentsPart {

    /** Namespace of the PowerPoint 2018 threaded-comment elements. */
    public static final String NS_P188 =
            "http://schemas.microsoft.com/office/powerpoint/2018/8/main";

    /** Namespace of the DrawingML elements used inside a comment's text body. */
    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /** Content type of a threaded-comment part. */
    public static final String CONTENT_TYPE = "application/vnd.ms-powerpoint.threadedcomments+xml";

    /** Content type of the author list. */
    public static final String AUTHORS_CONTENT_TYPE = "application/vnd.ms-powerpoint.authors+xml";

    /** Relationship type from a slide to its threaded-comment part. */
    public static final String REL_TYPE =
            "http://schemas.microsoft.com/office/2018/10/relationships/threadedComment";

    /** Relationship type from the presentation to the author list. */
    public static final String AUTHORS_REL_TYPE =
            "http://schemas.microsoft.com/office/2018/10/relationships/authors";

    /** Part name of the author list. */
    public static final String AUTHORS_PART_NAME = "ppt/authors.xml";

    private static final DateTimeFormatter CREATED =
            DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSS");

    private ThreadedCommentsPart() {
        // utility class
    }

    /**
     * One comment as it is written to a threaded-comment part.
     *
     * @param authorKey   the key identifying the author, from {@link #authorGuid}
     * @param created     when the comment was made
     * @param text        the comment text
     * @param posXEmu     the anchor's x offset in EMU; ignored for a reply
     * @param posYEmu     the anchor's y offset in EMU; ignored for a reply
     * @param parentIndex the position in this list of the comment being replied
     *                    to, or {@code -1} for a comment that starts a thread
     */
    public record Entry(String authorKey, LocalDateTime created, String text,
                        long posXEmu, long posYEmu, int parentIndex) {
    }

    /**
     * Returns the part name of the threaded-comment part for a slide.
     *
     * @param slideNumber the slide part number
     * @return the part name
     */
    public static String partName(int slideNumber) {
        return "ppt/threadedComments/threadedComment" + slideNumber + ".xml";
    }

    /**
     * Returns the stable GUID identifying an author.
     *
     * @param name     the author's name
     * @param initials the author's initials
     * @return a GUID in the brace-delimited upper-case form OOXML uses
     */
    public static String authorGuid(String name, String initials) {
        return guid("author:" + name + "|" + initials);
    }

    private static String guid(String seed) {
        return "{" + UUID.nameUUIDFromBytes(seed.getBytes(StandardCharsets.UTF_8))
                .toString().toUpperCase(Locale.ROOT) + "}";
    }

    /**
     * Writes the threaded-comment part for a slide, with its content type and
     * its relationship from the slide.
     *
     * <p>Writes nothing and removes any existing part when no comment on the
     * slide is part of a thread: a deck of unrelated comments does not need the
     * modern part, and shipping an empty one would be noise in every file.</p>
     *
     * @param pkg          the package to write into
     * @param slidePartUri the part name of the slide the comments belong to
     * @param slideNumber  the slide part number
     * @param entries      the slide's comments, in the order the classic part
     *                     writes them
     */
    public static void write(OpcPackage pkg, String slidePartUri, int slideNumber,
                             List<Entry> entries) {
        String partUri = partName(slideNumber);
        boolean anyThread = entries.stream().anyMatch(e -> e.parentIndex() >= 0);
        if (!anyThread) {
            delete(pkg, slidePartUri, slideNumber);
            return;
        }

        List<String> ids = new ArrayList<>(entries.size());
        for (int i = 0; i < entries.size(); i++) {
            ids.add(guid(partUri + "#" + i));
        }

        Document doc = OpcPackage.newDocument();
        Element root = doc.createElementNS(NS_P188, "p188:cmLst");
        root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:a", NS_A);
        doc.appendChild(root);

        for (int i = 0; i < entries.size(); i++) {
            Entry entry = entries.get(i);
            Element cm = doc.createElementNS(NS_P188, "p188:cm");
            cm.setAttribute("id", ids.get(i));
            cm.setAttribute("authorId", entry.authorKey());
            cm.setAttribute("created",
                    (entry.created() != null ? entry.created() : LocalDateTime.now()).format(CREATED));
            if (entry.parentIndex() >= 0 && entry.parentIndex() < ids.size()) {
                cm.setAttribute("parentId", ids.get(entry.parentIndex()));
            } else {
                Element pos = doc.createElementNS(NS_P188, "p188:pos");
                pos.setAttribute("x", String.valueOf(entry.posXEmu()));
                pos.setAttribute("y", String.valueOf(entry.posYEmu()));
                cm.appendChild(pos);
            }
            cm.appendChild(textBody(doc, entry.text()));
            root.appendChild(cm);
        }
        pkg.serializeXml(partUri, doc);

        var contentTypes = new ContentTypesManager(pkg);
        contentTypes.addOverride("/" + partUri, CONTENT_TYPE);
        contentTypes.save();

        var rels = new RelsHelper(pkg, slidePartUri);
        boolean present = rels.getAllRelationships().stream()
                .anyMatch(rel -> REL_TYPE.equals(rel.type()));
        if (!present) {
            rels.addRelationship(REL_TYPE, "../threadedComments/threadedComment"
                    + slideNumber + ".xml");
            rels.save();
        }
    }

    private static Element textBody(Document doc, String text) {
        Element txBody = doc.createElementNS(NS_P188, "p188:txBody");
        txBody.appendChild(doc.createElementNS(NS_A, "a:bodyPr"));
        txBody.appendChild(doc.createElementNS(NS_A, "a:lstStyle"));
        Element p = doc.createElementNS(NS_A, "a:p");
        Element r = doc.createElementNS(NS_A, "a:r");
        Element t = doc.createElementNS(NS_A, "a:t");
        t.setTextContent(text != null ? text : "");
        r.appendChild(t);
        p.appendChild(r);
        txBody.appendChild(p);
        return txBody;
    }

    /**
     * Removes a slide's threaded-comment part together with its relationship and
     * its content-type override.
     *
     * @param pkg          the package
     * @param slidePartUri the part name of the slide
     * @param slideNumber  the slide part number
     */
    public static void delete(OpcPackage pkg, String slidePartUri, int slideNumber) {
        String partUri = partName(slideNumber);
        if (!pkg.hasPart(partUri)) {
            return;
        }
        pkg.removePartCascading(partUri, slidePartUri);
    }

    /**
     * Writes {@code ppt/authors.xml} and relates it from the presentation.
     *
     * @param pkg     the package
     * @param authors the authors, as name/initials pairs in collection order
     */
    public static void writeAuthors(OpcPackage pkg, List<String[]> authors) {
        if (authors.isEmpty()) {
            return;
        }
        Document doc = OpcPackage.newDocument();
        Element root = doc.createElementNS(NS_P188, "p188:authorLst");
        doc.appendChild(root);
        for (String[] author : authors) {
            String name = author[0] != null ? author[0] : "";
            String initials = author[1] != null ? author[1] : "";
            Element el = doc.createElementNS(NS_P188, "p188:author");
            el.setAttribute("id", authorGuid(name, initials));
            el.setAttribute("name", name);
            el.setAttribute("initials", initials);
            el.setAttribute("userId", name);
            el.setAttribute("providerId", "None");
            root.appendChild(el);
        }
        pkg.serializeXml(AUTHORS_PART_NAME, doc);

        var contentTypes = new ContentTypesManager(pkg);
        contentTypes.addOverride("/" + AUTHORS_PART_NAME, AUTHORS_CONTENT_TYPE);
        contentTypes.save();

        var rels = new RelsHelper(pkg, PresentationPart.PART_NAME);
        boolean present = rels.getAllRelationships().stream()
                .anyMatch(rel -> AUTHORS_REL_TYPE.equals(rel.type()));
        if (!present) {
            rels.addRelationship(AUTHORS_REL_TYPE, "authors.xml");
            rels.save();
        }
    }

    /**
     * Reads back which comment each comment on a slide replies to.
     *
     * @param pkg         the package
     * @param slideNumber the slide part number
     * @return for each comment, in the order the part lists them, the index of
     *         the comment it replies to, or {@code -1}; empty when the slide has
     *         no threaded-comment part
     */
    public static List<Integer> readParentIndices(OpcPackage pkg, int slideNumber) {
        Document doc = pkg.parseXml(partName(slideNumber));
        if (doc == null) {
            return List.of();
        }
        NodeList nodes = doc.getElementsByTagNameNS(NS_P188, "cm");
        Map<String, Integer> byId = new HashMap<>();
        for (int i = 0; i < nodes.getLength(); i++) {
            byId.put(((Element) nodes.item(i)).getAttribute("id"), i);
        }
        List<Integer> parents = new ArrayList<>(nodes.getLength());
        for (int i = 0; i < nodes.getLength(); i++) {
            String parentId = ((Element) nodes.item(i)).getAttribute("parentId");
            parents.add(parentId.isEmpty() ? -1 : byId.getOrDefault(parentId, -1));
        }
        return parents;
    }
}
