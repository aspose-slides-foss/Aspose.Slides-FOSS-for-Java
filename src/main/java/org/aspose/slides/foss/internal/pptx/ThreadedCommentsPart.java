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
 * that a reader that does not understand threading still shows the comments —
 * and the thread is repeated on them as a {@code p15:threadingInfo} extension,
 * because that is where a reader that renders comments looks for it. A package
 * that carries the modern part alone is valid and opens, and the reply is shown
 * as a second unrelated comment: the discussion is silently flattened. Both
 * places are written, and {@link #setClassicThreadParent} is the second.</p>
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

    /** Namespace of the PowerPoint 2012 elements that thread the classic comment list. */
    public static final String NS_P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main";

    /**
     * The {@code p:ext/@uri} that identifies a {@code <p15:threadingInfo>} extension.
     *
     * <p>An extension is recognised by this URI, not by the element inside it, so it has to
     * be written exactly.</p>
     */
    public static final String THREADING_INFO_URI = "{C676402C-5697-4E1C-873F-D02D1690AC5C}";

    /** Namespace of the classic PresentationML comment elements. */
    private static final String NS_P =
            "http://schemas.openxmlformats.org/presentationml/2006/main";

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
     * Returns the threaded-comment part a slide is related to.
     *
     * <p>Found through the slide's relationships, never by its number: a part is whatever the
     * relationship names, and a cloned slide's part carries whatever number was free.</p>
     *
     * @param pkg          the package
     * @param slidePartUri the part name of the slide
     * @return the related part name, or {@code null} if the slide has none
     */
    public static String relatedPartName(OpcPackage pkg, String slidePartUri) {
        for (var rel : new RelsHelper(pkg, slidePartUri).getAllRelationships()) {
            if (REL_TYPE.equals(rel.type()) && !"External".equals(rel.targetMode())) {
                return SlidePart.resolveTargetStatic(slidePartUri, rel.target());
            }
        }
        return null;
    }

    private static String newPartName(OpcPackage pkg, int slideNumber) {
        String preferred = partName(slideNumber);
        for (int number = 1; pkg.hasPart(preferred); number++) {
            preferred = partName(number);
        }
        return preferred;
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
        boolean anyThread = entries.stream().anyMatch(e -> e.parentIndex() >= 0);
        if (!anyThread) {
            delete(pkg, slidePartUri);
            return;
        }
        String related = relatedPartName(pkg, slidePartUri);
        String partUri = related != null ? related : newPartName(pkg, slideNumber);

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

        if (related == null) {
            var rels = new RelsHelper(pkg, slidePartUri);
            rels.addRelationship(REL_TYPE, SlidePart.computeRelativeTarget(slidePartUri, partUri));
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
     * Identifies the comment a reply answers, as the classic list names it.
     *
     * <p>By author and index rather than by one id: a comment's {@code idx} is
     * only unique within its author.</p>
     *
     * @param authorId the {@code p:cmAuthor/@id} of the parent's author
     * @param idx      the parent's {@code p:cm/@idx}
     */
    public record ParentRef(int authorId, int idx) {
    }

    /**
     * Records on a classic {@code <p:cm>} the comment it replies to, or clears it.
     *
     * <p>{@code CT_Comment} declares no attribute for a parent; the extension list
     * is where markup outside the standard belongs, and {@code p:extLst} is the
     * last element of the sequence, so it is appended.</p>
     *
     * @param cm     the {@code <p:cm>} element
     * @param parent the comment being replied to, or {@code null} for none
     */
    public static void setClassicThreadParent(Element cm, ParentRef parent) {
        removeThreadingExtension(cm);
        if (parent == null) {
            return;
        }
        Document doc = cm.getOwnerDocument();
        Element extLst = firstChild(cm, NS_P, "extLst");
        if (extLst == null) {
            extLst = doc.createElementNS(NS_P, "p:extLst");
            cm.appendChild(extLst);
        }
        Element ext = doc.createElementNS(NS_P, "p:ext");
        ext.setAttribute("uri", THREADING_INFO_URI);
        extLst.appendChild(ext);
        Element info = doc.createElementNS(NS_P15, "p15:threadingInfo");
        info.setAttribute("timeZoneBias", "0");
        ext.appendChild(info);
        Element parentCm = doc.createElementNS(NS_P15, "p15:parentCm");
        parentCm.setAttribute("authorId", String.valueOf(parent.authorId()));
        parentCm.setAttribute("idx", String.valueOf(parent.idx()));
        info.appendChild(parentCm);
    }

    /**
     * Reads back the comment a classic {@code <p:cm>} replies to.
     *
     * @param cm the {@code <p:cm>} element
     * @return the parent reference, or {@code null} if the comment is not a reply
     */
    public static ParentRef classicThreadParent(Element cm) {
        Element ext = threadingExtension(cm);
        if (ext == null) {
            return null;
        }
        Element info = firstChild(ext, NS_P15, "threadingInfo");
        Element parentCm = info == null ? null : firstChild(info, NS_P15, "parentCm");
        if (parentCm == null) {
            return null;
        }
        try {
            return new ParentRef(Integer.parseInt(parentCm.getAttribute("authorId")),
                    Integer.parseInt(parentCm.getAttribute("idx")));
        } catch (NumberFormatException e) {
            return null;
        }
    }

    /** Removes the threading extension, and the extension list if it is left empty. */
    private static void removeThreadingExtension(Element cm) {
        Element ext = threadingExtension(cm);
        if (ext == null) {
            return;
        }
        Element extLst = (Element) ext.getParentNode();
        extLst.removeChild(ext);
        if (firstChild(extLst, NS_P, "ext") == null) {
            extLst.getParentNode().removeChild(extLst);
        }
    }

    /** Returns the {@code <p:ext>} carrying the threading extension, or {@code null}. */
    private static Element threadingExtension(Element cm) {
        Element extLst = firstChild(cm, NS_P, "extLst");
        if (extLst == null) {
            return null;
        }
        NodeList children = extLst.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && NS_P.equals(el.getNamespaceURI())
                    && "ext".equals(el.getLocalName())
                    && THREADING_INFO_URI.equals(el.getAttribute("uri"))) {
                return el;
            }
        }
        return null;
    }

    /** Returns the first child element with the given namespace and local name, or null. */
    private static Element firstChild(Element parent, String namespace, String localName) {
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && namespace.equals(el.getNamespaceURI())
                    && localName.equals(el.getLocalName())) {
                return el;
            }
        }
        return null;
    }

    /**
     * Removes a slide's threaded-comment part together with its relationship and
     * its content-type override.
     *
     * @param pkg          the package
     * @param slidePartUri the part name of the slide
     */
    public static void delete(OpcPackage pkg, String slidePartUri) {
        String partUri = relatedPartName(pkg, slidePartUri);
        if (partUri == null) {
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
     * @param pkg          the package
     * @param slidePartUri the part name of the slide
     * @return for each comment, in the order the part lists them, the index of
     *         the comment it replies to, or {@code -1}; empty when the slide has
     *         no threaded-comment part
     */
    public static List<Integer> readParentIndices(OpcPackage pkg, String slidePartUri) {
        String partUri = relatedPartName(pkg, slidePartUri);
        Document doc = partUri == null ? null : pkg.parseXml(partUri);
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
