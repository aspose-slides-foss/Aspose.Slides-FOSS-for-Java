package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.ArrayList;
import java.util.List;

/**
 * Keeps every text body written to the package valid.
 *
 * <p>{@code CT_TextBody} is {@code bodyPr, lstStyle?, p+}: the paragraph is required, so a
 * body written with no {@code <a:p>} is invalid whatever produced it. Several routes lead
 * there — emptying the paragraph collection, removing paragraphs one by one, setting a text
 * frame's text to {@code null} — and each would otherwise need its own guard.</p>
 *
 * <p>Guarding them individually is also the wrong shape: a collection whose {@code clear()}
 * left a paragraph behind would report a count of one after being emptied, and a caller
 * draining it with {@code while (size() > 0) removeAt(0)} would never finish. The
 * requirement is about the file rather than about the object model, so it is met where the
 * file is produced: an empty body gains an empty paragraph on its way out, which is what
 * PowerPoint leaves in a placeholder whose text has been deleted.</p>
 *
 * <p>The repair is applied to a copy, never to the document the object model is reading.
 * The two are the same tree, and mutating it would put the paragraph straight back into a
 * collection that had just been emptied — again, without end, since every mutation
 * serializes.</p>
 *
 * <p>A document whose text bodies all have a paragraph is returned unchanged and is not
 * copied, so this never alters, or costs anything for, a document that was already valid —
 * including one loaded from another producer and saved again untouched.</p>
 */
public final class TextBodies {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static final String NS_P =
            "http://schemas.openxmlformats.org/presentationml/2006/main";

    private TextBodies() {
        // utility class
    }

    /**
     * Returns a document whose every text body carries at least one paragraph.
     *
     * @param doc the document about to be serialized; may be {@code null}
     * @return {@code doc} itself when nothing needs repairing, otherwise a repaired copy
     */
    public static Document withParagraphs(Document doc) {
        if (doc == null || emptyTextBodies(doc).isEmpty()) {
            return doc;
        }
        Document copy = OpcPackage.newDocument();
        copy.appendChild(copy.importNode(doc.getDocumentElement(), true));
        for (Element body : emptyTextBodies(copy)) {
            SchemaOrder.insert(body, copy.createElementNS(NS_A, "a:p"));
        }
        return copy;
    }

    /** Returns the {@code <p:txBody>} and {@code <a:txBody>} elements that have no paragraph. */
    private static List<Element> emptyTextBodies(Document doc) {
        List<Element> empty = new ArrayList<>();
        collectEmpty(doc.getElementsByTagNameNS(NS_P, "txBody"), empty);
        collectEmpty(doc.getElementsByTagNameNS(NS_A, "txBody"), empty);
        return empty;
    }

    private static void collectEmpty(NodeList nodes, List<Element> into) {
        for (int i = 0; i < nodes.getLength(); i++) {
            if (nodes.item(i) instanceof Element body && !hasParagraph(body)) {
                into.add(body);
            }
        }
    }

    /** Reports whether the body has an {@code <a:p>} child. */
    private static boolean hasParagraph(Element body) {
        NodeList children = body.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element child
                    && NS_A.equals(child.getNamespaceURI())
                    && "p".equals(child.getLocalName())) {
                return true;
            }
        }
        return false;
    }
}
