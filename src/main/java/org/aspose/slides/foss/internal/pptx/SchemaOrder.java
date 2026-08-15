package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Inserts lazily-created children at the position their complex type's
 * sequence requires.
 *
 * <p>Most of the ECMA-376 complex types used by this library are
 * {@code xsd:sequence}s, not {@code xsd:all}: a consumer is entitled to
 * reject a document whose children appear in a different order, and the
 * Open XML schema validator does. Appending a child that is created on
 * first use therefore produces invalid markup whenever a later sibling
 * happens to have been created first — for example {@code <a:rPr>}
 * created by a formatting call after {@code <a:t>} was created by the
 * constructor, or {@code <a:scene3d>} created after {@code <a:sp3d>}.</p>
 *
 * <p>The tables below list, per parent element, the local names of its
 * children in schema order. Children whose local name is not in the table
 * are ignored when choosing an insertion point, so an element carrying
 * markup this library does not model is never reordered by accident.</p>
 */
public final class SchemaOrder {

    private SchemaOrder() {
        // utility class
    }

    /** {@code CT_RegularTextRun}: {@code rPr?, t}. */
    private static final String[] A_R = {"rPr", "t"};

    /**
     * {@code CT_ShapeProperties} ({@code <p:spPr>} and {@code <a:spPr>}):
     * {@code xfrm?, <geometry>?, <fill>?, ln?, <effect>?, scene3d?, sp3d?, extLst?}.
     */
    private static final String[] SP_PR = {
            "xfrm",
            "custGeom", "prstGeom",
            "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
            "ln",
            "effectLst", "effectDag",
            "scene3d", "sp3d",
            "extLst",
    };

    /** {@code CT_Scene3D}: {@code camera, lightRig, backdrop?, extLst?}. */
    private static final String[] A_SCENE3D = {"camera", "lightRig", "backdrop", "extLst"};

    /** {@code CT_Shape3D}: {@code bevelT?, bevelB?, extrusionClr?, contourClr?, extLst?}. */
    private static final String[] A_SP3D = {"bevelT", "bevelB", "extrusionClr", "contourClr", "extLst"};

    /** {@code CT_TextParagraph}: {@code pPr?, (run choice)*, endParaRPr?}. */
    private static final String[] A_P = {
            "pPr",
            "r", "br", "fld",
            "endParaRPr",
    };

    /** {@code CT_TextBody}: {@code bodyPr, lstStyle?, p+}. */
    private static final String[] TX_BODY = {"bodyPr", "lstStyle", "p"};

    private static final Map<String, List<String>> ORDERS = orders();

    private static Map<String, List<String>> orders() {
        Map<String, List<String>> m = new HashMap<>();
        m.put("r", Arrays.asList(A_R));
        m.put("spPr", Arrays.asList(SP_PR));
        m.put("scene3d", Arrays.asList(A_SCENE3D));
        m.put("sp3d", Arrays.asList(A_SP3D));
        m.put("p", Arrays.asList(A_P));
        m.put("txBody", Arrays.asList(TX_BODY));
        return Collections.unmodifiableMap(m);
    }

    /**
     * Returns whether a child order is known for the given parent element.
     *
     * @param parent the parent element; may be {@code null}
     * @return {@code true} if {@link #insert} will order children of this parent
     */
    public static boolean isKnown(Element parent) {
        return parent != null && ORDERS.containsKey(parent.getLocalName());
    }

    /**
     * Inserts {@code child} into {@code parent} at the position the parent's
     * complex type requires.
     *
     * <p>Falls back to appending when no child order is known for the parent,
     * when the child is not part of that order, or when no existing child
     * must follow it.</p>
     *
     * @param parent the parent element
     * @param child  the child element to insert
     * @return {@code child}, for fluent use
     */
    public static Element insert(Element parent, Element child) {
        List<String> order = ORDERS.get(parent.getLocalName());
        if (order == null) {
            parent.appendChild(child);
            return child;
        }
        int rank = order.indexOf(child.getLocalName());
        if (rank < 0) {
            parent.appendChild(child);
            return child;
        }
        Node before = null;
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node node = children.item(i);
            if (!(node instanceof Element el)) continue;
            int existing = order.indexOf(el.getLocalName());
            if (existing >= 0 && existing > rank) {
                before = node;
                break;
            }
        }
        if (before == null) {
            parent.appendChild(child);
        } else {
            parent.insertBefore(child, before);
        }
        return child;
    }

    /**
     * Moves an element that is already a child of {@code parent} to the
     * position its complex type requires, if it is not there already.
     *
     * @param parent the parent element
     * @param child  an existing child of {@code parent}
     */
    public static void reposition(Element parent, Element child) {
        if (child.getParentNode() != parent) return;
        parent.removeChild(child);
        insert(parent, child);
    }
}
