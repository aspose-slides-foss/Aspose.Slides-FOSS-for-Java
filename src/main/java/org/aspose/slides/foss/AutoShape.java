package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.Optional;

/**
 * Represents an AutoShape.
 *
 * <p>Extends {@link GeometryShape} with text frame and text box capabilities.</p>
 */
public final class AutoShape extends GeometryShape implements IAutoShape {

    /**
     * Creates an AutoShape backed by the given XML element.
     *
     * @param xmlElement   the {@code <p:sp>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public AutoShape(Element xmlElement, Runnable saveCallback) {
        super(xmlElement, saveCallback);
    }

    /**
     * Creates an AutoShape with no backing element.
     */
    public AutoShape() {
        super();
    }

    /**
     * {@inheritDoc}
     *
     * <p>Checks for custom geometry ({@code <a:custGeom>}) in addition to the
     * base class preset geometry lookup.</p>
     */
    @Override
    public ShapeType getShapeType() {
        if (xmlElement == null) return ShapeType.NOT_DEFINED;
        Element spPr = findChild(xmlElement, NS_P, "spPr");
        if (spPr == null) return ShapeType.NOT_DEFINED;
        Element prstGeom = findChild(spPr, NS_A, "prstGeom");
        if (prstGeom == null) {
            Element custGeom = findChild(spPr, NS_A, "custGeom");
            if (custGeom != null) {
                return ShapeType.CUSTOM;
            }
            return ShapeType.NOT_DEFINED;
        }
        String prst = prstGeom.getAttribute("prst");
        if (prst == null || prst.isEmpty()) return ShapeType.NOT_DEFINED;
        return ShapeType.fromOoxml(prst).orElse(ShapeType.NOT_DEFINED);
    }

    /**
     * {@inheritDoc}
     *
     * <p>Removes any existing custom geometry ({@code <a:custGeom>}) before
     * applying the preset, and clears children of an existing {@code <a:prstGeom>}.</p>
     */
    @Override
    public void setShapeType(ShapeType value) {
        if (xmlElement == null) return;
        if (value == ShapeType.NOT_DEFINED || value == ShapeType.CUSTOM) return;
        Optional<String> prst = value.toOoxml();
        if (prst.isEmpty()) return;

        Element spPr = ensureChild(xmlElement, NS_P, "spPr", "p:spPr");
        Element prstGeom = findChild(spPr, NS_A, "prstGeom");
        if (prstGeom == null) {
            // Remove custGeom if present
            Element custGeom = findChild(spPr, NS_A, "custGeom");
            if (custGeom != null) {
                spPr.removeChild(custGeom);
            }
            prstGeom = spPr.getOwnerDocument().createElementNS(NS_A, "a:prstGeom");
            spPr.appendChild(prstGeom);
        } else {
            // Clear children of existing prstGeom
            while (prstGeom.hasChildNodes()) {
                prstGeom.removeChild(prstGeom.getFirstChild());
            }
        }
        prstGeom.setAttribute("prst", prst.get());
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public ITextFrame getTextFrame() {
        if (xmlElement == null) return null;
        Element txBody = findChild(xmlElement, NS_P, "txBody");
        if (txBody == null) {
            txBody = findChild(xmlElement, NS_A, "txBody");
        }
        if (txBody == null) return null;
        return new TextFrame(txBody, saveCallback);
    }

    @Override
    public boolean isTextBox() {
        if (xmlElement == null) return false;
        return findChild(xmlElement, NS_P, "txBody") != null
                || findChild(xmlElement, NS_A, "txBody") != null;
    }

    @Override
    public IGeometryShape getAsIGeometryShape() {
        return this;
    }

    @Override
    public ITextFrame addTextFrame(String text) {
        if (xmlElement == null) return null;
        Document doc = xmlElement.getOwnerDocument();

        // Remove existing txBody if any
        Element existing = findChild(xmlElement, NS_P, "txBody");
        if (existing != null) {
            xmlElement.removeChild(existing);
        }
        existing = findChild(xmlElement, NS_A, "txBody");
        if (existing != null) {
            xmlElement.removeChild(existing);
        }

        // Mark shape as text box
        Element nvSpPr = findChild(xmlElement, NS_P, "nvSpPr");
        if (nvSpPr != null) {
            Element cNvSpPr = findChild(nvSpPr, NS_P, "cNvSpPr");
            if (cNvSpPr != null) {
                cNvSpPr.setAttribute("txBox", "1");
            }
        }

        // Create new txBody
        Element txBody = doc.createElementNS(NS_P, "p:txBody");
        xmlElement.appendChild(txBody);

        Element bodyPr = doc.createElementNS(NS_A, "a:bodyPr");
        bodyPr.setAttribute("rtlCol", "0");
        bodyPr.setAttribute("anchor", "ctr");
        txBody.appendChild(bodyPr);

        Element lstStyle = doc.createElementNS(NS_A, "a:lstStyle");
        txBody.appendChild(lstStyle);

        // Split on \r\n, \r, or \n to create separate paragraphs
        String[] lines;
        if (text != null && !text.isEmpty()) {
            lines = text.split("\\r\\n|\\r|\\n", -1);
        } else {
            lines = new String[]{""};
        }

        for (String line : lines) {
            Element p = doc.createElementNS(NS_A, "a:p");
            Element pPr = doc.createElementNS(NS_A, "a:pPr");
            pPr.setAttribute("algn", "ctr");
            p.appendChild(pPr);

            Element r = doc.createElementNS(NS_A, "a:r");
            Element t = doc.createElementNS(NS_A, "a:t");
            t.setTextContent(line);
            r.appendChild(t);
            p.appendChild(r);

            txBody.appendChild(p);
        }

        if (saveCallback != null) saveCallback.run();
        return new TextFrame(txBody, saveCallback);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public IFillFormat getFillFormat() {
        if (xmlElement == null) return null;
        Element spPr = findChild(xmlElement, NS_P, "spPr");
        if (spPr == null) {
            spPr = ensureChild(xmlElement, NS_P, "spPr", "p:spPr");
        }
        return withPartContext(new FillFormat(spPr, saveCallback));
    }
}
