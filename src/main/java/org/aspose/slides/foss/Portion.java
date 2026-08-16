package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.SchemaOrder;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.ParserConfigurationException;
import org.aspose.slides.foss.internal.xml.SecureXml;

/**
 * Represents a text portion (run) within a paragraph.
 *
 * <p>Wraps an OOXML {@code <a:r>} element containing text content
 * and run-level formatting properties.</p>
 */
public final class Portion implements IPortion {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private Element runElement;
    private Element pElement;
    private Element txBodyElement;
    private Object slidePart;
    private Runnable saveCallback;
    private IBaseSlide parentSlide;

    /**
     * Creates a new empty Portion with a detached {@code <a:r>} element.
     */
    public Portion() {
        try {
            Document doc = SecureXml.documentBuilderFactory()
                    .newDocumentBuilder()
                    .newDocument();
            this.runElement = doc.createElementNS(NS_A, "a:r");
            doc.appendChild(this.runElement);
            Element tEl = doc.createElementNS(NS_A, "a:t");
            this.runElement.appendChild(tEl);
        } catch (ParserConfigurationException e) {
            throw new IllegalStateException("Failed to create XML document", e);
        }
    }

    /**
     * Creates a new Portion with the given text.
     *
     * @param text the initial text
     */
    public Portion(String text) {
        this();
        setText(text);
    }

    /**
     * Creates a Portion backed by an existing XML run element.
     *
     * @param runElement   the {@code <a:r>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public Portion(Element runElement, Runnable saveCallback) {
        this.runElement = runElement;
        this.saveCallback = saveCallback;
    }

    private void save() {
        if (saveCallback != null) saveCallback.run();
    }

    private Element findChild(String localName) {
        NodeList children = runElement.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el && localName.equals(el.getLocalName())) {
                return el;
            }
        }
        return null;
    }

    private Element getOrCreateChild(String localName) {
        Element el = findChild(localName);
        if (el == null) {
            Document doc = runElement.getOwnerDocument();
            el = doc.createElementNS(NS_A, "a:" + localName);
            SchemaOrder.insert(runElement, el);
        }
        return el;
    }

    @Override
    public String getText() {
        Element tEl = findChild("t");
        return tEl != null ? tEl.getTextContent() : "";
    }

    @Override
    public void setText(String value) {
        Element tEl = getOrCreateChild("t");
        tEl.setTextContent(value != null ? value : "");
        save();
    }

    @Override
    public IPortionFormat getPortionFormat() {
        Element rPr = findChild("rPr");
        if (rPr == null) {
            rPr = getOrCreateChild("rPr");
        }
        return new PortionFormat(rPr, this::save);
    }

    @Override
    public ISlideComponent asISlideComponent() {
        return this;
    }

    @Override
    public IPresentationComponent asIPresentationComponent() {
        return this;
    }

    @Override
    public IBaseSlide getSlide() {
        return parentSlide;
    }

    /**
     * Sets the parent slide for this portion.
     *
     * @param slide the parent slide, or {@code null} to clear
     */
    public void setSlide(IBaseSlide slide) {
        this.parentSlide = slide;
    }

    @Override
    public IPresentation getPresentation() {
        return null;
    }

    /**
     * Re-initialises this portion with a (possibly adopted) run element
     * and a new save callback. Used by {@link PortionCollection} when
     * inserting a portion into an XML-backed paragraph.
     *
     * @param runElement   the {@code <a:r>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    void initInternal(Element runElement, Runnable saveCallback) {
        this.runElement = runElement;
        this.saveCallback = saveCallback;
    }

    /**
     * Fully initialises this portion with its XML context and slide hierarchy.
     *
     * <p>Stores the run element together with its parent paragraph element,
     * text-body element, owning slide part, and parent slide reference.
     * Returns {@code this} to allow fluent usage.</p>
     *
     * @param rElement      the {@code <a:r>} XML element
     * @param pElement      the {@code <a:p>} paragraph element containing the run
     * @param txBodyElement the {@code <a:txBody>} element containing the paragraph
     * @param slidePart     the slide part that owns this portion's XML
     * @param parentSlide   the parent slide
     * @return this portion, for fluent chaining
     */
    Portion initInternal(Element rElement, Element pElement, Element txBodyElement,
                         Object slidePart, IBaseSlide parentSlide) {
        this.runElement = rElement;
        this.pElement = pElement;
        this.txBodyElement = txBodyElement;
        this.slidePart = slidePart;
        this.parentSlide = parentSlide;
        return this;
    }

    /**
     * Returns the backing XML element.
     *
     * @return the run element
     */
    Element getRunElement() {
        return runElement;
    }
}
