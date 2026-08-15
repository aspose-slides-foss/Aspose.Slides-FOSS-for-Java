package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

/**
 * Represents a text paragraph.
 *
 * <p>Wraps an OOXML {@code <a:p>} element containing text runs and
 * paragraph-level formatting properties.</p>
 */
public final class Paragraph implements IParagraph, ISlideComponent, IPresentationComponent {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private Element pElement;
    private Element txbodyElement;
    private Object slidePart;
    private Runnable saveCallback;
    private PortionCollection portions;
    private ParagraphFormat paragraphFormat;
    private IBaseSlide parentSlide;

    /**
     * Creates a new empty Paragraph with a detached {@code <a:p>} element.
     */
    public Paragraph() {
        try {
            Document doc = DocumentBuilderFactory.newInstance()
                    .newDocumentBuilder()
                    .newDocument();
            this.pElement = doc.createElementNS(NS_A, "a:p");
            doc.appendChild(this.pElement);
        } catch (ParserConfigurationException e) {
            throw new IllegalStateException("Failed to create XML document", e);
        }
        this.portions = new PortionCollection(this.pElement, null);
    }

    /**
     * Creates a Paragraph backed by an existing XML element.
     *
     * @param pElement     the {@code <a:p>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public Paragraph(Element pElement, Runnable saveCallback) {
        this.pElement = pElement;
        this.saveCallback = saveCallback;
        this.portions = new PortionCollection(pElement, this::save);
        loadPortions();
    }

    /**
     * Initialises this paragraph from existing OOXML elements.
     *
     * <p>Sets the backing XML elements and parent references, reloads the
     * portion collection from the supplied {@code <a:p>} element, and
     * resolves any pending blip image references when both a slide part
     * and parent slide are provided.</p>
     *
     * @param pElement      the {@code <a:p>} XML element
     * @param txbodyElement the {@code <a:txBody>} XML element containing this paragraph
     * @param slidePart     the OPC slide part that owns the paragraph, or {@code null}
     * @param parentSlide   the parent slide, or {@code null}
     * @return this paragraph, for method chaining
     */
    public Paragraph initInternal(Element pElement, Element txbodyElement,
                                  Object slidePart, IBaseSlide parentSlide) {
        return initInternal(pElement, txbodyElement, slidePart, parentSlide, this.saveCallback);
    }

    /**
     * Initialises this paragraph from existing OOXML elements, replacing its
     * save callback.
     *
     * @param pElement      the {@code <a:p>} XML element
     * @param txbodyElement the {@code <a:txBody>} XML element containing this paragraph
     * @param slidePart     the OPC slide part that owns the paragraph, or {@code null}
     * @param parentSlide   the parent slide, or {@code null}
     * @param saveCallback  callback invoked after mutations, or {@code null}
     * @return this paragraph, for method chaining
     */
    public Paragraph initInternal(Element pElement, Element txbodyElement,
                                  Object slidePart, IBaseSlide parentSlide,
                                  Runnable saveCallback) {
        this.pElement = pElement;
        this.txbodyElement = txbodyElement;
        this.slidePart = slidePart;
        this.parentSlide = parentSlide;
        this.saveCallback = saveCallback;
        this.portions = new PortionCollection(pElement, this::save);
        loadPortions();
        return this;
    }

    private void loadPortions() {
        NodeList children = pElement.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && NS_A.equals(el.getNamespaceURI())
                    && "r".equals(el.getLocalName())) {
                portions.add(new Portion(el, this::save));
            }
        }
    }

    private void save() {
        if (saveCallback != null) saveCallback.run();
    }

    @Override
    public IPortionCollection getPortions() {
        return portions;
    }

    @Override
    public IParagraphFormat getParagraphFormat() {
        if (paragraphFormat == null) {
            paragraphFormat = new ParagraphFormat(pElement, this::save);
        }
        // Passed on every call, not only on creation: the paragraph learns which part it
        // belongs to when its text frame is attached to a shape, which can be later.
        paragraphFormat.setPartContext(slidePart instanceof PartContext ctx ? ctx : null);
        return paragraphFormat;
    }

    @Override
    public String getText() {
        var sb = new StringBuilder();
        for (int i = 0; i < portions.size(); i++) {
            sb.append(portions.get(i).getText());
        }
        return sb.toString();
    }

    @Override
    public void setText(String value) {
        portions.clear();
        if (value != null && !value.isEmpty()) {
            var portion = new Portion(value);
            portions.add(portion);
        }
        save();
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
     * Sets the parent slide for this paragraph.
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
     * Returns the backing XML element.
     *
     * @return the paragraph element
     */
    Element getPElement() {
        return pElement;
    }
}
