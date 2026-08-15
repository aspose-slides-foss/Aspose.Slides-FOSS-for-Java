package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.BaseCollection;
import org.aspose.slides.foss.internal.pptx.SchemaOrder;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Represents a collection of paragraphs.
 *
 * <p>When backed by an OOXML {@code <a:txBody>} element, paragraphs are
 * resolved dynamically from the XML on each access. Otherwise, the
 * collection operates on an in-memory list.</p>
 */
public final class ParagraphCollection extends BaseCollection<IParagraph>
        implements IParagraphCollection, ISlideComponent, IPresentationComponent {

    private static final String NS_A =
            "http://schemas.openxmlformats.org/drawingml/2006/main";

    private final List<IParagraph> paragraphs = new ArrayList<>();

    private Element txbodyElement;
    private Object slidePart;
    private IBaseSlide parentSlide;
    private Runnable saveCallback;

    /**
     * Creates an empty ParagraphCollection.
     */
    public ParagraphCollection() {
    }

    /**
     * Creates a ParagraphCollection with the given initial paragraphs.
     *
     * @param paragraphs the initial paragraphs
     */
    public ParagraphCollection(List<IParagraph> paragraphs) {
        if (paragraphs != null) {
            this.paragraphs.addAll(paragraphs);
        }
    }

    /**
     * Initialises this collection from an OOXML text-body element.
     *
     * <p>After this call, {@link #getParagraphs()} dynamically resolves
     * {@code <a:p>} children from the supplied element.</p>
     *
     * @param txbodyElement the {@code <a:txBody>} XML element
     * @param slidePart     the OPC slide part that owns the paragraphs, or {@code null}
     * @param parentSlide   the parent slide, or {@code null}
     * @return this collection, for method chaining
     */
    public ParagraphCollection initInternal(Element txbodyElement,
                                            Object slidePart,
                                            IBaseSlide parentSlide) {
        return initInternal(txbodyElement, slidePart, parentSlide, null);
    }

    /**
     * Initialises this collection from an OOXML text-body element, with a
     * callback to invoke after every mutation.
     *
     * <p>After this call the collection is <em>attached</em>: reads resolve
     * {@code <a:p>} children of the supplied element, and {@link #add},
     * {@link #insert}, {@link #removeAt}, {@link #remove} and {@link #clear}
     * change those children rather than a private list.</p>
     *
     * @param txbodyElement the {@code <a:txBody>} XML element
     * @param slidePart     the OPC slide part that owns the paragraphs, or {@code null}
     * @param parentSlide   the parent slide, or {@code null}
     * @param saveCallback  callback invoked after mutations, or {@code null}
     * @return this collection, for method chaining
     */
    public ParagraphCollection initInternal(Element txbodyElement,
                                            Object slidePart,
                                            IBaseSlide parentSlide,
                                            Runnable saveCallback) {
        this.txbodyElement = txbodyElement;
        this.slidePart = slidePart;
        this.parentSlide = parentSlide;
        this.saveCallback = saveCallback;
        return this;
    }

    private void save() {
        if (saveCallback != null) saveCallback.run();
    }

    /** Returns the {@code <a:p>} children of the backing text body, in document order. */
    private List<Element> paragraphElements() {
        List<Element> result = new ArrayList<>();
        if (txbodyElement == null) return result;
        NodeList children = txbodyElement.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && NS_A.equals(el.getNamespaceURI())
                    && "p".equals(el.getLocalName())) {
                result.add(el);
            }
        }
        return result;
    }

    /**
     * Puts the paragraph's {@code <a:p>} element into the backing text body and
     * rebinds the paragraph object to it, so that later changes to the object
     * reach the same document the file is written from.
     *
     * @param value  the paragraph to attach
     * @param before the element to insert before, or {@code null} to append
     */
    private void attach(IParagraph value, Element before) {
        if (!(value instanceof Paragraph paragraph)) {
            throw new IllegalArgumentException(
                    "A paragraph added to a text frame must be an instance of Paragraph, was "
                            + (value == null ? "null" : value.getClass().getName()));
        }
        Element pElement = paragraph.getPElement();
        if (pElement.getParentNode() != txbodyElement) {
            if (pElement.getOwnerDocument() != txbodyElement.getOwnerDocument()) {
                pElement = (Element) txbodyElement.getOwnerDocument().importNode(pElement, true);
            }
            if (before != null) {
                txbodyElement.insertBefore(pElement, before);
            } else {
                SchemaOrder.insert(txbodyElement, pElement);
            }
        }
        paragraph.initInternal(pElement, txbodyElement, slidePart, parentSlide, saveCallback);
        save();
    }

    /** Returns the index of the paragraph backed by the given element, or -1. */
    private int indexOfElement(Element pElement) {
        if (pElement == null) return -1;
        List<Element> elements = paragraphElements();
        for (int i = 0; i < elements.size(); i++) {
            if (elements.get(i) == pElement) return i;
        }
        return -1;
    }

    /**
     * Builds and returns the list of paragraphs.
     *
     * <p>When the collection is backed by an XML element, paragraphs are
     * resolved from the {@code <a:p>} children of the text-body element
     * on every call. Otherwise, returns the in-memory list.</p>
     *
     * @return a list of paragraphs
     */
    public List<IParagraph> getParagraphs() {
        if (txbodyElement == null) {
            return List.copyOf(paragraphs);
        }
        List<IParagraph> result = new ArrayList<>();
        for (Element el : paragraphElements()) {
            Paragraph para = new Paragraph();
            para.initInternal(el, txbodyElement, slidePart, parentSlide, saveCallback);
            result.add(para);
        }
        return result;
    }

    /**
     * Sets the parent slide for this collection.
     *
     * @param parentSlide the parent slide
     */
    void setParentSlide(IBaseSlide parentSlide) {
        this.parentSlide = parentSlide;
    }

    @Override
    public IParagraph get(int index) {
        return getParagraphs().get(index);
    }

    @Override
    public int size() {
        return getParagraphs().size();
    }

    @Override
    public int count() {
        return getParagraphs().size();
    }

    @Override
    public boolean isReadOnly() {
        return false;
    }

    @Override
    public ISlideComponent asISlideComponent() {
        return this;
    }

    @Override
    public Iterable<IParagraph> asIEnumerable() {
        return getParagraphs();
    }

    @Override
    public IPresentationComponent asIPresentationComponent() {
        return this;
    }

    @Override
    public IBaseSlide getSlide() {
        return parentSlide;
    }

    @Override
    public IPresentation getPresentation() {
        if (parentSlide != null) {
            return parentSlide.getPresentation();
        }
        return null;
    }

    @Override
    public void add(IParagraph value) {
        if (txbodyElement == null) {
            paragraphs.add(value);
            return;
        }
        attach(value, null);
    }

    @Override
    public void insert(int index, IParagraph value) {
        if (txbodyElement == null) {
            if (index >= paragraphs.size()) {
                paragraphs.add(value);
            } else {
                paragraphs.add(index, value);
            }
            return;
        }
        List<Element> elements = paragraphElements();
        attach(value, index >= 0 && index < elements.size() ? elements.get(index) : null);
    }

    @Override
    public int indexOf(IParagraph item) {
        if (txbodyElement == null) {
            return paragraphs.indexOf(item);
        }
        return item instanceof Paragraph p ? indexOfElement(p.getPElement()) : -1;
    }

    @Override
    public boolean contains(IParagraph item) {
        return indexOf(item) >= 0;
    }

    @Override
    public void clear() {
        if (txbodyElement != null) {
            for (Element el : paragraphElements()) {
                txbodyElement.removeChild(el);
            }
            save();
        }
        paragraphs.clear();
    }

    @Override
    public void removeAt(int index) {
        if (txbodyElement == null) {
            if (index >= 0 && index < paragraphs.size()) {
                paragraphs.remove(index);
            }
            return;
        }
        List<Element> elements = paragraphElements();
        if (index < 0 || index >= elements.size()) return;
        txbodyElement.removeChild(elements.get(index));
        save();
    }

    @Override
    public boolean remove(IParagraph item) {
        if (txbodyElement == null) {
            return paragraphs.remove(item);
        }
        int index = indexOf(item);
        if (index < 0) return false;
        removeAt(index);
        return true;
    }

    @Override
    public Iterator<IParagraph> iterator() {
        return getParagraphs().iterator();
    }

    /**
     * Returns the internal list (for framework use).
     *
     * @return the internal list
     */
    List<IParagraph> getInternalList() {
        return paragraphs;
    }
}
