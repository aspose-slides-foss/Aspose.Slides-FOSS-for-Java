package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.util.Set;

/**
 * Represents fill formatting properties.
 *
 * <p>Wraps an OOXML parent element (e.g. {@code <p:spPr>}, {@code <a:tcPr>})
 * for reading and writing fill format properties.</p>
 */
public class FillFormat extends PVIObject implements IFillFormat, IFillParamSource {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static final Set<String> FILL_LOCAL_NAMES = Set.of(
            "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"
    );

    /** Local names that should appear after fill elements in OOXML element order. */
    private static final Set<String> POST_FILL_LOCAL_NAMES = Set.of(
            "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"
    );

    protected Element parentElement;
    protected Runnable saveCallback;

    /** The package and part this fill is serialized into, when it belongs to one. */
    private org.aspose.slides.foss.internal.pptx.OpcPackage partPackage;
    private String partName;

    /**
     * Records which package part this fill is serialized into, so that a picture fill
     * can declare its image relationship in that part's {@code .rels}.
     *
     * @param opcPackage the package, or {@code null} if unattached
     * @param partName   the part name, or {@code null} if unattached
     */
    void setPartContext(org.aspose.slides.foss.internal.pptx.OpcPackage opcPackage, String partName) {
        this.partPackage = opcPackage;
        this.partName = partName;
    }

    /**
     * Creates a new FillFormat backed by the given parent XML element.
     *
     * @param parentElement the parent XML element
     * @param saveCallback  callback invoked after mutations; may be {@code null}
     */
    public FillFormat(Element parentElement, Runnable saveCallback) {
        this.parentElement = parentElement;
        this.saveCallback = saveCallback;
    }

    /**
     * Creates an uninitialized FillFormat. Call {@link #initInternal} before use.
     */
    public FillFormat() {
    }

    /**
     * Internal initialization.
     *
     * @param parentElement any XML element that contains fill children
     *                      (e.g. {@code <p:spPr>}, {@code <p:bgPr>}, {@code <a:tcPr>}, {@code <a:rPr>})
     * @param saveCallback  callback invoked after mutations to persist changes; may be {@code null}
     * @param parentSlide   the parent slide object
     */
    public void initInternal(Element parentElement, Runnable saveCallback, IBaseSlide parentSlide) {
        this.parentElement = parentElement;
        this.saveCallback = saveCallback;
        this.parentSlide = parentSlide;
    }

    // --- internal helpers ---

    protected Element findFillElement() {
        NodeList children = parentElement.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && NS_A.equals(el.getNamespaceURI())
                    && FILL_LOCAL_NAMES.contains(el.getLocalName())) {
                return el;
            }
        }
        return null;
    }

    protected void removeFillElements() {
        NodeList children = parentElement.getChildNodes();
        for (int i = children.getLength() - 1; i >= 0; i--) {
            Node child = children.item(i);
            if (child instanceof Element el
                    && NS_A.equals(el.getNamespaceURI())
                    && FILL_LOCAL_NAMES.contains(el.getLocalName())) {
                parentElement.removeChild(el);
            }
        }
    }

    protected Element insertFillElement(String localName) {
        Document doc = parentElement.getOwnerDocument();
        Element el = doc.createElementNS(NS_A, "a:" + localName);
        // Insert before post-fill elements to maintain OOXML element order
        Node insertBefore = null;
        NodeList children = parentElement.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element child
                    && NS_A.equals(child.getNamespaceURI())
                    && POST_FILL_LOCAL_NAMES.contains(child.getLocalName())) {
                insertBefore = child;
                break;
            }
        }
        if (insertBefore != null) {
            parentElement.insertBefore(el, insertBefore);
        } else {
            parentElement.appendChild(el);
        }
        return el;
    }

    private Element getOrCreateFill(String localName) {
        Element el = findFillElement();
        if (el != null && localName.equals(el.getLocalName())) {
            return el;
        }
        removeFillElements();
        el = insertFillElement(localName);
        initFillDefaults(el, localName);
        return el;
    }

    private void initFillDefaults(Element el, String localName) {
        Document doc = el.getOwnerDocument();
        if ("gradFill".equals(localName)) {
            Element lin = doc.createElementNS(NS_A, "a:lin");
            lin.setAttribute("ang", "0");
            lin.setAttribute("scaled", "1");
            el.appendChild(lin);
        } else if ("blipFill".equals(localName)) {
            el.appendChild(doc.createElementNS(NS_A, "a:blip"));
            Element stretch = doc.createElementNS(NS_A, "a:stretch");
            stretch.appendChild(doc.createElementNS(NS_A, "a:fillRect"));
            el.appendChild(stretch);
        }
    }

    protected void save() {
        if (saveCallback != null) saveCallback.run();
    }

    private static Element findChild(Element parent, String nsUri, String localName) {
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && nsUri.equals(el.getNamespaceURI())
                    && localName.equals(el.getLocalName())) {
                return el;
            }
        }
        return null;
    }

    // --- public API ---

    @Override
    public FillType getFillType() {
        Element el = findFillElement();
        if (el == null) return FillType.NOT_DEFINED;
        return switch (el.getLocalName()) {
            case "noFill" -> FillType.NO_FILL;
            case "solidFill" -> FillType.SOLID;
            case "gradFill" -> FillType.GRADIENT;
            case "pattFill" -> FillType.PATTERN;
            case "blipFill" -> FillType.PICTURE;
            case "grpFill" -> FillType.GROUP;
            default -> FillType.NOT_DEFINED;
        };
    }

    @Override
    public void setFillType(FillType value) {
        String localName = switch (value) {
            case NO_FILL -> "noFill";
            case SOLID -> "solidFill";
            case GRADIENT -> "gradFill";
            case PATTERN -> "pattFill";
            case PICTURE -> "blipFill";
            case GROUP -> "grpFill";
            default -> null;
        };
        // If the fill type already matches, preserve existing element
        Element existing = findFillElement();
        if (existing != null && localName != null && localName.equals(existing.getLocalName())) {
            return;
        }
        removeFillElements();
        if (localName != null) {
            Element el = insertFillElement(localName);
            initFillDefaults(el, localName);
        }
        save();
    }

    @Override
    public IColorFormat getSolidFillColor() {
        Element solidFill = getOrCreateFill("solidFill");
        return new ColorFormat(solidFill, saveCallback);
    }

    @Override
    public IGradientFormat getGradientFormat() {
        Element el = findFillElement();
        if (el == null || !"gradFill".equals(el.getLocalName())) {
            el = getOrCreateFill("gradFill");
        }
        return new GradientFormat(el, saveCallback);
    }

    @Override
    public IPatternFormat getPatternFormat() {
        Element el = findFillElement();
        if (el == null || !"pattFill".equals(el.getLocalName())) {
            el = getOrCreateFill("pattFill");
        }
        return new PatternFormat(el, saveCallback);
    }

    @Override
    public IPictureFillFormat getPictureFillFormat() {
        Element el = findFillElement();
        if (el == null || !"blipFill".equals(el.getLocalName())) {
            el = getOrCreateFill("blipFill");
        }
        var pictureFill = new PictureFillFormat(el, saveCallback);
        pictureFill.setPartContext(partPackage, partName);
        return pictureFill;
    }

    @Override
    public NullableBool getRotateWithShape() {
        Element el = findFillElement();
        if (el == null) return NullableBool.NOT_DEFINED;
        String val = el.getAttribute("rotWithShape");
        if (val == null || val.isEmpty()) return NullableBool.NOT_DEFINED;
        return "1".equals(val) ? NullableBool.TRUE : NullableBool.FALSE;
    }

    @Override
    public void setRotateWithShape(NullableBool value) {
        Element el = findFillElement();
        if (el == null) return;
        if (value == NullableBool.NOT_DEFINED) {
            el.removeAttribute("rotWithShape");
        } else {
            el.setAttribute("rotWithShape", value == NullableBool.TRUE ? "1" : "0");
        }
        save();
    }
}
