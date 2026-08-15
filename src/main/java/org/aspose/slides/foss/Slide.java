package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.OpcPackage;
import org.aspose.slides.foss.internal.pptx.RelsHelper;
import org.aspose.slides.foss.internal.pptx.SlideReference;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.function.Function;

/**
 * Represents a slide in a presentation.
 */
public final class Slide implements ISlide {

    private static final String NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private final Presentation presentation;
    private final int index;
    private NotesSlideManager notesSlideManager;
    private ShapeCollection shapes;
    private String name = "";
    private boolean hidden;
    private ILayoutSlide layoutSlide;
    private Document slideDoc;
    private String slidePartUri;

    /* Fields populated by initInternal */
    private org.aspose.slides.foss.internal.opc.OpcPackage opcPackage;
    private String partName;
    private SlideReference slideRef;
    private Object slidePart;
    private Function<String, ILayoutSlide> layoutResolver;
    private Optional<ILayoutSlide> layoutSlideCache = Optional.empty();
    private Optional<INotesSlideManager> notesSlideManagerCache = Optional.empty();

    /**
     * Creates a slide with the given index in the given presentation.
     *
     * @param presentation the owning presentation
     * @param index        the zero-based slide index
     */
    Slide(Presentation presentation, int index) {
        this.presentation = presentation;
        this.index = index;
        this.notesSlideManager = new NotesSlideManager(this, presentation.getPackage());
        loadShapes();
        loadHiddenState();
    }

    /**
     * Internal initialization for a slide.
     *
     * <p>Sets up the slide's internal state from the OPC package components.
     * This is used as an alternative initialization path when constructing
     * slides from parsed package data.</p>
     *
     * @param presentation   the parent {@link Presentation} object
     * @param opcPackage     the OPC package containing slide data
     * @param partName       the part name of this slide within the package
     * @param slideRef       the {@link SlideReference} from the presentation part
     * @param slidePart      the parsed slide part data
     * @param layoutResolver optional resolver that maps a part name to an {@link ILayoutSlide},
     *                       or {@code null} if layout resolution is not needed
     */
    public void initInternal(Presentation presentation,
                             org.aspose.slides.foss.internal.opc.OpcPackage opcPackage,
                             String partName,
                             SlideReference slideRef,
                             Object slidePart,
                             Function<String, ILayoutSlide> layoutResolver) {
        this.opcPackage = opcPackage;
        this.partName = partName;
        this.slideRef = slideRef;
        this.slidePart = slidePart;
        this.layoutResolver = layoutResolver;
        this.layoutSlideCache = Optional.empty();
        this.notesSlideManagerCache = Optional.empty();
    }

    private void loadShapes() {
        OpcPackage pkg = presentation.getPackage();
        slidePartUri = "ppt/slides/slide" + (index + 1) + ".xml";
        Document doc = pkg.parseXml(slidePartUri);
        if (doc == null) {
            shapes = new ShapeCollection();
            return;
        }
        this.slideDoc = doc;
        loadNameFromXml();
        Element spTree = findSpTree(doc.getDocumentElement());
        Runnable saveCallback = () -> pkg.serializeXml(slidePartUri, doc);
        shapes = new ShapeCollection(spTree, saveCallback);
        resolveLayoutSlide(pkg);
    }

    private void loadHiddenState() {
        if (slideDoc == null) return;
        Element root = slideDoc.getDocumentElement();
        String showAttr = root.getAttribute("show");
        // In PPTX, show="0" means hidden
        if ("0".equals(showAttr)) {
            hidden = true;
        }
    }

    private Element findSpTree(Element root) {
        NodeList cSldList = root.getElementsByTagNameNS(NS_P, "cSld");
        if (cSldList.getLength() == 0) return null;
        Element cSld = (Element) cSldList.item(0);
        NodeList spTreeList = cSld.getElementsByTagNameNS(NS_P, "spTree");
        if (spTreeList.getLength() == 0) return null;
        return (Element) spTreeList.item(0);
    }

    private void loadNameFromXml() {
        if (slideDoc == null) return;
        Element root = slideDoc.getDocumentElement();
        NodeList cSldList = root.getElementsByTagNameNS(NS_P, "cSld");
        if (cSldList.getLength() > 0) {
            Element cSld = (Element) cSldList.item(0);
            String nameAttr = cSld.getAttribute("name");
            if (nameAttr != null && !nameAttr.isEmpty()) {
                this.name = nameAttr;
            }
        }
    }

    private void persistName() {
        if (slideDoc == null) return;
        Element root = slideDoc.getDocumentElement();
        NodeList cSldList = root.getElementsByTagNameNS(NS_P, "cSld");
        if (cSldList.getLength() > 0) {
            Element cSld = (Element) cSldList.item(0);
            if (name != null && !name.isEmpty()) {
                cSld.setAttribute("name", name);
            } else {
                cSld.removeAttribute("name");
            }
            presentation.getPackage().serializeXml(slidePartUri, slideDoc);
        }
    }

    private void resolveLayoutSlide(OpcPackage pkg) {
        var slideRels = new RelsHelper(pkg, slidePartUri);
        String layoutRelType =
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
        for (var rel : slideRels.getAllRelationships()) {
            if (layoutRelType.equals(rel.type())) {
                String target = rel.target();
                String layoutPartName;
                if (target.startsWith("/")) {
                    layoutPartName = target.replaceFirst("^/+", "");
                } else {
                    // Resolve relative to slide directory (ppt/slides/)
                    layoutPartName = "ppt/slides/" + target;
                    // Normalize: ../slideLayouts/slideLayout1.xml -> ppt/slideLayouts/slideLayout1.xml
                    layoutPartName = normalizePath(layoutPartName);
                }
                LayoutSlide resolved = presentation.resolveLayoutSlide(layoutPartName);
                if (resolved != null) {
                    this.layoutSlide = resolved;
                }
                break;
            }
        }
    }

    private static String normalizePath(String path) {
        String[] parts = path.split("/");
        var result = new ArrayList<String>();
        for (String part : parts) {
            if ("..".equals(part)) {
                if (!result.isEmpty()) {
                    result.remove(result.size() - 1);
                }
            } else if (!".".equals(part) && !part.isEmpty()) {
                result.add(part);
            }
        }
        return String.join("/", result);
    }

    private void loadExistingShapes(Element spTree, Runnable saveCallback) {
        NodeList children = spTree.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el) {
                String localName = el.getLocalName();
                if ("sp".equals(localName)) {
                    AutoShape shape = new AutoShape(el, saveCallback);
                    shape.setParentShapes(shapes.getInternalList());
                    shapes.getInternalList().add(shape);
                } else if ("pic".equals(localName)) {
                    PictureFrame frame = new PictureFrame(el, saveCallback);
                    frame.setParentShapes(shapes.getInternalList());
                    shapes.getInternalList().add(frame);
                } else if ("cxnSp".equals(localName)) {
                    Connector connector = new Connector(el, saveCallback);
                    connector.setParentShapes(shapes.getInternalList());
                    shapes.getInternalList().add(connector);
                } else if ("graphicFrame".equals(localName)) {
                    // Check if this graphic frame contains a table
                    if (containsTable(el)) {
                        Table table = new Table(el, saveCallback);
                        table.setParentShapes(shapes.getInternalList());
                        shapes.getInternalList().add(table);
                    }
                }
            }
        }
    }

    private static boolean containsTable(Element graphicFrame) {
        NodeList graphics = graphicFrame.getElementsByTagNameNS(
                "http://schemas.openxmlformats.org/drawingml/2006/main", "tbl");
        return graphics.getLength() > 0;
    }

    /**
     * Returns the zero-based index of this slide.
     *
     * @return the slide index
     */
    public int getIndex() {
        return index;
    }

    @Override
    public IBaseSlide getSlide() {
        return this;
    }

    @Override
    public IPresentationComponent asIPresentationComponent() {
        return this;
    }

    @Override
    public IPresentation getPresentation() {
        return presentation;
    }

    @Override
    public String getName() {
        return name;
    }

    @Override
    public void setName(String name) {
        this.name = name != null ? name : "";
        persistName();
    }

    @Override
    public int getSlideId() {
        return index;
    }

    @Override
    public IShapeCollection getShapes() {
        return shapes;
    }

    @Override
    public int getSlideNumber() {
        return index + 1;
    }

    @Override
    public void setSlideNumber(int value) {
        // Slide number is derived from the slide's index in the collection.
    }

    @Override
    public boolean isHidden() {
        return hidden;
    }

    @Override
    public void setHidden(boolean value) {
        this.hidden = value;
        persistHiddenState();
    }

    private void persistHiddenState() {
        if (slideDoc == null) return;
        Element root = slideDoc.getDocumentElement();
        if (hidden) {
            root.setAttribute("show", "0");
        } else {
            root.removeAttribute("show");
        }
        presentation.getPackage().serializeXml(slidePartUri, slideDoc);
    }

    @Override
    public ILayoutSlide getLayoutSlide() {
        return layoutSlide;
    }

    @Override
    public void setLayoutSlide(ILayoutSlide value) {
        this.layoutSlide = value;
    }

    @Override
    public INotesSlideManager getNotesSlideManager() {
        return notesSlideManager;
    }

    @Override
    public IComment[] getSlideComments(ICommentAuthor author) {
        List<IComment> result = new ArrayList<>();
        for (CommentAuthor ca : presentation.getCommentAuthorCollection().getInternalList()) {
            if (author != null && ca != author) continue;
            for (Comment c : ca.getCommentCollection().getInternalList()) {
                if (c.getSlide() == this) {
                    result.add(c);
                }
            }
        }
        return result.toArray(new IComment[0]);
    }

    @Override
    public void remove() {
        presentation.getSlideCollection().remove(this);
    }

    /**
     * Returns the OPC part URI for this slide (e.g. {@code "ppt/slides/slide1.xml"}).
     *
     * @return the slide part URI
     */
    String getSlidePartUri() {
        return slidePartUri;
    }

    /**
     * Flushes the in-memory DOM back to the OPC package.
     * Called by {@link Presentation#save} to ensure all DOM modifications
     * are persisted even if the save callback was not invoked.
     */
    void flush() {
        if (slideDoc != null && slidePartUri != null) {
            presentation.getPackage().serializeXml(slidePartUri, slideDoc);
        }
    }
}
