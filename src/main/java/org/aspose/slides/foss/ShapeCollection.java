package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.util.*;

/**
 * Represents a collection of shapes on a slide.
 *
 * <p>Shapes are lazily loaded from the underlying {@code <p:spTree>} XML element
 * and cached until the collection is mutated. An identity map ensures that the
 * same XML element always yields the same Java shape wrapper.</p>
 */
public final class ShapeCollection implements IShapeCollection, Iterable<IShape> {

    private static final String NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final double EMU_PER_POINT = 12700.0;

    /**
     * Default adjustment values (name, val) for connector preset types.
     * Connectors with no entries use an empty avLst.
     */
    private static final Map<String, List<int[]>> CONNECTOR_DEFAULT_ADJUSTMENTS;

    static {
        var map = new LinkedHashMap<String, List<int[]>>();
        map.put("straightConnector1", List.of());
        map.put("bentConnector2", List.of());
        map.put("bentConnector3", List.of(new int[]{/* adj1 */ 50000}));
        map.put("bentConnector4", List.of(new int[]{/* adj1 */ 50000}, new int[]{/* adj2 */ 50000}));
        map.put("bentConnector5", List.of(new int[]{/* adj1 */ 50000}, new int[]{/* adj2 */ 50000}, new int[]{/* adj3 */ 50000}));
        map.put("curvedConnector2", List.of());
        map.put("curvedConnector3", List.of(new int[]{/* adj1 */ 50000}));
        map.put("curvedConnector4", List.of(new int[]{/* adj1 */ 50000}, new int[]{/* adj2 */ 50000}));
        map.put("curvedConnector5", List.of(new int[]{/* adj1 */ 50000}, new int[]{/* adj2 */ 50000}, new int[]{/* adj3 */ 50000}));
        CONNECTOR_DEFAULT_ADJUSTMENTS = Collections.unmodifiableMap(map);
    }

    /** Connector adjustment names indexed by position. */
    private static final String[] ADJ_NAMES = {"adj1", "adj2", "adj3"};

    private Element spTree;
    private Runnable saveCallback;
    private IGroupShape parentGroup;

    /** The package and part these shapes live in, when they belong to one. */
    private org.aspose.slides.foss.internal.pptx.OpcPackage partPackage;
    private String partName;

    /** Lazily-loaded shapes cache; {@code null} means not yet loaded. */
    private List<IShape> shapesCache;

    /** Identity map: XML element &rarr; shape wrapper (preserves object identity). */
    private final IdentityHashMap<Element, IShape> elementToShape = new IdentityHashMap<>();

    /**
     * Creates an empty ShapeCollection.
     */
    public ShapeCollection() {
    }

    /**
     * Creates a ShapeCollection backed by the given shape tree element.
     *
     * @param spTree       the {@code <p:spTree>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public ShapeCollection(Element spTree, Runnable saveCallback) {
        this.spTree = spTree;
        this.saveCallback = saveCallback;
    }

    // ── Two-phase initialization ────────────────────────────────────────

    /**
     * Internal initialization with slide context.
     *
     * @param spTree       the {@code <p:spTree>} XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public void initInternal(Element spTree, Runnable saveCallback) {
        this.spTree = spTree;
        this.saveCallback = saveCallback;
        this.shapesCache = null;
        this.elementToShape.clear();
    }

    /**
     * Records which package part these shapes are serialized into.
     *
     * <p>A shape that embeds an image has to declare a relationship in its own part's
     * {@code .rels}, which means it has to know which part that is.</p>
     *
     * @param opcPackage the package the part belongs to
     * @param partName   the part these shapes live in, e.g. {@code "ppt/slides/slide1.xml"}
     */
    void setPartContext(org.aspose.slides.foss.internal.pptx.OpcPackage opcPackage, String partName) {
        this.partPackage = opcPackage;
        this.partName = partName;
    }

    /**
     * @return the package these shapes are serialized into, or {@code null} if unattached
     */
    org.aspose.slides.foss.internal.pptx.OpcPackage getPartPackage() {
        return partPackage;
    }

    /**
     * @return the part these shapes are serialized into, or {@code null} if unattached
     */
    String getPartName() {
        return partName;
    }

    // ── XML access ──────────────────────────────────────────────────────

    /**
     * Gets the {@code <p:spTree>} element from the slide XML.
     *
     * @return the spTree element, or {@code null}
     */
    public Element getSpTree() {
        return spTree;
    }

    // ── Lazy shape loading ──────────────────────────────────────────────

    /**
     * Loads all shapes from the XML, using the cache when available.
     *
     * <p>Iterates the children of {@code <p:spTree>}, skipping group-shape
     * property elements ({@code nvGrpSpPr}, {@code grpSpPr}), and wraps each
     * recognised shape element ({@code sp}, {@code pic}, {@code graphicFrame},
     * {@code grpSp}, {@code cxnSp}) into the appropriate Java shape object.</p>
     *
     * @return an unmodifiable list of shapes
     */
    public List<IShape> loadShapes() {
        if (shapesCache != null) {
            return shapesCache;
        }

        var shapes = new ArrayList<IShape>();
        if (spTree == null) {
            return shapes;
        }

        NodeList children = spTree.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node node = children.item(i);
            if (!(node instanceof Element elem)) {
                continue;
            }
            String localName = elem.getLocalName();
            if (localName == null) {
                continue;
            }

            // Skip group shape property elements
            if ("nvGrpSpPr".equals(localName) || "grpSpPr".equals(localName)) {
                continue;
            }

            // Recognised shape elements
            if ("sp".equals(localName) || "pic".equals(localName)
                    || "graphicFrame".equals(localName) || "grpSp".equals(localName)
                    || "cxnSp".equals(localName)) {
                IShape shape = elementToShape.get(elem);
                if (shape == null) {
                    shape = createShapeFromElement(elem);
                    if (shape != null) {
                        elementToShape.put(elem, shape);
                    }
                }
                if (shape != null) {
                    shapes.add(shape);
                }
            }
        }

        shapesCache = shapes;
        return shapes;
    }

    /**
     * Creates the appropriate shape wrapper for a given XML element.
     */
    private IShape createShapeFromElement(Element elem) {
        String localName = elem.getLocalName();
        return switch (localName) {
            case "sp" -> {
                var s = new AutoShape(elem, this::save);
                s.setParentShapes(this);
                yield s;
            }
            case "pic" -> {
                var f = new PictureFrame(elem, this::save);
                f.setParentShapes(this);
                yield f;
            }
            case "cxnSp" -> {
                var c = new Connector(elem, this::save);
                c.setParentShapes(this);
                yield c;
            }
            case "graphicFrame" -> {
                if (containsTable(elem)) {
                    var t = new Table(elem, this::save);
                    t.setParentShapes(this);
                    yield t;
                }
                yield null;
            }
            default -> null;
        };
    }

    private static boolean containsTable(Element graphicFrame) {
        NodeList tables = graphicFrame.getElementsByTagNameNS(NS_A, "tbl");
        return tables.getLength() > 0;
    }

    /**
     * Invalidates the shapes cache, forcing a reload on next access.
     */
    public void invalidateCache() {
        shapesCache = null;
    }

    /**
     * Persists changes to the slide part via the save callback.
     */
    public void save() {
        if (saveCallback != null) {
            saveCallback.run();
        }
    }

    // ── Shape ID allocation ─────────────────────────────────────────────

    /**
     * Finds the next available shape ID by scanning all {@code id} attributes
     * in the spTree.
     *
     * @return the next available ID (always &ge; 2)
     */
    public int nextShapeId() {
        if (spTree == null) {
            return 2;
        }
        int maxId = 1;
        maxId = scanMaxId(spTree, maxId);
        return maxId + 1;
    }

    private int scanMaxId(Element element, int currentMax) {
        String idVal = element.getAttribute("id");
        if (idVal != null && !idVal.isEmpty()) {
            try {
                currentMax = Math.max(currentMax, Integer.parseInt(idVal));
            } catch (NumberFormatException ignored) {
                // Non-numeric value; use default
            }
        }
        NodeList children = element.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element child) {
                currentMax = scanMaxId(child, currentMax);
            }
        }
        return currentMax;
    }

    // ── XML builders ────────────────────────────────────────────────────

    /**
     * Builds XML for a new AutoShape ({@code <p:sp>}) element.
     *
     * @param shapeId            the shape ID
     * @param name               the shape name
     * @param shapeType          the shape type
     * @param x                  x-coordinate in points
     * @param y                  y-coordinate in points
     * @param width              width in points
     * @param height             height in points
     * @param createFromTemplate whether to add theme-based style
     * @return the new {@code <p:sp>} element
     */
    public Element buildAutoShapeXml(int shapeId, String name, ShapeType shapeType,
                                     double x, double y, double width, double height,
                                     boolean createFromTemplate) {
        Document doc = spTree.getOwnerDocument();
        String prst = shapeType.toOoxml().orElse("rect");

        String xEmu = String.valueOf(Math.round(x * EMU_PER_POINT));
        String yEmu = String.valueOf(Math.round(y * EMU_PER_POINT));
        String wEmu = String.valueOf(Math.round(width * EMU_PER_POINT));
        String hEmu = String.valueOf(Math.round(height * EMU_PER_POINT));

        Element sp = doc.createElementNS(NS_P, "p:sp");

        // nvSpPr
        Element nvSpPr = doc.createElementNS(NS_P, "p:nvSpPr");
        Element cNvPr = doc.createElementNS(NS_P, "p:cNvPr");
        cNvPr.setAttribute("id", String.valueOf(shapeId));
        cNvPr.setAttribute("name", name);
        nvSpPr.appendChild(cNvPr);
        nvSpPr.appendChild(doc.createElementNS(NS_P, "p:cNvSpPr"));
        nvSpPr.appendChild(doc.createElementNS(NS_P, "p:nvPr"));
        sp.appendChild(nvSpPr);

        // spPr
        Element spPr = doc.createElementNS(NS_P, "p:spPr");
        Element xfrm = doc.createElementNS(NS_A, "a:xfrm");
        Element off = doc.createElementNS(NS_A, "a:off");
        off.setAttribute("x", xEmu);
        off.setAttribute("y", yEmu);
        xfrm.appendChild(off);
        Element ext = doc.createElementNS(NS_A, "a:ext");
        ext.setAttribute("cx", wEmu);
        ext.setAttribute("cy", hEmu);
        xfrm.appendChild(ext);
        spPr.appendChild(xfrm);
        Element prstGeom = doc.createElementNS(NS_A, "a:prstGeom");
        prstGeom.setAttribute("prst", prst);
        prstGeom.appendChild(doc.createElementNS(NS_A, "a:avLst"));
        spPr.appendChild(prstGeom);
        sp.appendChild(spPr);

        if (createFromTemplate) {
            // p:style for theme-based appearance (must come before txBody)
            Element style = doc.createElementNS(NS_P, "p:style");

            Element lnRef = doc.createElementNS(NS_A, "a:lnRef");
            lnRef.setAttribute("idx", "2");
            Element lnRefClr = doc.createElementNS(NS_A, "a:schemeClr");
            lnRefClr.setAttribute("val", "accent1");
            Element shade = doc.createElementNS(NS_A, "a:shade");
            shade.setAttribute("val", "50000");
            lnRefClr.appendChild(shade);
            lnRef.appendChild(lnRefClr);
            style.appendChild(lnRef);

            Element fillRef = doc.createElementNS(NS_A, "a:fillRef");
            fillRef.setAttribute("idx", "1");
            Element fillRefClr = doc.createElementNS(NS_A, "a:schemeClr");
            fillRefClr.setAttribute("val", "accent1");
            fillRef.appendChild(fillRefClr);
            style.appendChild(fillRef);

            Element effectRef = doc.createElementNS(NS_A, "a:effectRef");
            effectRef.setAttribute("idx", "0");
            Element effectRefClr = doc.createElementNS(NS_A, "a:schemeClr");
            effectRefClr.setAttribute("val", "accent1");
            effectRef.appendChild(effectRefClr);
            style.appendChild(effectRef);

            Element fontRef = doc.createElementNS(NS_A, "a:fontRef");
            fontRef.setAttribute("idx", "minor");
            Element fontRefClr = doc.createElementNS(NS_A, "a:schemeClr");
            fontRefClr.setAttribute("val", "lt1");
            fontRef.appendChild(fontRefClr);
            style.appendChild(fontRef);

            sp.appendChild(style);
        }

        // txBody (must come after p:style per OOXML element ordering)
        Element txBody = doc.createElementNS(NS_P, "p:txBody");
        Element bodyPr = doc.createElementNS(NS_A, "a:bodyPr");
        bodyPr.setAttribute("rtlCol", "0");
        bodyPr.setAttribute("anchor", "ctr");
        txBody.appendChild(bodyPr);
        txBody.appendChild(doc.createElementNS(NS_A, "a:lstStyle"));
        Element aP = doc.createElementNS(NS_A, "a:p");
        aP.appendChild(doc.createElementNS(NS_A, "a:endParaRPr"));
        txBody.appendChild(aP);
        sp.appendChild(txBody);

        return sp;
    }

    /**
     * Builds XML for a new connector ({@code <p:cxnSp>}) element.
     *
     * @param shapeId            the shape ID
     * @param name               the shape name
     * @param shapeType          the connector type
     * @param x                  x-coordinate in points
     * @param y                  y-coordinate in points
     * @param width              width in points
     * @param height             height in points
     * @param createFromTemplate whether to add theme-based style
     * @return the new {@code <p:cxnSp>} element
     */
    public Element buildConnectorXml(int shapeId, String name, ShapeType shapeType,
                                     double x, double y, double width, double height,
                                     boolean createFromTemplate) {
        Document doc = spTree.getOwnerDocument();
        String prst = shapeType.toOoxml().orElse("bentConnector3");

        String xEmu = String.valueOf(Math.round(x * EMU_PER_POINT));
        String yEmu = String.valueOf(Math.round(y * EMU_PER_POINT));
        String wEmu = String.valueOf(Math.round(width * EMU_PER_POINT));
        String hEmu = String.valueOf(Math.round(height * EMU_PER_POINT));

        Element cxnSp = doc.createElementNS(NS_P, "p:cxnSp");

        // nvCxnSpPr
        Element nvCxnSpPr = doc.createElementNS(NS_P, "p:nvCxnSpPr");
        Element cNvPr = doc.createElementNS(NS_P, "p:cNvPr");
        cNvPr.setAttribute("id", String.valueOf(shapeId));
        cNvPr.setAttribute("name", name);
        nvCxnSpPr.appendChild(cNvPr);
        nvCxnSpPr.appendChild(doc.createElementNS(NS_P, "p:cNvCxnSpPr"));
        nvCxnSpPr.appendChild(doc.createElementNS(NS_P, "p:nvPr"));
        cxnSp.appendChild(nvCxnSpPr);

        // spPr
        Element spPr = doc.createElementNS(NS_P, "p:spPr");
        Element xfrm = doc.createElementNS(NS_A, "a:xfrm");
        Element off = doc.createElementNS(NS_A, "a:off");
        off.setAttribute("x", xEmu);
        off.setAttribute("y", yEmu);
        xfrm.appendChild(off);
        Element ext = doc.createElementNS(NS_A, "a:ext");
        ext.setAttribute("cx", wEmu);
        ext.setAttribute("cy", hEmu);
        xfrm.appendChild(ext);
        spPr.appendChild(xfrm);
        Element prstGeom = doc.createElementNS(NS_A, "a:prstGeom");
        prstGeom.setAttribute("prst", prst);
        Element avLst = doc.createElementNS(NS_A, "a:avLst");
        // Add default adjustment values for this connector type
        List<int[]> defaults = CONNECTOR_DEFAULT_ADJUSTMENTS.getOrDefault(prst, List.of());
        for (int idx = 0; idx < defaults.size(); idx++) {
            Element gd = doc.createElementNS(NS_A, "a:gd");
            gd.setAttribute("name", ADJ_NAMES[idx]);
            gd.setAttribute("fmla", "val " + defaults.get(idx)[0]);
            avLst.appendChild(gd);
        }
        prstGeom.appendChild(avLst);
        spPr.appendChild(prstGeom);
        cxnSp.appendChild(spPr);

        if (createFromTemplate) {
            Element style = doc.createElementNS(NS_P, "p:style");

            Element lnRef = doc.createElementNS(NS_A, "a:lnRef");
            lnRef.setAttribute("idx", "1");
            Element lnRefClr = doc.createElementNS(NS_A, "a:schemeClr");
            lnRefClr.setAttribute("val", "accent1");
            lnRef.appendChild(lnRefClr);
            style.appendChild(lnRef);

            Element fillRef = doc.createElementNS(NS_A, "a:fillRef");
            fillRef.setAttribute("idx", "0");
            Element fillRefClr = doc.createElementNS(NS_A, "a:schemeClr");
            fillRefClr.setAttribute("val", "accent1");
            fillRef.appendChild(fillRefClr);
            style.appendChild(fillRef);

            Element effectRef = doc.createElementNS(NS_A, "a:effectRef");
            effectRef.setAttribute("idx", "0");
            Element effectRefClr = doc.createElementNS(NS_A, "a:schemeClr");
            effectRefClr.setAttribute("val", "accent1");
            effectRef.appendChild(effectRefClr);
            style.appendChild(effectRef);

            Element fontRef = doc.createElementNS(NS_A, "a:fontRef");
            fontRef.setAttribute("idx", "minor");
            Element fontRefClr = doc.createElementNS(NS_A, "a:schemeClr");
            fontRefClr.setAttribute("val", "tx1");
            fontRef.appendChild(fontRefClr);
            style.appendChild(fontRef);

            cxnSp.appendChild(style);
        }

        return cxnSp;
    }

    // ── Core shape addition implementations ─────────────────────────────

    /**
     * Core implementation for adding or inserting an AutoShape.
     *
     * @param index              insertion index, or {@code null} to append
     * @param shapeType          the shape type
     * @param x                  x-coordinate in points
     * @param y                  y-coordinate in points
     * @param width              width in points
     * @param height             height in points
     * @param createFromTemplate whether to add theme-based style
     * @return the created AutoShape
     */
    public IAutoShape addAutoShapeImpl(Integer index, ShapeType shapeType, double x, double y,
                                       double width, double height, boolean createFromTemplate) {
        if (spTree == null) {
            throw new IllegalStateException("Cannot add shape: slide has no shape tree");
        }

        int shapeId = nextShapeId();
        String shapeName = formatShapeTypeName(shapeType) + " " + shapeId;

        Element spXml = buildAutoShapeXml(shapeId, shapeName, shapeType, x, y, width, height, createFromTemplate);

        insertElementIntoSpTree(index, spXml);

        invalidateCache();
        save();

        AutoShape autoShape = new AutoShape(spXml, this::save);
        autoShape.setParentShapes(this);
        elementToShape.put(spXml, autoShape);
        return autoShape;
    }

    /**
     * Core implementation for adding or inserting a connector.
     *
     * @param index              insertion index, or {@code null} to append
     * @param shapeType          the connector type
     * @param x                  x-coordinate in points
     * @param y                  y-coordinate in points
     * @param width              width in points
     * @param height             height in points
     * @param createFromTemplate whether to add theme-based style
     * @return the created Connector
     */
    public IConnector addConnectorImpl(Integer index, ShapeType shapeType, double x, double y,
                                       double width, double height, boolean createFromTemplate) {
        if (spTree == null) {
            throw new IllegalStateException("Cannot add shape: slide has no shape tree");
        }

        int shapeId = nextShapeId();
        String connectorName = "Connector " + shapeId;

        Element cxnXml = buildConnectorXml(shapeId, connectorName, shapeType, x, y, width, height, createFromTemplate);

        insertElementIntoSpTree(index, cxnXml);

        invalidateCache();
        save();

        Connector connector = new Connector(cxnXml, this::save);
        connector.setParentShapes(this);
        elementToShape.put(cxnXml, connector);
        return connector;
    }

    /**
     * Core implementation for adding or inserting a picture frame.
     *
     * @param index     insertion index, or {@code null} to append
     * @param shapeType the shape type for the picture frame
     * @param x         x-coordinate in points
     * @param y         y-coordinate in points
     * @param width     width in points
     * @param height    height in points
     * @param image     the image to embed
     * @return the created PictureFrame
     */
    public IPictureFrame addPictureFrameImpl(Integer index, ShapeType shapeType, double x, double y,
                                             double width, double height, IPPImage image) {
        if (spTree == null) {
            throw new IllegalStateException("Cannot add shape: slide has no shape tree");
        }

        Document doc = spTree.getOwnerDocument();
        int shapeId = nextShapeId();
        String picName = "Picture " + shapeId;

        String prst = shapeType.toOoxml().orElse("rect");
        String xEmu = String.valueOf(Math.round(x * EMU_PER_POINT));
        String yEmu = String.valueOf(Math.round(y * EMU_PER_POINT));
        String wEmu = String.valueOf(Math.round(width * EMU_PER_POINT));
        String hEmu = String.valueOf(Math.round(height * EMU_PER_POINT));

        Element pic = doc.createElementNS(NS_P, "p:pic");

        // nvPicPr
        Element nvPicPr = doc.createElementNS(NS_P, "p:nvPicPr");
        Element cNvPr = doc.createElementNS(NS_P, "p:cNvPr");
        cNvPr.setAttribute("id", String.valueOf(shapeId));
        cNvPr.setAttribute("name", picName);
        nvPicPr.appendChild(cNvPr);
        Element cNvPicPr = doc.createElementNS(NS_P, "p:cNvPicPr");
        Element picLocks = doc.createElementNS(NS_A, "a:picLocks");
        picLocks.setAttribute("noChangeAspect", "1");
        cNvPicPr.appendChild(picLocks);
        nvPicPr.appendChild(cNvPicPr);
        nvPicPr.appendChild(doc.createElementNS(NS_P, "p:nvPr"));
        pic.appendChild(nvPicPr);

        // blipFill. The r:embed is written once the image is embedded below, since it must
        // name a relationship that this part's .rels actually declares.
        Element blipFill = doc.createElementNS(NS_P, "p:blipFill");
        Element blip = doc.createElementNS(NS_A, "a:blip");
        blipFill.appendChild(blip);
        Element stretch = doc.createElementNS(NS_A, "a:stretch");
        stretch.appendChild(doc.createElementNS(NS_A, "a:fillRect"));
        blipFill.appendChild(stretch);
        pic.appendChild(blipFill);

        // spPr
        Element spPr = doc.createElementNS(NS_P, "p:spPr");
        Element xfrm = doc.createElementNS(NS_A, "a:xfrm");
        Element off = doc.createElementNS(NS_A, "a:off");
        off.setAttribute("x", xEmu);
        off.setAttribute("y", yEmu);
        xfrm.appendChild(off);
        Element ext = doc.createElementNS(NS_A, "a:ext");
        ext.setAttribute("cx", wEmu);
        ext.setAttribute("cy", hEmu);
        xfrm.appendChild(ext);
        spPr.appendChild(xfrm);
        Element prstGeom = doc.createElementNS(NS_A, "a:prstGeom");
        prstGeom.setAttribute("prst", prst);
        prstGeom.appendChild(doc.createElementNS(NS_A, "a:avLst"));
        spPr.appendChild(prstGeom);
        pic.appendChild(spPr);

        insertElementIntoSpTree(index, pic);

        invalidateCache();
        save();

        PictureFrame frame = new PictureFrame(pic, this::save);
        frame.setParentShapes(this);
        elementToShape.put(pic, frame);

        // Point the blip at the image through a relationship declared in this part's .rels
        if (image != null) {
            if (partPackage == null || partName == null) {
                throw new IllegalStateException(
                        "Cannot embed an image: these shapes are not attached to a package part");
            }
            Picture.setBlipImage(blip, partPackage, partName, image);
        }

        return frame;
    }

    /**
     * Core implementation for adding or inserting a table.
     *
     * @param index      insertion index, or {@code null} to append
     * @param x          x-coordinate in points
     * @param y          y-coordinate in points
     * @param colWidths  column widths in points
     * @param rowHeights row heights in points
     * @return the created Table
     */
    public ITable addTableImpl(Integer index, double x, double y,
                               double[] colWidths, double[] rowHeights) {
        if (spTree == null) {
            throw new IllegalStateException("Cannot add shape: slide has no shape tree");
        }

        Document doc = spTree.getOwnerDocument();
        int shapeId = nextShapeId();
        String tableName = "Table " + shapeId;

        double totalW = 0;
        for (double w : colWidths) totalW += w;
        double totalH = 0;
        for (double h : rowHeights) totalH += h;

        String xEmu = String.valueOf(Math.round(x * EMU_PER_POINT));
        String yEmu = String.valueOf(Math.round(y * EMU_PER_POINT));
        String wEmu = String.valueOf(Math.round(totalW * EMU_PER_POINT));
        String hEmu = String.valueOf(Math.round(totalH * EMU_PER_POINT));

        int numCols = colWidths.length;

        // Build the graphicFrame XML
        Element gf = doc.createElementNS(NS_P, "p:graphicFrame");

        // nvGraphicFramePr
        Element nvGfPr = doc.createElementNS(NS_P, "p:nvGraphicFramePr");
        Element cNvPr = doc.createElementNS(NS_P, "p:cNvPr");
        cNvPr.setAttribute("id", String.valueOf(shapeId));
        cNvPr.setAttribute("name", tableName);
        nvGfPr.appendChild(cNvPr);
        Element cNvGfPr = doc.createElementNS(NS_P, "p:cNvGraphicFramePr");
        Element gfLocks = doc.createElementNS(NS_A, "a:graphicFrameLocks");
        gfLocks.setAttribute("noGrp", "1");
        cNvGfPr.appendChild(gfLocks);
        nvGfPr.appendChild(cNvGfPr);
        nvGfPr.appendChild(doc.createElementNS(NS_P, "p:nvPr"));
        gf.appendChild(nvGfPr);

        // p:xfrm
        Element xfrmEl = doc.createElementNS(NS_P, "p:xfrm");
        Element off = doc.createElementNS(NS_A, "a:off");
        off.setAttribute("x", xEmu);
        off.setAttribute("y", yEmu);
        xfrmEl.appendChild(off);
        Element ext = doc.createElementNS(NS_A, "a:ext");
        ext.setAttribute("cx", wEmu);
        ext.setAttribute("cy", hEmu);
        xfrmEl.appendChild(ext);
        gf.appendChild(xfrmEl);

        // a:graphic > a:graphicData > a:tbl
        Element graphic = doc.createElementNS(NS_A, "a:graphic");
        Element graphicData = doc.createElementNS(NS_A, "a:graphicData");
        graphicData.setAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/table");

        Element tbl = doc.createElementNS(NS_A, "a:tbl");

        // a:tblPr with default style
        Element tblPr = doc.createElementNS(NS_A, "a:tblPr");
        tblPr.setAttribute("firstRow", "1");
        tblPr.setAttribute("bandRow", "1");
        Element tblStyleId = doc.createElementNS(NS_A, "a:tableStyleId");
        tblStyleId.setTextContent("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
        tblPr.appendChild(tblStyleId);
        tbl.appendChild(tblPr);

        // a:tblGrid
        Element tblGrid = doc.createElementNS(NS_A, "a:tblGrid");
        for (double cw : colWidths) {
            Element gridCol = doc.createElementNS(NS_A, "a:gridCol");
            gridCol.setAttribute("w", String.valueOf(Math.round(cw * EMU_PER_POINT)));
            tblGrid.appendChild(gridCol);
        }
        tbl.appendChild(tblGrid);

        // Rows
        for (double rh : rowHeights) {
            Element tr = doc.createElementNS(NS_A, "a:tr");
            tr.setAttribute("h", String.valueOf(Math.round(rh * EMU_PER_POINT)));
            for (int c = 0; c < numCols; c++) {
                Element tc = doc.createElementNS(NS_A, "a:tc");
                Element txBody = doc.createElementNS(NS_A, "a:txBody");
                txBody.appendChild(doc.createElementNS(NS_A, "a:bodyPr"));
                txBody.appendChild(doc.createElementNS(NS_A, "a:lstStyle"));
                txBody.appendChild(doc.createElementNS(NS_A, "a:p"));
                tc.appendChild(txBody);
                tc.appendChild(doc.createElementNS(NS_A, "a:tcPr"));
                tr.appendChild(tc);
            }
            tbl.appendChild(tr);
        }

        graphicData.appendChild(tbl);
        graphic.appendChild(graphicData);
        gf.appendChild(graphic);

        insertElementIntoSpTree(index, gf);

        invalidateCache();
        save();

        Table table = new Table(gf, this::save);
        table.setParentShapes(this);
        elementToShape.put(gf, table);
        return table;
    }

    // ── Reordering ──────────────────────────────────────────────────────

    /**
     * Reorders a single shape to a new position within the spTree XML.
     *
     * @param newIndex the target shape index (0-based, skipping property elements)
     * @param shape    the shape to reorder
     */
    public void reorderSingle(int newIndex, IShape shape) {
        if (spTree == null) {
            return;
        }

        // Find the shape's XML element
        Element shapeElem = findXmlElement(shape);
        if (shapeElem == null || shapeElem.getParentNode() != spTree) {
            return;
        }

        // Remove the shape from its current position
        spTree.removeChild(shapeElem);

        // Find the correct insertion point (skip nvGrpSpPr and grpSpPr)
        int insertPosition = 0;
        int currentShapeIndex = -1;
        NodeList children = spTree.getChildNodes();
        List<Node> childList = new ArrayList<>();
        for (int i = 0; i < children.getLength(); i++) {
            childList.add(children.item(i));
        }

        boolean inserted = false;
        for (Node node : childList) {
            if (!(node instanceof Element elem)) {
                continue;
            }
            String localName = elem.getLocalName();
            if ("nvGrpSpPr".equals(localName) || "grpSpPr".equals(localName)) {
                continue;
            }
            currentShapeIndex++;
            if (currentShapeIndex >= newIndex) {
                spTree.insertBefore(shapeElem, elem);
                inserted = true;
                break;
            }
        }

        if (!inserted) {
            // Append at the end
            spTree.appendChild(shapeElem);
        }
    }

    /**
     * Attempts to find the XML element backing a shape.
     */
    private Element findXmlElement(IShape shape) {
        for (Map.Entry<Element, IShape> entry : elementToShape.entrySet()) {
            if (entry.getValue() == shape) {
                return entry.getKey();
            }
        }
        // Fallback: try reflection-free access via GeometryShape field
        if (shape instanceof GeometryShape gs) {
            return gs.xmlElement;
        }
        return null;
    }

    // ── Helper: insert element at shape index within spTree ─────────────

    /**
     * Inserts an element into the spTree at the given shape index,
     * or appends if index is {@code null}.
     */
    private void insertElementIntoSpTree(Integer index, Element element) {
        if (index == null) {
            spTree.appendChild(element);
        } else {
            int shapeCount = 0;
            NodeList children = spTree.getChildNodes();
            Element insertBefore = null;
            for (int i = 0; i < children.getLength(); i++) {
                if (!(children.item(i) instanceof Element elem)) {
                    continue;
                }
                String localName = elem.getLocalName();
                if ("nvGrpSpPr".equals(localName) || "grpSpPr".equals(localName)) {
                    continue;
                }
                if (shapeCount == index) {
                    insertBefore = elem;
                    break;
                }
                shapeCount++;
            }
            if (insertBefore != null) {
                spTree.insertBefore(element, insertBefore);
            } else {
                spTree.appendChild(element);
            }
        }
    }

    /**
     * Formats a shape type name for use as a shape name prefix.
     * Converts underscores to spaces and title-cases each word.
     */
    private static String formatShapeTypeName(ShapeType shapeType) {
        String raw = shapeType.name().replace('_', ' ').toLowerCase(java.util.Locale.ROOT);
        StringBuilder sb = new StringBuilder();
        boolean capitalizeNext = true;
        for (char c : raw.toCharArray()) {
            if (c == ' ') {
                sb.append(' ');
                capitalizeNext = true;
            } else if (capitalizeNext) {
                sb.append(Character.toUpperCase(c));
                capitalizeNext = false;
            } else {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    // ── IShapeCollection interface ──────────────────────────────────────

    @Override
    public IGroupShape getParentGroup() {
        return parentGroup;
    }

    /**
     * Sets the parent group shape for this collection.
     *
     * @param parentGroup the parent group shape
     */
    public void setParentGroup(IGroupShape parentGroup) {
        this.parentGroup = parentGroup;
    }

    @Override
    public List<IShape> asICollection() {
        return Collections.unmodifiableList(loadShapes());
    }

    @Override
    public Iterable<IShape> asIEnumerable() {
        return Collections.unmodifiableList(loadShapes());
    }

    @Override
    public IShape get(int index) {
        List<IShape> shapes = loadShapes();
        if (index < 0) {
            index = shapes.size() + index;
        }
        if (index < 0 || index >= shapes.size()) {
            throw new IndexOutOfBoundsException("Index " + index + " is out of range");
        }
        return shapes.get(index);
    }

    @Override
    public int size() {
        return loadShapes().size();
    }

    /**
     * Returns an iterator over the shapes in this collection.
     *
     * @return an iterator
     */
    @Override
    public Iterator<IShape> iterator() {
        return Collections.unmodifiableList(loadShapes()).iterator();
    }

    @Override
    public IAutoShape addAutoShape(ShapeType shapeType, double x, double y, double width, double height) {
        return addAutoShapeImpl(null, shapeType, x, y, width, height, true);
    }

    @Override
    public IAutoShape addAutoShape(ShapeType shapeType, double x, double y, double width, double height, boolean createFromTemplate) {
        return addAutoShapeImpl(null, shapeType, x, y, width, height, createFromTemplate);
    }

    @Override
    public IAutoShape insertAutoShape(int index, ShapeType shapeType, double x, double y, double width, double height) {
        return addAutoShapeImpl(index, shapeType, x, y, width, height, true);
    }

    @Override
    public IAutoShape insertAutoShape(int index, ShapeType shapeType, double x, double y, double width, double height, boolean createFromTemplate) {
        return addAutoShapeImpl(index, shapeType, x, y, width, height, createFromTemplate);
    }

    @Override
    public IConnector addConnector(ShapeType shapeType, double x, double y, double width, double height) {
        return addConnectorImpl(null, shapeType, x, y, width, height, true);
    }

    @Override
    public IConnector addConnector(ShapeType shapeType, double x, double y, double width, double height, boolean createFromTemplate) {
        return addConnectorImpl(null, shapeType, x, y, width, height, createFromTemplate);
    }

    @Override
    public IConnector insertConnector(int index, ShapeType shapeType, double x, double y, double width, double height) {
        return addConnectorImpl(index, shapeType, x, y, width, height, true);
    }

    @Override
    public IConnector insertConnector(int index, ShapeType shapeType, double x, double y, double width, double height, boolean createFromTemplate) {
        return addConnectorImpl(index, shapeType, x, y, width, height, createFromTemplate);
    }

    @Override
    public IPictureFrame addPictureFrame(ShapeType shapeType, double x, double y, double width, double height, IPPImage image) {
        return addPictureFrameImpl(null, shapeType, x, y, width, height, image);
    }

    @Override
    public IPictureFrame insertPictureFrame(int index, ShapeType shapeType, double x, double y, double width, double height, IPPImage image) {
        return addPictureFrameImpl(index, shapeType, x, y, width, height, image);
    }

    @Override
    public ITable addTable(double x, double y, double[] colWidths, double[] rowHeights) {
        return addTableImpl(null, x, y, colWidths, rowHeights);
    }

    @Override
    public ITable insertTable(int index, double x, double y, double[] colWidths, double[] rowHeights) {
        return addTableImpl(index, x, y, colWidths, rowHeights);
    }

    @Override
    public int indexOf(IShape shape) {
        return loadShapes().indexOf(shape);
    }

    @Override
    public IShape[] toArray() {
        return loadShapes().toArray(new IShape[0]);
    }

    @Override
    public IShape[] toArray(int startIndex, int count) {
        return loadShapes().subList(startIndex, startIndex + count).toArray(new IShape[0]);
    }

    @Override
    public void reorder(int index, IShape shape) {
        reorderSingle(index, shape);
        invalidateCache();
        save();
    }

    @Override
    public void reorder(int index, IShape[] shapesToReorder) {
        if (spTree == null) return;

        // Remove all target shapes from spTree first
        List<Element> elements = new ArrayList<>();
        for (IShape shape : shapesToReorder) {
            Element elem = findXmlElement(shape);
            if (elem != null && elem.getParentNode() == spTree) {
                spTree.removeChild(elem);
                elements.add(elem);
            }
        }

        // Re-insert at consecutive positions starting from index
        for (int i = 0; i < elements.size(); i++) {
            reinsertAtShapeIndex(index + i, elements.get(i));
        }

        invalidateCache();
        save();
    }

    /**
     * Inserts an element at a given shape index (skipping property elements).
     */
    private void reinsertAtShapeIndex(int targetIndex, Element element) {
        int currentShapeIndex = -1;
        NodeList children = spTree.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (!(children.item(i) instanceof Element elem)) continue;
            String localName = elem.getLocalName();
            if ("nvGrpSpPr".equals(localName) || "grpSpPr".equals(localName)) continue;
            currentShapeIndex++;
            if (currentShapeIndex >= targetIndex) {
                spTree.insertBefore(element, elem);
                return;
            }
        }
        spTree.appendChild(element);
    }

    @Override
    public void removeAt(int index) {
        List<IShape> shapes = loadShapes();
        if (index < 0 || index >= shapes.size()) {
            throw new IndexOutOfBoundsException("Index " + index + " is out of range");
        }
        IShape shape = shapes.get(index);
        Element elem = findXmlElement(shape);
        if (elem != null && spTree != null) {
            spTree.removeChild(elem);
            elementToShape.remove(elem);
        }
        invalidateCache();
        save();
    }

    @Override
    public void remove(IShape shape) {
        Element elem = findXmlElement(shape);
        if (elem != null && spTree != null) {
            spTree.removeChild(elem);
            elementToShape.remove(elem);
        }
        invalidateCache();
        save();
    }

    @Override
    public void clear() {
        if (spTree != null) {
            // Remove all shape elements, keeping nvGrpSpPr and grpSpPr
            List<Element> toRemove = new ArrayList<>();
            NodeList children = spTree.getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                if (children.item(i) instanceof Element elem) {
                    String localName = elem.getLocalName();
                    if (!"nvGrpSpPr".equals(localName) && !"grpSpPr".equals(localName)) {
                        toRemove.add(elem);
                    }
                }
            }
            for (Element elem : toRemove) {
                spTree.removeChild(elem);
            }
        }
        elementToShape.clear();
        invalidateCache();
        save();
    }

    /**
     * Returns the internal list (for framework use).
     *
     * @return the mutable shape list
     */
    List<IShape> getInternalList() {
        return loadShapes();
    }
}
