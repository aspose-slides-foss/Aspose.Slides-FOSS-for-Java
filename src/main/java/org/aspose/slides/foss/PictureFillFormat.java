package org.aspose.slides.foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.Map;

/**
 * Represents a picture fill style.
 *
 * <p>Wraps an OOXML {@code <a:blipFill>} element, reading and writing
 * crop, stretch-offset, and tile properties directly from/to XML.</p>
 */
public final class PictureFillFormat extends PVIObject implements IPictureFillFormat, IFillParamSource {

    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final double EMU_PER_POINT = 12700.0;

    private static final Map<String, RectangleAlignment> ALIGN_MAP = Map.of(
            "tl", RectangleAlignment.TOP_LEFT,
            "t", RectangleAlignment.TOP,
            "tr", RectangleAlignment.TOP_RIGHT,
            "l", RectangleAlignment.LEFT,
            "ctr", RectangleAlignment.CENTER,
            "r", RectangleAlignment.RIGHT,
            "bl", RectangleAlignment.BOTTOM_LEFT,
            "b", RectangleAlignment.BOTTOM,
            "br", RectangleAlignment.BOTTOM_RIGHT
    );
    private static final Map<RectangleAlignment, String> ALIGN_MAP_REV = Map.of(
            RectangleAlignment.TOP_LEFT, "tl",
            RectangleAlignment.TOP, "t",
            RectangleAlignment.TOP_RIGHT, "tr",
            RectangleAlignment.LEFT, "l",
            RectangleAlignment.CENTER, "ctr",
            RectangleAlignment.RIGHT, "r",
            RectangleAlignment.BOTTOM_LEFT, "bl",
            RectangleAlignment.BOTTOM, "b",
            RectangleAlignment.BOTTOM_RIGHT, "br"
    );

    private static final Map<String, TileFlip> FLIP_MAP = Map.of(
            "none", TileFlip.NO_FLIP,
            "x", TileFlip.FLIP_X,
            "y", TileFlip.FLIP_Y,
            "xy", TileFlip.FLIP_BOTH
    );
    private static final Map<TileFlip, String> FLIP_MAP_REV = Map.of(
            TileFlip.NO_FLIP, "none",
            TileFlip.FLIP_X, "x",
            TileFlip.FLIP_Y, "y",
            TileFlip.FLIP_BOTH, "xy"
    );

    private Element blipFill;
    private Runnable saveCallback;
    private int dpi = -1;

    /** The package and part this fill is serialized into, when it belongs to one. */
    private org.aspose.slides.foss.internal.pptx.OpcPackage partPackage;
    private String partName;

    /**
     * Records which package part this fill is serialized into, so that the picture it
     * holds can declare its image relationship in that part's {@code .rels}.
     *
     * @param opcPackage the package, or {@code null} if unattached
     * @param partName   the part name, or {@code null} if unattached
     */
    void setPartContext(org.aspose.slides.foss.internal.pptx.OpcPackage opcPackage, String partName) {
        this.partPackage = opcPackage;
        this.partName = partName;
    }

    /**
     * Creates an uninitialized PictureFillFormat. Call {@link #initInternal} to complete setup.
     */
    public PictureFillFormat() {
    }

    /**
     * Creates a new PictureFillFormat backed by the given {@code <a:blipFill>} element.
     *
     * @param blipFill     the blipFill XML element
     * @param saveCallback callback invoked after mutations; may be {@code null}
     */
    public PictureFillFormat(Element blipFill, Runnable saveCallback) {
        this.blipFill = blipFill;
        this.saveCallback = saveCallback;
    }

    /**
     * Internal initialisation with the {@code <a:blipFill>} XML element.
     *
     * @param blipFillElement the {@code <a:blipFill>} element
     * @param parentSlide     the parent slide object
     * @param saveCallback    callback invoked after mutations; may be {@code null}
     */
    public void initInternal(Element blipFillElement, IBaseSlide parentSlide, Runnable saveCallback) {
        this.blipFill = blipFillElement;
        this.parentSlide = parentSlide;
        this.saveCallback = saveCallback;
    }

    /**
     * Persists changes via the save callback if one is configured.
     */
    void save() {
        if (saveCallback != null) {
            saveCallback.run();
        }
    }

    // --- XML element helpers ---

    private Element findChild(Element parent, String localName) {
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            if (children.item(i) instanceof Element el
                    && localName.equals(el.getLocalName())) {
                return el;
            }
        }
        return null;
    }

    private Element ensureChild(Element parent, String localName) {
        Element child = findChild(parent, localName);
        if (child == null) {
            Document doc = parent.getOwnerDocument();
            child = doc.createElementNS(NS_A, "a:" + localName);
            parent.appendChild(child);
        }
        return child;
    }

    private void removeChild(Element parent, String localName) {
        Element child = findChild(parent, localName);
        if (child != null) {
            parent.removeChild(child);
        }
    }

    /**
     * Returns the {@code <a:blip>} child element, or {@code null} if absent.
     *
     * @return the blip element, or {@code null}
     */
    Element getBlip() {
        return findChild(blipFill, "blip");
    }

    /**
     * Returns the {@code <a:srcRect>} element, creating it if absent.
     *
     * @return the srcRect element (never {@code null})
     */
    Element getOrCreateSrcRect() {
        return ensureChild(blipFill, "srcRect");
    }

    /**
     * Returns the {@code <a:stretch>} child element, or {@code null} if absent.
     *
     * @return the stretch element, or {@code null}
     */
    Element getStretch() {
        return findChild(blipFill, "stretch");
    }

    /**
     * Returns the {@code <a:fillRect>} element under {@code <a:stretch>},
     * creating both elements if absent.
     *
     * @return the fillRect element (never {@code null})
     */
    Element getOrCreateFillRect() {
        Element stretch = findChild(blipFill, "stretch");
        if (stretch == null) {
            stretch = ensureChild(blipFill, "stretch");
        }
        return ensureChild(stretch, "fillRect");
    }

    /**
     * Returns the {@code <a:tile>} child element, or {@code null} if absent.
     *
     * @return the tile element, or {@code null}
     */
    Element getTile() {
        return findChild(blipFill, "tile");
    }

    /**
     * Returns the {@code <a:tile>} element, creating it if absent and
     * removing any {@code <a:stretch>} element (switches to tile mode).
     *
     * @return the tile element (never {@code null})
     */
    Element ensureTile() {
        Element tile = getTile();
        if (tile != null) {
            return tile;
        }
        removeChild(blipFill, "stretch");
        return ensureChild(blipFill, "tile");
    }

    // --- DPI ---

    @Override
    public int getDpi() {
        return dpi;
    }

    @Override
    public void setDpi(int value) {
        this.dpi = value;
        save();
    }

    // --- Picture fill mode ---

    @Override
    public PictureFillMode getPictureFillMode() {
        if (findChild(blipFill, "tile") != null) {
            return PictureFillMode.TILE;
        }
        return PictureFillMode.STRETCH;
    }

    @Override
    public void setPictureFillMode(PictureFillMode value) {
        if (value == PictureFillMode.TILE) {
            removeChild(blipFill, "stretch");
            if (findChild(blipFill, "tile") == null) {
                ensureChild(blipFill, "tile");
            }
        } else {
            removeChild(blipFill, "tile");
            if (findChild(blipFill, "stretch") == null) {
                Element stretch = ensureChild(blipFill, "stretch");
                ensureChild(stretch, "fillRect");
            }
        }
        save();
    }

    // --- Picture ---

    @Override
    public ISlidesPicture getPicture() {
        Element blip = getBlip();
        if (blip == null) {
            return null;
        }
        var picture = new Picture();
        if (partPackage != null && partName != null) {
            picture.initInternal(blip, partPackage, partName, parentSlide);
        } else {
            picture.initInternal(blip, parentSlide);
        }
        return picture;
    }

    // --- Crop properties (backed by a:srcRect attributes) ---

    /**
     * Gets a crop value from the {@code <a:srcRect>} element.
     *
     * @param attr the attribute name ({@code "l"}, {@code "t"}, {@code "r"}, or {@code "b"})
     * @return the crop percentage (e.g., 10.0 for 10%)
     */
    float getCropValue(String attr) {
        Element srcRect = findChild(blipFill, "srcRect");
        if (srcRect == null) {
            return 0.0f;
        }
        String val = srcRect.getAttribute(attr);
        if (val == null || val.isEmpty()) {
            return 0.0f;
        }
        return Integer.parseInt(val) / 1000.0f;
    }

    /**
     * Sets a crop value on the {@code <a:srcRect>} element.
     *
     * @param attr  the attribute name ({@code "l"}, {@code "t"}, {@code "r"}, or {@code "b"})
     * @param value the crop percentage (e.g., 10.0 for 10%)
     */
    void setCropValue(String attr, float value) {
        Element srcRect = getOrCreateSrcRect();
        srcRect.setAttribute(attr, String.valueOf(Math.round(value * 1000)));
        save();
    }

    @Override
    public float getCropLeft() {
        return getCropValue("l");
    }

    @Override
    public void setCropLeft(float value) {
        setCropValue("l", value);
    }

    @Override
    public float getCropTop() {
        return getCropValue("t");
    }

    @Override
    public void setCropTop(float value) {
        setCropValue("t", value);
    }

    @Override
    public float getCropRight() {
        return getCropValue("r");
    }

    @Override
    public void setCropRight(float value) {
        setCropValue("r", value);
    }

    @Override
    public float getCropBottom() {
        return getCropValue("b");
    }

    @Override
    public void setCropBottom(float value) {
        setCropValue("b", value);
    }

    // --- Stretch offset properties (backed by a:fillRect under a:stretch) ---

    /**
     * Gets a stretch offset value from the {@code <a:fillRect>} element under {@code <a:stretch>}.
     *
     * @param attr the attribute name ({@code "l"}, {@code "t"}, {@code "r"}, or {@code "b"})
     * @return the stretch offset percentage
     */
    float getStretchOffset(String attr) {
        Element stretch = getStretch();
        if (stretch == null) {
            return 0.0f;
        }
        Element fillRect = findChild(stretch, "fillRect");
        if (fillRect == null) {
            return 0.0f;
        }
        String val = fillRect.getAttribute(attr);
        if (val == null || val.isEmpty()) {
            return 0.0f;
        }
        return Integer.parseInt(val) / 1000.0f;
    }

    /**
     * Sets a stretch offset value on the {@code <a:fillRect>} element under {@code <a:stretch>}.
     *
     * @param attr  the attribute name ({@code "l"}, {@code "t"}, {@code "r"}, or {@code "b"})
     * @param value the stretch offset percentage
     */
    void setStretchOffset(String attr, float value) {
        Element fillRect = getOrCreateFillRect();
        fillRect.setAttribute(attr, String.valueOf(Math.round(value * 1000)));
        save();
    }

    @Override
    public float getStretchOffsetLeft() {
        return getStretchOffset("l");
    }

    @Override
    public void setStretchOffsetLeft(float value) {
        setStretchOffset("l", value);
    }

    @Override
    public float getStretchOffsetTop() {
        return getStretchOffset("t");
    }

    @Override
    public void setStretchOffsetTop(float value) {
        setStretchOffset("t", value);
    }

    @Override
    public float getStretchOffsetRight() {
        return getStretchOffset("r");
    }

    @Override
    public void setStretchOffsetRight(float value) {
        setStretchOffset("r", value);
    }

    @Override
    public float getStretchOffsetBottom() {
        return getStretchOffset("b");
    }

    @Override
    public void setStretchOffsetBottom(float value) {
        setStretchOffset("b", value);
    }

    // --- Tile offset properties (backed by a:tile tx/ty in EMU) ---

    @Override
    public float getTileOffsetX() {
        Element tile = getTile();
        if (tile == null) {
            return 0.0f;
        }
        String val = tile.getAttribute("tx");
        if (val == null || val.isEmpty()) {
            return 0.0f;
        }
        return (float) (Integer.parseInt(val) / EMU_PER_POINT);
    }

    @Override
    public void setTileOffsetX(float value) {
        Element tile = ensureTile();
        tile.setAttribute("tx", String.valueOf(Math.round(value * EMU_PER_POINT)));
        save();
    }

    @Override
    public float getTileOffsetY() {
        Element tile = getTile();
        if (tile == null) {
            return 0.0f;
        }
        String val = tile.getAttribute("ty");
        if (val == null || val.isEmpty()) {
            return 0.0f;
        }
        return (float) (Integer.parseInt(val) / EMU_PER_POINT);
    }

    @Override
    public void setTileOffsetY(float value) {
        Element tile = ensureTile();
        tile.setAttribute("ty", String.valueOf(Math.round(value * EMU_PER_POINT)));
        save();
    }

    // --- Tile scale properties (backed by a:tile sx/sy, percentage * 1000) ---

    @Override
    public float getTileScaleX() {
        Element tile = getTile();
        if (tile == null) {
            return 100.0f;
        }
        String val = tile.getAttribute("sx");
        if (val == null || val.isEmpty()) {
            return 100.0f;
        }
        return Integer.parseInt(val) / 1000.0f;
    }

    @Override
    public void setTileScaleX(float value) {
        Element tile = ensureTile();
        tile.setAttribute("sx", String.valueOf(Math.round(value * 1000)));
        save();
    }

    @Override
    public float getTileScaleY() {
        Element tile = getTile();
        if (tile == null) {
            return 100.0f;
        }
        String val = tile.getAttribute("sy");
        if (val == null || val.isEmpty()) {
            return 100.0f;
        }
        return Integer.parseInt(val) / 1000.0f;
    }

    @Override
    public void setTileScaleY(float value) {
        Element tile = ensureTile();
        tile.setAttribute("sy", String.valueOf(Math.round(value * 1000)));
        save();
    }

    // --- Tile alignment (backed by a:tile algn attribute) ---

    @Override
    public RectangleAlignment getTileAlignment() {
        Element tile = getTile();
        if (tile == null) {
            return RectangleAlignment.NOT_DEFINED;
        }
        String val = tile.getAttribute("algn");
        if (val == null || val.isEmpty()) {
            return RectangleAlignment.NOT_DEFINED;
        }
        RectangleAlignment result = ALIGN_MAP.get(val);
        return result != null ? result : RectangleAlignment.NOT_DEFINED;
    }

    @Override
    public void setTileAlignment(RectangleAlignment value) {
        Element tile = ensureTile();
        if (value == RectangleAlignment.NOT_DEFINED) {
            tile.removeAttribute("algn");
        } else {
            String ooxmlVal = ALIGN_MAP_REV.get(value);
            if (ooxmlVal != null) {
                tile.setAttribute("algn", ooxmlVal);
            }
        }
        save();
    }

    // --- Tile flip (backed by a:tile flip attribute) ---

    @Override
    public TileFlip getTileFlip() {
        Element tile = getTile();
        if (tile == null) {
            return TileFlip.NOT_DEFINED;
        }
        String val = tile.getAttribute("flip");
        if (val == null || val.isEmpty()) {
            return TileFlip.NOT_DEFINED;
        }
        TileFlip result = FLIP_MAP.get(val);
        return result != null ? result : TileFlip.NOT_DEFINED;
    }

    @Override
    public void setTileFlip(TileFlip value) {
        Element tile = ensureTile();
        if (value == TileFlip.NOT_DEFINED) {
            tile.removeAttribute("flip");
        } else {
            String ooxmlVal = FLIP_MAP_REV.get(value);
            if (ooxmlVal != null) {
                tile.setAttribute("flip", ooxmlVal);
            }
        }
        save();
    }
}
