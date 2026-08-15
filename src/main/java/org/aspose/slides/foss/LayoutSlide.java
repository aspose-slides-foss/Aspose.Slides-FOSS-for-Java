package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.LayoutSlidePart;
import org.aspose.slides.foss.internal.pptx.OpcPackage;

import java.util.function.Function;

/**
 * Represents a layout slide.
 */
public final class LayoutSlide extends BaseSlide implements ILayoutSlide {

    private IMasterSlide masterSlide;
    private SlideLayoutType layoutType = SlideLayoutType.CUSTOM;

    // Internal fields for lazy master resolution
    private IPresentation presentationRef;
    private OpcPackage opcPackage;
    private String partName;
    private LayoutSlidePart layoutPart;
    private Function<String, IMasterSlide> masterResolver;

    /**
     * Creates a new empty LayoutSlide.
     */
    public LayoutSlide() {
    }

    /**
     * Internal initialization for a layout slide loaded from a package.
     *
     * @param presentation   the parent Presentation object
     * @param opcPackage     the OPC package
     * @param partName       the part name of this layout slide
     * @param layoutPart     the parsed LayoutSlidePart
     * @param masterResolver callable that resolves a master part name to a MasterSlide,
     *                       or {@code null}
     */
    public void initInternal(IPresentation presentation, OpcPackage opcPackage,
                             String partName, LayoutSlidePart layoutPart,
                             Function<String, IMasterSlide> masterResolver) {
        this.presentationRef = presentation;
        this.opcPackage = opcPackage;
        this.partName = partName;
        this.layoutPart = layoutPart;
        this.masterResolver = masterResolver;
    }

    @Override
    public IMasterSlide getMasterSlide() {
        if (masterResolver != null && masterSlide == null) {
            String masterPartName = layoutPart.getMasterPartName();
            if (masterPartName != null) {
                masterSlide = masterResolver.apply(masterPartName);
            }
        }
        return masterSlide;
    }

    @Override
    public void setMasterSlide(IMasterSlide value) {
        this.masterSlide = value;
    }

    @Override
    public SlideLayoutType getLayoutType() {
        if (layoutPart != null) {
            String typeValue = layoutPart.getLayoutTypeValue();
            return SlideLayoutType.fromValue(typeValue);
        }
        return layoutType;
    }

    /**
     * Sets the layout type (internal).
     *
     * @param layoutType the layout type
     */
    public void setLayoutType(SlideLayoutType layoutType) {
        this.layoutType = layoutType;
    }

    /**
     * {@inheritDoc}
     *
     * <p>For a layout read from a package the name lives in
     * {@code <p:cSld name="…">} of the layout part, not in this object, so it is
     * read from there rather than reported as empty.</p>
     */
    @Override
    public String getName() {
        if (layoutPart != null) {
            String partName = layoutPart.getName();
            if (partName != null && !partName.isEmpty()) {
                return partName;
            }
        }
        return super.getName();
    }

    /**
     * {@inheritDoc}
     *
     * <p>For a layout read from a package the name is written back to
     * {@code <p:cSld name="…">}, so that it survives a save.</p>
     */
    @Override
    public void setName(String name) {
        super.setName(name);
        if (layoutPart != null) {
            layoutPart.setName(name != null ? name : "");
        }
    }

    @Override
    public IPresentation getPresentation() {
        return presentationRef;
    }

    /**
     * Returns the OPC part name of this layout slide (e.g. {@code "ppt/slideLayouts/slideLayout1.xml"}).
     *
     * @return the part name, or {@code null} if not loaded from a package
     */
    String getPartName() {
        return partName;
    }
}
