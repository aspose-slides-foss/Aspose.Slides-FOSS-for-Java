# Aspose.Slides FOSS — Guide for AI Agents

You are working with `aspose-slides-foss`, the official open-source Java library by Aspose.Slides for creating, reading, and editing PowerPoint (.pptx) presentations.

## Build

```bash
mvn compile                  # compile
mvn test                     # run all tests
mvn verify -Dgpg.skip=true   # tests, plus the jar, sources, javadoc and SBOM
```

Requires JDK 21+ and Maven 3.9+. No runtime dependencies beyond the JDK. Compiler warnings fail the
build. See CONTRIBUTING.md.

## Core Concepts

- **`Presentation`** is the root object. It owns slides, masters, layouts, images, document properties, and comments.
- Always use `Presentation` in a try-with-resources block to ensure proper cleanup.
- Save with `prs.save("out.pptx", SaveFormat.PPTX)`. `SaveFormat.PPSX` and `SaveFormat.POTX` are
  also written; every other `SaveFormat` value raises `UnsupportedOperationException`.
- Unknown XML parts are preserved verbatim on save — round-tripping is safe.

## Import Pattern

```java
import org.aspose.slides.foss.*;
import org.aspose.slides.foss.export.SaveFormat;
import org.aspose.slides.foss.drawing.Color;
import org.aspose.slides.foss.drawing.PointF;
import org.aspose.slides.foss.drawing.SizeF;
import org.aspose.slides.foss.drawing.Size;
```

## Quick Reference

### Create a presentation

```java
try (Presentation prs = new Presentation()) {
    ISlide slide = prs.getSlides().get(0);  // first slide exists by default
    prs.save("new.pptx", SaveFormat.PPTX);
}
```

### Open an existing file

```java
try (Presentation prs = new Presentation("input.pptx")) {
    for (ISlide slide : prs.getSlides()) {
        // ISlideCollection is Iterable; IShapeCollection is not - use asIEnumerable() or an index.
        for (IShape shape : slide.getShapes().asIEnumerable()) {
            System.out.println(shape.getName());
        }
    }
    prs.save("output.pptx", SaveFormat.PPTX);
}
```

### Add shapes

```java
IAutoShape shape = slide.getShapes().addAutoShape(ShapeType.RECTANGLE, x, y, width, height);
shape.addTextFrame("Hello");
```

Coordinates and dimensions are in points (1 point = 1/72 inch).

### Text formatting

```java
IPortionFormat fmt = shape.getTextFrame().getParagraphs().get(0)
        .getPortions().get(0).getPortionFormat();
fmt.setFontHeight(24);
fmt.setFontBold(NullableBool.TRUE);
fmt.setFontItalic(NullableBool.TRUE);
fmt.getFillFormat().setFillType(FillType.SOLID);
fmt.getFillFormat().getSolidFillColor().setColor(Color.fromArgb(255, 0, 70, 127));
```

### Tables

```java
double[] colWidths = {120.0, 120.0, 120.0};
double[] rowHeights = {40.0, 40.0};
ITable table = slide.getShapes().addTable(x, y, colWidths, rowHeights);
table.getRows().get(0).get(0).getTextFrame().setText("Header");
```

### Connectors

```java
IConnector conn = slide.getShapes().addConnector(ShapeType.BENT_CONNECTOR3, 0, 0, 10, 10);
conn.setStartShapeConnectedTo(shapeA);
conn.setStartShapeConnectionSiteIndex(3);  // 0=top, 1=left, 2=bottom, 3=right
conn.setEndShapeConnectedTo(shapeB);
conn.setEndShapeConnectionSiteIndex(1);
```

### Fills

```java
shape.getFillFormat().setFillType(FillType.SOLID);
shape.getFillFormat().getSolidFillColor().setColor(Color.fromArgb(255, 30, 120, 200));
```

Also supports: `FillType.GRADIENT`, `FillType.PATTERN`, `FillType.PICTURE`, `FillType.NO_FILL`.

### Images

```java
byte[] imageData = Files.readAllBytes(Path.of("photo.png"));
IPPImage image = prs.getImages().addImage(imageData);
slide.getShapes().addPictureFrame(ShapeType.RECTANGLE, x, y, w, h, image);
```

### Notes

```java
INotesSlide notes = slide.getNotesSlideManager().addNotesSlide();
notes.getNotesTextFrame().setText("Speaker notes here.");
```

### Comments

```java
import java.time.LocalDateTime;

ICommentAuthor author = prs.getCommentAuthors().addAuthor("Jane Smith", "JS");
author.getComments().addComment("Review this", slide, new PointF(2.0f, 2.0f), LocalDateTime.now());
```

### Document properties

```java
prs.getDocumentProperties().setTitle("Quarterly Report");
prs.getDocumentProperties().setAuthor("Finance Team");
prs.getDocumentProperties().setCustomPropertyValue("Version", 3);
```

### Slide operations

```java
prs.getSlides().addEmptySlide(prs.getLayoutSlides().get(0));  // add slide
prs.getSlides().removeAt(1);                                    // remove by index
ISlide cloned = prs.getSlides().addClone(prs.getSlides().get(0)); // clone slide
slide.setHidden(true);                                          // hide slide
```

### Effects and 3D

```java
// Outer shadow
IEffectFormat ef = shape.getEffectFormat();
ef.enableOuterShadowEffect();
ef.getOuterShadowEffect().setBlurRadius(10);
ef.getOuterShadowEffect().setDistance(5);

// 3D bevel
IThreeDFormat td = shape.getThreeDFormat();
td.getBevelTop().setBevelType(BevelPresetType.CIRCLE);
td.getBevelTop().setHeight(6);
td.getBevelTop().setWidth(6);
```

### Line formatting

```java
ILineFormat lf = shape.getLineFormat();
lf.setWidth(2.5);
lf.setDashStyle(LineDashStyle.DASH_DOT);
lf.getFillFormat().setFillType(FillType.SOLID);
lf.getFillFormat().getSolidFillColor().setColor(Color.RED);
```

## Package Structure

```
src/main/java/org/aspose/slides/foss/
    Presentation.java         # Root object
    Slide.java                # Slide, LayoutSlide, MasterSlide
    ShapeCollection.java      # Shape management
    AutoShape.java            # AutoShape with text frames
    Table.java                # Tables, rows, columns, cells
    Connector.java            # Shape-to-shape connectors
    TextFrame.java            # Text content model
    FillFormat.java           # Fill styling
    LineFormat.java           # Line styling
    EffectFormat.java         # Visual effects
    ThreeDFormat.java         # 3D formatting
    Comment.java              # Slide comments
    DocumentProperties.java   # Document metadata
    internal/                 # Implementation internals (do not import directly)
    drawing/                  # Color, PointF, SizeF, Size
    export/                   # SaveFormat
    effects/                  # Effect-related classes
    theme/                    # Theme-related types
```

## Do

- Always wrap `Presentation` in a try-with-resources block
- Use `SaveFormat.PPTX`, `PPSX` or `POTX` when saving — every other value raises, and the file
  extension has to match the format you asked for
- Use `Color.fromArgb(a, r, g, b)` or named constants like `Color.RED`, `Color.BLUE`
- Access slides via `prs.getSlides().get(index)` — slides are 0-indexed
- Use `NullableBool` enum (`NullableBool.FALSE`, `NullableBool.TRUE`, `NullableBool.NOT_DEFINED`) for boolean formatting properties like `setFontBold`
- Import drawing types from `org.aspose.slides.foss.drawing`

## Don't

- Don't import from `org.aspose.slides.foss.internal` — it is a private implementation detail
- Don't attempt PDF, HTML, SVG, or image export — the library does not render or convert
- Don't use charts, SmartArt, animations, or VBA — they are not supported
- Don't modify the public API class signatures — they are fixed

## Limitations

Not supported:

- Charts, SmartArt, OLE objects, mathematical text
- Animations and slide transitions
- Export to PDF, HTML, SVG, or images; only PPTX, PPSX and POTX are written
- VBA macros, digital signatures
- Hyperlinks and action settings

## Links

- [GitHub](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java)
- [Issue Tracker](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java/issues)
