# Aspose.Slides FOSS

The official open-source Java library by Aspose.Slides for creating, reading, and editing PowerPoint (`.pptx`) presentations.

---

## Installation

Maven:

```xml
<dependency>
    <groupId>org.aspose</groupId>
    <artifactId>aspose-slides-foss</artifactId>
    <version>26.7.0</version>
</dependency>
```

Gradle:

```groovy
implementation 'org.aspose:aspose-slides-foss:26.7.0'
```

**Java 21 or later.** The library uses Java 21 language and library features and will not run on an
older runtime. It has no runtime dependencies of its own: adding it adds one jar and nothing else.

Searching for the library on `search.maven.org` finds nothing, and that is a limitation of that
search index rather than of the release — the artifact is on Maven Central and resolves normally.
[central.sonatype.com](https://central.sonatype.com/artifact/org.aspose/aspose-slides-foss) lists
it; [PUBLISHING.md](PUBLISHING.md) explains why the other one does not.

---

## Quick Start

```java
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.export.SaveFormat;

// Open an existing presentation
try (Presentation prs = new Presentation("input.pptx")) {
    System.out.println("Slides: " + prs.getSlides().size());
    prs.save("output.pptx", SaveFormat.PPTX);
}

// Create a new presentation
try (Presentation prs = new Presentation()) {
    var slide = prs.getSlides().get(0);
    prs.save("new.pptx", SaveFormat.PPTX);
}
```

---

## Features

- **Presentation I/O** — Open, create, and save `.pptx` files with full round-trip fidelity
- **Slides** — Add, remove, clone, and iterate slides
- **Shapes** — AutoShapes, PictureFrames, Tables, Connectors
- **Text** — TextFrame, Paragraph, Portion with character, paragraph, and text frame formatting (including bullets)
- **Fill** — Solid, gradient, pattern, and picture fills
- **Lines** — Width, dash style, arrows, join and alignment
- **Effects** — Outer shadow, glow, soft edge, blur, reflection, inner shadow
- **3D** — Bevel, camera, light rig, material, extrusion depth
- **Document properties** — Core, app, and custom properties
- **Notes slides** — Per-slide notes with header/footer management
- **Comments** — Threaded comments with authors, timestamps, and positions
- **Images** — Embed from file, bytes, or stream

---

## Usage Examples

### Shapes

```java
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.ShapeType;
import org.aspose.slides.foss.IAutoShape;
import org.aspose.slides.foss.ISlide;
import org.aspose.slides.foss.export.SaveFormat;

try (Presentation prs = new Presentation()) {
    ISlide slide = prs.getSlides().get(0);
    IAutoShape shape = slide.getShapes().addAutoShape(ShapeType.RECTANGLE, 50, 50, 300, 100);
    shape.addTextFrame("Hello, world!");
    prs.save("shapes.pptx", SaveFormat.PPTX);
}
```

### Text Formatting

```java
import org.aspose.slides.foss.*;
import org.aspose.slides.foss.drawing.Color;
import org.aspose.slides.foss.export.SaveFormat;

try (Presentation prs = new Presentation()) {
    IAutoShape shape = prs.getSlides().get(0).getShapes()
            .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 150);
    shape.addTextFrame("Formatted text");
    IPortionFormat fmt = shape.getTextFrame().getParagraphs().get(0)
            .getPortions().get(0).getPortionFormat();
    fmt.setFontHeight(24);
    fmt.setFontBold(NullableBool.TRUE);
    fmt.getFillFormat().setFillType(FillType.SOLID);
    fmt.getFillFormat().getSolidFillColor().setColor(Color.fromArgb(255, 0, 70, 127));
    prs.save("text.pptx", SaveFormat.PPTX);
}
```

### Table

```java
import org.aspose.slides.foss.*;
import org.aspose.slides.foss.export.SaveFormat;

try (Presentation prs = new Presentation()) {
    ITable table = prs.getSlides().get(0).getShapes()
            .addTable(50, 50, new double[]{120, 120, 120}, new double[]{40, 40});
    table.getRows().get(0).get(0).getTextFrame().setText("Name");
    table.getRows().get(0).get(1).getTextFrame().setText("Value");
    prs.save("table.pptx", SaveFormat.PPTX);
}
```

### Connector

```java
import org.aspose.slides.foss.*;
import org.aspose.slides.foss.export.SaveFormat;

try (Presentation prs = new Presentation()) {
    ISlide slide = prs.getSlides().get(0);
    IAutoShape box1 = slide.getShapes().addAutoShape(ShapeType.RECTANGLE, 50, 100, 150, 60);
    IAutoShape box2 = slide.getShapes().addAutoShape(ShapeType.RECTANGLE, 350, 100, 150, 60);
    IConnector conn = slide.getShapes().addConnector(ShapeType.BENT_CONNECTOR3, 0, 0, 10, 10);
    conn.setStartShapeConnectedTo(box1);
    conn.setStartShapeConnectionSiteIndex(3);  // right
    conn.setEndShapeConnectedTo(box2);
    conn.setEndShapeConnectionSiteIndex(1);    // left
    prs.save("connector.pptx", SaveFormat.PPTX);
}
```

### Fill

```java
import org.aspose.slides.foss.*;
import org.aspose.slides.foss.drawing.Color;
import org.aspose.slides.foss.export.SaveFormat;

try (Presentation prs = new Presentation()) {
    IAutoShape shape = prs.getSlides().get(0).getShapes()
            .addAutoShape(ShapeType.RECTANGLE, 50, 50, 300, 150);
    shape.getFillFormat().setFillType(FillType.SOLID);
    shape.getFillFormat().getSolidFillColor().setColor(Color.fromArgb(255, 30, 120, 200));
    prs.save("fill.pptx", SaveFormat.PPTX);
}
```

### Notes

```java
import org.aspose.slides.foss.*;
import org.aspose.slides.foss.export.SaveFormat;

try (Presentation prs = new Presentation()) {
    INotesSlide notes = prs.getSlides().get(0).getNotesSlideManager().addNotesSlide();
    notes.getNotesTextFrame().setText("Speaker notes go here.");
    prs.save("notes.pptx", SaveFormat.PPTX);
}
```

### Comments

```java
import org.aspose.slides.foss.*;
import org.aspose.slides.foss.drawing.PointF;
import org.aspose.slides.foss.export.SaveFormat;

import java.time.LocalDateTime;

try (Presentation prs = new Presentation()) {
    ICommentAuthor author = prs.getCommentAuthors().addAuthor("Jane Smith", "JS");
    ISlide slide = prs.getSlides().get(0);
    author.getComments().addComment("Review this slide", slide,
            new PointF(2.0f, 2.0f), LocalDateTime.now());
    prs.save("comments.pptx", SaveFormat.PPTX);
}
```

### Document Properties

```java
import org.aspose.slides.foss.*;
import org.aspose.slides.foss.export.SaveFormat;

try (Presentation prs = new Presentation()) {
    prs.getDocumentProperties().setTitle("Q1 Results");
    prs.getDocumentProperties().setAuthor("Finance Team");
    prs.getDocumentProperties().setCustomPropertyValue("Version", 3);
    prs.save("deck.pptx", SaveFormat.PPTX);
}
```

---

## Limitations

The following areas are not implemented:

- Charts, SmartArt, OLE objects, mathematical text
- Animations and slide transitions
- Export to non-PPTX formats (PDF, HTML, SVG, images)
- VBA macros, digital signatures
- Hyperlinks and action settings

**Writable formats are `SaveFormat.PPTX`, `SaveFormat.PPSX` and `SaveFormat.POTX`.** Every
other `SaveFormat` value — including the macro-enabled `PPTM`, `POTM` and `PPSM`, whose
content types declare a VBA project this library does not write — raises
`UnsupportedOperationException` naming the formats that can be written. Nothing is created
when it does.

**Give the file the extension its format uses.** `save("deck.pptx", SaveFormat.POTX)` writes
a correct template, and PowerPoint refuses to open it: the content type says template and the
name says presentation, and PowerPoint trusts neither over the other. A template is `.potx`,
a slideshow is `.ppsx`, a presentation is `.pptx`.

Unknown XML parts encountered during load are preserved verbatim on save —
opening and re-saving a file will never strip content this library does not understand.

Changes between releases, including behaviour changes and removed API, are in
[CHANGELOG.md](CHANGELOG.md).

---

## Building from source

```
mvn verify -Dgpg.skip=true
```

JDK 21 or later and Maven 3.9 or later are required; the build refuses to start otherwise. That
command runs the whole test suite — unit tests, integration tests, and conformance tests that
unzip the produced `.pptx` and assert on the package — and builds the jar, the sources jar, the
javadoc jar and a CycloneDX SBOM into `target/`. `-Dgpg.skip=true` skips artifact signing, which
is part of `verify` and needs the release key.

Compiler warnings fail the build. Two builds of the same source produce byte-identical jars.

Every push and every pull request runs that build on Linux, Windows and macOS, on Java 21 and
Java 25 (`.github/workflows/build.yml`), and the jar, sources jar, javadoc jar and SBOM from each
run can be downloaded from the run's page.

---

## Links

- [GitHub Repository](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java)
- [The library on Maven Central](https://central.sonatype.com/artifact/org.aspose/aspose-slides-foss)
- [Issue Tracker](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java/issues)

---

## License

[MIT License](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java/blob/main/LICENSE)
