# Aspose.Slides FOSS for Java

[![Build and test](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java/actions/workflows/build.yml/badge.svg)](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java/actions/workflows/build.yml)
[![Maven Central](https://img.shields.io/maven-central/v/org.aspose/aspose-slides-foss.svg)](https://central.sonatype.com/artifact/org.aspose/aspose-slides-foss)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Java 21+](https://img.shields.io/badge/Java-21%2B-orange.svg)](https://adoptium.net/)

An MIT-licensed Java library that creates, reads and edits PowerPoint `.pptx` presentations by
building the Office Open XML package itself — no PowerPoint installation, no native code, and no
runtime dependencies at all. The CycloneDX SBOM the build produces lists zero components.

It is for developers who generate or modify decks on a server or in a build, and who are working
*inside* the file format. It is not a renderer or a converter: there is no PDF, HTML, SVG or image
output, and everything it cannot do is listed under [Limitations](#limitations).

---

## Requirements

| | |
|---|---|
| Java | **21 or later.** The sources use `List.getFirst()`, `List.removeLast()` and an unconditional `instanceof` pattern, so they do not compile below 21, and the build refuses to start on an older JDK. |
| Runtime dependencies | none |
| Tested on | Linux, Windows and macOS, on Java 21 and Java 25 — every push and every pull request |

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

**Before you depend on `26.7.0`, read this.** It predates the XML hardening described under
*Security* in [CHANGELOG.md](CHANGELOG.md), and [SECURITY.md](SECURITY.md) lists it as not
supported for that reason. If your program opens `.pptx` files that come from somewhere you do not
control, [build from source](#building-from-source) rather than depending on `26.7.0`.

`26.7.0`, published 2026-07-27, is the only version on Maven Central today —
`repo1.maven.org/maven2/org/aspose/aspose-slides-foss/maven-metadata.xml` lists it as both
`<latest>` and `<release>`. **This source tree is `26.8.0` and is not published yet**, and the
behaviour changes at the top of [CHANGELOG.md](CHANGELOG.md) — above all the refusal to write a save
format it cannot produce — are in the source and not in `26.7.0`. Until `26.8.0` is released,
[build from source](#building-from-source) if you need them.

Searching for the library on `search.maven.org` finds nothing, and that is a limitation of that
search index rather than of the release — the artifact resolves normally from `repo1.maven.org`.
[central.sonatype.com](https://central.sonatype.com/artifact/org.aspose/aspose-slides-foss) lists
it; [PUBLISHING.md](PUBLISHING.md) explains why the other one does not.

The jar declares `Automatic-Module-Name: org.aspose.slides.foss`, so it can go on the module path
under a stable name.

---

## Quick start

```java
import java.io.IOException;

import org.aspose.slides.foss.IAutoShape;
import org.aspose.slides.foss.ISlide;
import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.ShapeType;
import org.aspose.slides.foss.export.SaveFormat;

public class QuickStart {
    public static void main(String[] args) throws IOException {
        // Create a deck, put one shape with some text on the first slide, save it.
        try (Presentation prs = new Presentation()) {
            ISlide slide = prs.getSlides().get(0);
            IAutoShape shape = slide.getShapes()
                    .addAutoShape(ShapeType.RECTANGLE, 50, 50, 400, 100);
            shape.addTextFrame("Hello from Aspose.Slides FOSS");
            prs.save("hello.pptx", SaveFormat.PPTX);
        }

        // Read it back.
        try (Presentation prs = new Presentation("hello.pptx")) {
            ISlide slide = prs.getSlides().get(0);
            IAutoShape shape = (IAutoShape) slide.getShapes().get(0);
            System.out.println("slides: " + prs.getSlides().size());
            System.out.println("shapes: " + slide.getShapes().size());
            System.out.println("text:   " + shape.getTextFrame().getText());
        }
    }
}
```

Compiled against this tree and run, it prints

```
slides: 1
shapes: 1
text:   Hello from Aspose.Slides FOSS
```

and leaves a 5,458-byte `hello.pptx` behind.

Three things worth knowing before the second program you write:

- **`save` and the `Presentation(String)` constructor throw `IOException`.** `close()` does not, so
  try-with-resources needs nothing extra.
- **`IShapeCollection.get(int)` returns `IShape`.** Cast to `IAutoShape`, `ITable`, `IConnector` or
  `IPictureFrame` to reach the members only those have. `ISlideCollection` implements `Iterable`, so
  a for-each over `getSlides()` works; `IShapeCollection` does not, so iterate it with an index or
  with `getShapes().asIEnumerable()`.
- **Coordinates and sizes are points** — 1 point = 1/72 inch. Slide and shape indices are
  zero-based.

---

## Examples

Every example below was compiled against this tree and run.

### Text formatting

```java
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

### Tables

```java
try (Presentation prs = new Presentation()) {
    ITable table = prs.getSlides().get(0).getShapes()
            .addTable(50, 50, new double[]{120, 120, 120}, new double[]{40, 40});
    table.getRows().get(0).get(0).getTextFrame().setText("Name");
    table.getRows().get(0).get(1).getTextFrame().setText("Value");
    prs.save("table.pptx", SaveFormat.PPTX);
}
```

### Connectors bound to shapes

```java
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

### Pictures

```java
try (Presentation prs = new Presentation()) {
    byte[] png = Files.readAllBytes(Path.of("photo.png"));
    IPPImage image = prs.getImages().addImage(png);          // or addImage(InputStream)
    prs.getSlides().get(0).getShapes()
            .addPictureFrame(ShapeType.RECTANGLE, 50, 50, 200, 150, image);
    prs.save("picture.pptx", SaveFormat.PPTX);
}
```

### Notes and comments

```java
try (Presentation prs = new Presentation()) {
    ISlide slide = prs.getSlides().get(0);

    INotesSlide notes = slide.getNotesSlideManager().addNotesSlide();
    notes.getNotesTextFrame().setText("Speaker notes go here.");

    ICommentAuthor author = prs.getCommentAuthors().addAuthor("Jane Smith", "JS");
    IComment root = author.getComments()
            .addComment("Review this slide", slide, new PointF(2.0f, 2.0f), LocalDateTime.now());
    IComment reply = author.getComments()
            .addComment("Done", slide, new PointF(2.0f, 2.0f), LocalDateTime.now());
    reply.setParentComment(root);   // writes ppt/threadedComments/ as well as the classic list

    prs.save("notes-and-comments.pptx", SaveFormat.PPTX);
}
```

### Streams

```java
try (Presentation prs = new Presentation(Files.newInputStream(Path.of("in.pptx")));
     OutputStream out = Files.newOutputStream(Path.of("out.pptx"))) {
    prs.save(out, SaveFormat.PPTX);
}
```

### A subset of slides

```java
try (Presentation prs = new Presentation("deck.pptx")) {
    // Zero-based. Repeats are ignored; the slides kept stay in document order.
    prs.save("first-and-third.pptx", new int[]{0, 2}, SaveFormat.PPTX);
}
```

---

## What it can do

Everything below was exercised through the public API and then confirmed by reading the XML inside
the file that came out — not by asking the library to read its own file back.

- **Presentations** — create, open a path or a stream, save to a path or a stream, save a subset of
  slides; `Presentation` is `AutoCloseable`.
- **Slides** — add empty, insert, remove (the part, its relationship and its content-type override
  all go), clone, hide (`<p:sld show="0">`), enumerate; the masters and layouts of a loaded deck are
  enumerable.
- **Shapes** — AutoShapes for all **190** `ShapeType` values, writing **187** distinct
  `<a:prstGeom prst="…">` tokens between them. `NOT_DEFINED` and `CUSTOM` have no preset and come
  out as `rect`; `ROUND_RECTANGLE` and `ROUND_CORNER_RECTANGLE` both write `roundRect`. Also
  picture frames, tables, and connectors bound to shapes by connection site.
- **Text** — text frames, paragraphs, portions; character formatting (bold, italic, size, spacing,
  caps, latin and east-asian fonts), paragraph formatting (alignment, indent, margins, spacing, and
  symbol, numbered and picture bullets), text-frame formatting (margins, wrap, anchor, autofit,
  columns).
- **Fill** — solid, gradient, pattern, picture and no-fill, on shapes, lines, text portions and
  table cells alike.
- **Lines** — width, dash style, cap, join, compound style, alignment, arrowheads at both ends.
- **Effects** — outer shadow, inner shadow, glow, soft edge, reflection, blur, preset shadow and
  fill overlay, each landing inside `<a:effectLst>`.
- **3-D** — bevel top and bottom, extrusion, contour, material, a camera with a preset and a light
  rig, written where `CT_ShapeProperties` puts them.
- **Tables** — rows, columns, cells, cell merging, cell fill, borders, margins, and the
  `<a:graphicFrameLocks>` the schema requires on the frame.
- **Pictures** — embed from a byte array, an `IImage` or an `InputStream`, through a real
  relationship declared in the owning part's `.rels`.
- **Notes** — a notes slide per slide, with header/footer management.
- **Comments** — authors, and comments with position and timestamp, on the classic `ppt/comments/`
  list; a reply additionally writes `ppt/threadedComments/` and `ppt/authors.xml`.
- **Document properties** — core, extended and custom, with the extended counts recomputed from the
  document on save.
- **Unknown parts** — parts the library does not model are carried through a load and a save
  unchanged.

---

## Limitations

This section is the point of this file. Nothing here is a "coming soon"; it is what the API does not
contain today.

### Not in the public API

| | |
|---|---|
| Charts | no `IShapeCollection.addChart` |
| SmartArt, OLE objects, video, audio | not modelled |
| Group shapes | no `IShapeCollection.addGroupShape` — no shape can hold child shapes |
| Animations and slide transitions | no `ISlide.getTimeline`, no `ISlide.getSlideShowTransition` |
| Hyperlinks | no `setHyperlinkClick` on a shape or on a text portion |
| Sections | no `IPresentation.getSections` |
| Slide backgrounds | no `ISlide.getBackground` |
| Themes | no `IPresentation.getMasterTheme` |
| Slide size | no `IPresentation.getSlideSize`. A new deck is 4:3 (`cx="9144000" cy="6858000" type="screen4x3"`) and there is no supported way to change it |
| Rendering and conversion | no PDF, HTML, SVG, image or text export of any kind |
| VBA macros, digital signatures, encryption | not modelled |

### A known defect: cloning a master

`getMasters().addClone(...)` returns, and the collection then reports a size of 2 — but the saved
package contains one `<p:sldMasterId>` and one `ppt/slideMasters/slideMaster1.xml`. The clone never
reaches the file, and nothing reports a problem. Do not rely on it.

### Save formats

`SaveFormat` declares **25** values. `save` writes **three** of them and raises
`UnsupportedOperationException` for the other **22**. It never writes a package under a name that
claims to be a format it did not produce, and nothing is left on disk when it refuses.

| Written | Refused |
|---|---|
| `PPTX`, `PPSX`, `POTX` | `PPT`, `PDF`, `XPS`, `TIFF`, `ODP`, `PPTM`, `PPSM`, `POTM`, `HTML`, `HTML5`, `SWF`, `OTP`, `PPS`, `POT`, `FODP`, `GIF`, `MD`, `XML`, `SVG`, `JPEG`, `PNG`, `BMP` |

The macro-enabled formats are refused because their content type declares a VBA project this library
does not write. The refusal names the alternatives:

```
Export format 'Pdf' is not supported. Writable formats are: Pptx, Ppsx, Potx
```

The three that are written are three genuinely different packages, each with its own main-part
content type, so a template saved as a template really is a template. **Give the file the extension
its format uses.** `save("deck.pptx", SaveFormat.POTX)` writes a correct template that PowerPoint
refuses to open: the content type says template, the name says presentation, and PowerPoint trusts
neither over the other. A template is `.potx`, a slideshow `.ppsx`, a presentation `.pptx`.

### What round-tripping does and does not guarantee

Opening a deck and saving it again preserves **every part**, and most of them byte for byte. It does
not rewrite nothing at all. Measured on a 41-part deck written by Apache POI — an independent
writer, not this library:

| | |
|---|---|
| Parts in / out | 41 / 41 |
| Parts dropped | 0 |
| Parts added | 0 |
| Parts byte-identical afterwards | 36 of 41 |
| Parts rewritten | `docProps/app.xml`, `ppt/presentation.xml`, and the three slide parts |

`docProps/app.xml` is regenerated from the presentation; the others are rebuilt from the package
model. Parts the library has no model for — `presProps`, `viewProps`, `tableStyles`, themes, unused
layouts — are among the 36 that come back unchanged. That is a specific, measured claim about one
real deck, and it is deliberately narrower than "full fidelity".

Re-saving is stable as well: a save with nothing modified in between does not change the file size
or the slide XML, so a byte-length delta stays a usable "did anything actually change" signal.

### Text language

A deck written by this library contains no `lang=` attribute anywhere in the package, and a portion
that was never formatted gets no `<a:rPr>` at all. PowerPoint will apply the authoring machine's
default language rather than one the file states.

---

## Choosing an edition

There are four editions of this library and they are **not** interchangeable. Three of them — .NET,
Java and C++ — are the same design in three languages; Python is a larger and different one. The
table was measured on 2026-08-16 by running each edition and reading the XML in the file it
produced.

| | .NET | **Java** | C++ | Python |
|---|---|---|---|---|
| Create, open, round-trip, save | yes | **yes** | yes | yes |
| Text, tables, connectors, fills, all 8 effects, 3-D | yes | **yes** | yes | yes |
| Notes and classic comments | yes | **yes** | yes | yes |
| Threaded comment part | no | **yes** | no | yes |
| Save to a stream | yes | **yes** | **no** | yes |
| Add a picture from a stream | yes | **yes** | **no** | yes |
| Clone a slide, with its text | yes | **yes** | **partial** | yes |
| Clone a master | yes | **no** | **no** | yes |
| Sections | **yes** | **no** | no | no |
| Charts | no | **no** | no | **yes** |
| Animations | no | **no** | no | **yes** |
| Slide transitions | no | **no** | no | **yes** |
| Themes | no | **no** | no | **yes** |
| Slide backgrounds | no | **no** | no | **yes** |
| Group shapes | no | **no** | no | **yes** |
| Hyperlinks | no | **no** | no | **yes** |
| Markdown export | no | **no** | no | **yes** |
| Formats `save` writes | 3 | **3** | 6 | 7 |
| Default slide size of a new deck | 4:3 | **4:3** | 4:3 | **16:9** |
| Layouts in a new deck | 1 | **1** | 1 | **11** |

The short version: **if you need charts, animations, transitions, themes, backgrounds, group shapes
or hyperlinks, none of them exist in this edition — use the Python one.** If you need sections, only
the .NET edition has them. Code written against the Python examples does not port here, and a new
deck is not even the same shape.

- [Aspose.Slides FOSS for .NET](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET)
- [Aspose.Slides FOSS for C++](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Cpp)
- [Aspose.Slides FOSS for Python](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Python)

---

## Building from source

```bash
git clone https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java.git
cd Aspose.Slides-FOSS-for-Java
mvn verify -Dgpg.skip=true
```

JDK 21 or later and Maven 3.9 or later are required; the build refuses to start otherwise. That one
command runs the whole test suite — **3,638 tests**, of which **232 are conformance tests** that
unzip the produced `.pptx` and assert on the package — and builds the jar, the sources jar, the
javadoc jar and a CycloneDX SBOM into `target/`. `-Dgpg.skip=true` skips artifact signing, which is
part of `verify` and needs the release key.

Compiler warnings fail the build, and so does a javadoc error. Two builds of the same source produce
byte-identical jars; CI proves it by building twice and comparing the hashes.

Every push and every pull request runs that build on Linux, Windows and macOS, on Java 21 and Java
25 (`.github/workflows/build.yml`). The jar, sources jar, javadoc jar and SBOM from each run are
downloadable from the run's page.

Releases go to Maven Central through `.github/workflows/maven-central-release.yml`;
[PUBLISHING.md](PUBLISHING.md) documents how one is cut and what fails one.

---

## Documentation and support

This library mirrors the naming of the commercial **Aspose.Slides for Java** product, so that
product's documentation is often the fastest way to understand what a shared type name means. It
describes a much larger API: check the tables above before relying on anything you read there.

- [Aspose.Slides for Java — product page](https://products.aspose.com/slides/java/)
- [Documentation](https://docs.aspose.com/slides/java/)
- [API reference](https://reference.aspose.com/slides/java/)
- [Free support forum](https://forum.aspose.com/c/slides/11)

**When you want the commercial product instead of this one:** if you need to render or convert —
PDF, images, HTML, thumbnails — or need charts, SmartArt, animations, OLE objects or macro-enabled
files, or need a supported product with a licence behind it, none of that is here and none of it is
planned in this repository. This library is the right choice when you are reading and writing
`.pptx` and want an MIT-licensed dependency that brings nothing else with it.

Bugs and feature requests for **this** library belong in
[this repository's issue tracker](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java/issues),
not in the commercial product's forum.

---

## Contributing

Pull requests are welcome. Read [CONTRIBUTING.md](CONTRIBUTING.md) first — it explains the build,
the test suites, and the one rule specific to this project: **a writer fix ships with a test that
asserts on the produced `.pptx` package, not on what the library reads back.**

By participating you agree to the [Code of Conduct](CODE_OF_CONDUCT.md). Changes between releases,
including behaviour changes and removed API, are in [CHANGELOG.md](CHANGELOG.md).

## Security

Do not report a vulnerability in a public issue. [SECURITY.md](SECURITY.md) has the reporting route.

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2026 Aspose Pty Ltd.
