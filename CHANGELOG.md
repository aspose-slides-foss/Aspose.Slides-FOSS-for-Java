# Changelog

All notable changes to this library are recorded here. Versions follow the published
Maven coordinates `org.aspose:aspose-slides-foss`.

## Unreleased

### Build and packaging

- **The published jar declares `Automatic-Module-Name: org.aspose.slides.foss`**, so it can be
  put on the module path under a stable name, and carries `Implementation-Title`,
  `-Version` and `-Vendor` in its manifest along with a copy of the licence at
  `META-INF/LICENSE`.
- **Builds are reproducible.** Two builds of the same source now produce byte-identical
  archives; the build date is fixed by `project.build.outputTimestamp` in `pom.xml` rather
  than taken from the clock.
- **A CycloneDX SBOM is produced** at `target/aspose-slides-foss-<version>-cyclonedx.json`.
  It lists no components, which is the correct answer: the library has no runtime
  dependencies.
- **Java 21 is enforced rather than assumed.** The build fails on a JDK older than 21 with a
  readable message, and compiles with `release` rather than `source`/`target`, so a build on
  a newer JDK cannot produce a jar that quietly needs it.
- **Compiler warnings fail the build** (`-Xlint:all` minus `try`, which would require an API
  change to satisfy), and a javadoc error fails the build instead of producing a javadoc jar
  with pages missing.
- **`PptException`, `PptReadException` and `PptCorruptFileException` declare a
  `serialVersionUID`.** Each is the value the compiler was already computing, so instances
  serialized by an earlier release still deserialize.

### Removed

- **`public static void Picture.flushPendingBlipImages(Element, IBaseSlide)`.** Binary
  incompatible: code compiled against an earlier release that calls it will fail to link.
  It existed to resolve a placeholder relationship id written by `Picture.setImage`; images
  are now embedded through a real relationship at the point they are set, so the method had
  nothing left to resolve and could only ever write an id that nothing declared. There is no
  replacement and none is needed — setting the image is enough.

### Changed — behaviour

- **`save(path, SaveFormat)` and its overloads now raise `UnsupportedOperationException`
  for a format the library cannot write**, instead of writing a PPTX package under the
  requested file name. Writable formats are **PPTX, PPSX and POTX**; every other
  `SaveFormat` value — including `PPTM`, `POTM`, `PPSM`, `PDF`, `ODP`, `HTML`, `XML` and the
  image formats — raises. The exception names the writable set. Nothing is written and no
  file is left behind. This is source-compatible and behaviour-breaking: a call that used to
  return now throws, which is the point, because what it used to return was a mislabelled
  file that PowerPoint refuses to open.
- **`Picture.setImage` raises `IllegalStateException` when the picture is not attached to a
  package part.** An `r:embed` names a relationship in the owning part's `.rels`; with no
  part there is none to declare. The call used to return normally, keep the image in memory,
  write nothing, and remove any reference the blip already carried. Set the image through a
  shape, a fill or a bullet that belongs to a slide, layout or master.
- **`Camera.setCameraType(NOT_DEFINED)`, `LightRig.setLightType(NOT_DEFINED)` and
  `LightRig.setDirection(NOT_DEFINED)` no longer round-trip.** `CT_Camera/@prst` and
  `CT_LightRig/@rig` and `@dir` are required attributes, so there is no way to record "not
  defined" in the file; the getters now return the schema default that was written.
- **Emptying a text frame leaves an empty paragraph in the file.** `CT_TextBody` requires at
  least one `<a:p>`, so clearing the paragraph collection, removing the last paragraph, or
  setting the text to `null` now writes one empty paragraph. The collection itself still
  reports a count of zero.

### Fixed

- **Threaded comments.** A reply was written only to the modern threaded-comment part. A
  reader that renders comments takes the thread from the classic comment list, where the
  reply appeared as an unrelated second comment; the thread is now recorded in both places.
- **Save formats.** POTX and PPSX declared a PPTX content type, so a template saved as a
  template was a presentation.
- **Added and removed slides.** A slide added through the collection was written as a part
  but never registered in `p:sldIdLst`, so it did not exist for a reader; a removed slide
  left its part, its relationship and its content type behind.
- **Embedded images.** A picture referred to a literal relationship id that no `.rels`
  declared, so readers showed an empty box.
- **Picture bullets.** The image set on a bullet was never embedded, and any bullet image
  already there was removed.
- **Effects.** Enabling an outer shadow, inner shadow, glow, preset shadow, soft edge or
  fill overlay wrote an element without the content its type requires, which PowerPoint
  refuses rather than ignores.
- **Pattern fills.** `a:pattFill/@prst` was written as the constant's name rather than the
  `ST_PresetPatternVal` token, so no pattern was recognised.
- **Text runs.** `<a:rPr>` was written after `<a:t>`, which `CT_RegularTextRun` does not
  allow.
- **Paragraphs and portions.** A paragraph added to a text frame, or a portion added to a
  paragraph, went into a private list that nothing serialized.
- **3-D.** A `<a:sp3d>` was written outside the position `CT_ShapeProperties` gives it, and a
  scene was written without the camera and light rig `CT_Scene3D` requires.
- **Tables.** A graphic frame was written without `<a:graphicFrameLocks>`.
- **Masters and layouts.** A loaded presentation reported one synthetic layout instead of the
  masters and layouts the package contains.
- **Saving a subset of slides.** The slide indices were ignored and the whole deck was
  written.
- **Removing notes and comments.** The part was deleted but its relationship and its
  content-type override were left pointing at nothing.
- **`docProps/app.xml`.** The slide, paragraph and word counts were carried over from
  whatever was loaded rather than recomputed, and the word count treated only spaces as
  separators.
- **Re-saving.** An unchanged file grew on every save, because already-indented XML was
  indented again.
