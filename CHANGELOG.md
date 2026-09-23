# Changelog

All notable changes to this library are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the versions are the published Maven
coordinates `org.aspose:aspose-slides-foss`.

## Unreleased

A change a caller can observe goes here, under a `### Added`, `### Changed`,
`### Fixed`, `### Removed` or `### Security` heading, and the section is renamed to the version and
dated when that version is released.

### Fixed

- **Speaker notes are wired to a notes master and to their slide.** `addNotesSlide()` wrote the
  notes slide as a bare part with no relationships of its own: nothing led from it back to its
  slide, and the presentation had no notes master, although ECMA-376 requires every notes slide
  to have one. PowerPoint still showed the text, but a reader that follows the relationships
  found notes that belong to nothing, and the notes placeholders had no master to inherit their
  formatting from. A notes slide is now related to its slide and to the presentation's notes
  master. If the presentation has no notes master, a minimal one is added with its own copy of
  the presentation's theme, a content-type override and a `p:notesMasterIdLst` entry; a deck that
  already has one keeps it. What a caller can see: the first notes added to a deck add
  `ppt/notesMasters/notesMaster1.xml` and one more theme part to the saved file, and the
  relationship from the slide to its notes now has an ordinary `rIdN` id instead of `rId_notes`.
  Covered by `NotesMasterConformanceTest`.
- **Notes and comments are found through the slide's relationships, not by part number.** The
  library paired `notesSlideN.xml` and `commentN.xml` with `slideN.xml` by their numbers. Other
  producers number these parts independently — a deck whose first slide has no notes commonly
  keeps the second slide's notes in `notesSlide1.xml` — and on such a deck the library reported
  one slide's notes or comments as another's, `addNotesSlide()` wrote into another slide's notes,
  and removing notes or a comment deleted another slide's part and left that slide pointing at a
  part that was gone, which PowerPoint refuses to open. Saving after cloning a slide whose
  comments had been loaded from the file could break the file the same way. Each slide's notes,
  comments and threaded-comment parts are now the ones its relationships name. A new part still
  takes the slide's number when that name is free, and the lowest free number otherwise. A cloned
  slide is still saved without the source slide's comments. Covered by
  `PartNumberingConformanceTest`.

## [26.8.0] - 2026-08-16

`26.7.0` is still the version on Maven Central as this is written; `26.8.0` is the version in the
source tree and is published when the release tag is pushed.

### Upgrading from 26.7.0 — read this first

Three changes will be noticed by code that already works, and all three are deliberate:

1. **`save` now refuses a format it cannot write** instead of writing a PPTX package under the name
   you asked for. If your code saves as `PPTM`, `PDF`, `ODP` or anything else outside PPTX, PPSX and
   POTX, it now throws where it used to return — and what it used to return was a mislabelled file
   PowerPoint refuses to open.
2. **`Picture.setImage` now throws when the picture is not attached to a package part**, where it
   used to return normally and write nothing.
3. **`Picture.flushPendingBlipImages` has been removed**, which is binary incompatible: recompile.
   Nothing replaces it, and nothing needs to.

Each is described in full below. A `.pptx` written by `26.7.0` still loads.

### Security

- **XML parsing no longer resolves anything outside the document being parsed.** Every parser and
  serializer in the library is now created by `org.aspose.slides.foss.internal.xml.SecureXml`, which
  refuses a `<!DOCTYPE>` declaration outright and disables external general and parameter entities,
  external DTD loading and XInclude. Before this change, a part of a `.pptx` could declare an
  external entity and the parser would resolve it while loading the file — so opening a presentation
  from an untrusted source could read a local file, or open a network connection, and place the
  result in the presentation where the calling application would treat it as slide text. Office Open
  XML packages carry no DTD, so nothing legitimate is lost. Opening a package whose parts declare a
  `DOCTYPE` now fails with a parse error. Covered by
  `tests/conformance/.../UntrustedInputConformanceTest.java`.

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
- **The test suites outside `src/test/java` are checked to be on the test classpath at all.** The
  integration and conformance suites live outside Maven's default test root and are added by
  `build-helper-maven-plugin`; a root missing from that configuration compiles nothing, runs
  nothing and reports success. A test in `src/test/java` — the one root that cannot itself go
  missing — now fails the build and names every class that was not compiled.
- **Every push and pull request builds and runs the whole test suite** on Linux, Windows and macOS,
  on Java 21 and Java 25, and a separate job checks that two builds of the same source produce
  identical jars. Before this, nothing ran the tests outside a release.
- **The release workflow runs the test suite and waits for the artifacts to be downloadable.** It
  used to build with `-DskipTests` and to treat "not on Maven Central yet" as a warning, so a
  release that published nothing could still report success. It now fails instead, and a release
  that published but failed a later step can be re-run without inventing a new version number.

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

## [26.7.0] - 2026-07-27

The first published release.

[26.8.0]: https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java/compare/v26.7.0...v26.8.0
[26.7.0]: https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java/releases/tag/v26.7.0
