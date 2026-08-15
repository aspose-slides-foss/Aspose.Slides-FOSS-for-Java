# Conformance tests

These tests answer one question the rest of the suite cannot: **is the file we just wrote a
PowerPoint file?**

They write a `.pptx` through the public API, unzip it, parse the XML inside, and assert on the
elements, the attributes and the relationships. Then they hand the same file to Apache POI — a
reader written by other people — and assert on what *it* sees.

## Why not just read the file back with this library?

Because that proves nothing.

A test that writes a value with this library and reads it back with this library passes whenever
the writer and the reader are wrong in the same way — which is the normal case, since they are
usually the same person's work on the same day. The rest of the suite is written that way, and it
is green on files that PowerPoint refuses to open.

The concrete example is the slide count. This library's loader has historically counted
`ppt/slides/slideN.xml` entries in the ZIP. But a slide part is not a slide: a reader finds slides
by walking `<p:sldIdLst>` in `ppt/presentation.xml` and resolving each `r:id` through
`ppt/_rels/presentation.xml.rels`. A part that no `p:sldId` references, that no relationship
points at and that no `[Content_Types].xml` override types is invisible to PowerPoint, to POI and
to everything else. Count parts and a deck that has silently lost two of its three slides looks
complete; read `sldIdLst` and it does not.

So: **assert on the package, never on the library's own read-back.** `PackageAssertions`
deliberately offers no route back into the library, and `PackageAssertions.registeredSlideCount`
reads `sldIdLst`.

The one deliberate exception is `LoadedDeckConformanceTest`, where the *reader* is what is being
tested. There the expected values are still taken from the package — counted out of the ZIP —
and only then compared with what the library reports.

## What is here

| File | What it gives you |
|---|---|
| `PptxPackage` | opens a produced file as a ZIP: entry names, part bytes, part text, part DOM, and the two OPC path rules (`_rels` location, relative target resolution) |
| `PackageAssertions` | the package-wide rules, the slide count, XPath selection, attribute assertions and schema child-order assertions |
| `ThirdPartyReadBack` | Apache POI as an independent reader: slide count, shape count, text, and whether a picture's image data actually resolves |
| `Fixtures` | decks written by something other than this library, with a real master, eleven layouts and a theme |
| `ZipSurgery` | copies a package with one deliberate edit, used only to calibrate the rules |
| `HarnessCalibrationTest` | proves each rule passes on a good package and fails on a damaged one |

The four package-wide rules, all bundled in `assertPackageIsSelfConsistent`:

1. every `r:id` / `r:embed` / `r:link` resolves to a `Relationship Id` in that part's own `.rels`;
2. every internal relationship target names a part that exists;
3. every part resolves a content type through an `Override` or a `Default`;
4. no `Override` names a part that is missing.

Rules 2 and 4 are the two halves of a part deletion. Deleting a part means removing the bytes,
the part's own `.rels`, the `<Relationship>` in the owner's `.rels` **and** the `<Override>` —
doing only the first leaves a package that advertises a part it does not contain, which strict
readers reject outright.

## Apache POI

POI (`poi-ooxml`, Apache-2.0) is a **test-scope** dependency. It is never used by main sources and
is not a dependency of the published artifact.

It is here because a file that only PowerPoint accepts is not good enough: server-side pipelines,
document management systems and search indexers are built on libraries like POI, and a deck POI
cannot open is a deck that silently fails in production. POI is also permissive, so it makes a
useful lower bound — if POI complains, everything stricter will too.

Note that POI *loading* a file proves less than it looks. It will happily load a slide whose
picture relationship is dangling; only asking for the image data shows there is nothing behind it.
Prefer `assertPicturesResolve` over `assertOpens`.

## Running them

```
mvn test -Dtest='*ConformanceTest,HarnessCalibrationTest'
```

Or just `mvn test`, which runs them with everything else.

## Adding a case

1. **Name the test after the user-visible failure**, in a full sentence — 
   `anAddedSlideMustBeRegisteredInTheSavedFile`, not `testAddSlide`. Someone reading a red build
   should learn what a user would have lost, without opening the file.
2. **Write the file through the public API only.** If the API cannot express it, that is itself
   the finding.
3. **Assert through `PptxPackage` and `PackageAssertions`.** If you find yourself calling
   `new Presentation(output)` to check the result, stop — see above.
4. **Say what was actually written when it fails.** Every assertion here carries an `.as(...)`
   describing the file and, where it is short enough, quoting the XML. A failure message of
   `expected: 2 but was: 1` costs the next person twenty minutes.
5. **Run it against the unfixed code first and watch it fail.** A conformance test that passes
   the moment it is written is either testing the wrong thing or the defect is not real. Record
   which.
6. **If the case needs a realistic input deck, add it to `Fixtures`** rather than committing
   binary files. The fixtures are generated, so they are diffable, and they are demonstrably not
   this library's own output.
