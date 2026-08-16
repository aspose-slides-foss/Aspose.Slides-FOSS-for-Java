# Contributing

Thank you for considering a contribution. This document is specific to the **Java** edition of
Aspose.Slides FOSS; the .NET, C++ and Python editions live in their own repositories and have their
own build commands and conventions.

## The one rule that is particular to this project

**A fix to a writer ships with a test that asserts on the produced `.pptx` package — not on what
the library reads back.**

A library agrees with itself for free. If a writer emits an element under the wrong name and the
matching reader looks for that same wrong name, then every property reads back exactly what was
set, every object-model test passes, and the file is still one PowerPoint quietly strips on load.
The same shape of mistake hides a relationship id that is never written into the `.rels`, a part
with no content type, and an effect element missing a required child: the object model is intact in
memory, so a test built on the object model sees nothing at all.

This repository has a concrete example. The loader used to count `ppt/slides/slideN.xml` entries in
the ZIP — but a slide part is not a slide. A reader finds slides by walking `<p:sldIdLst>` in
`ppt/presentation.xml` and resolving each `r:id`. A part nothing references is invisible to
PowerPoint, to Apache POI and to everything else, and counting parts made a deck that had silently
lost two of three slides look complete.

So: object-model tests are welcome and there are 3,306 of them, but they cannot close a writer bug.
Put the assertion in `tests/conformance/`, where it unzips the file and reads the XML.
[`tests/conformance/README.md`](tests/conformance/README.md) describes the harness in full.

## Prerequisites

**JDK 21 or later and Maven 3.9 or later.** The build enforces both and stops with a readable
message rather than a page of `cannot find symbol` if either is older. Java 21 is a hard floor, not
a preference: the sources call `List.getFirst()` and `List.removeLast()` and use an unconditional
`instanceof` pattern.

The library itself has **no dependencies**. Everything in `pom.xml` is `<scope>test</scope>`:
JUnit 5, AssertJ, and Apache POI as the independent reader the conformance tests use.

## Build

```bash
git clone https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java.git
cd Aspose.Slides-FOSS-for-Java
mvn verify -Dgpg.skip=true
```

`-Dgpg.skip=true` skips artifact signing, which is bound to the `verify` phase and needs the release
key. Four things about this build are worth knowing before your first pull request:

- **Warnings are build failures.** The compiler runs with `-Xlint:all,-try` and
  `failOnWarning=true`. The tree compiles clean under that set, main sources and tests alike, so a
  new warning is a regression and is treated as one. If a warning genuinely has to be allowed,
  suppress it at the declaration with a comment saying why — do not weaken the flag. (`try` is the
  one category left off, because `IImage` and `IPresentation` extend `AutoCloseable` without
  narrowing its `close() throws Exception`; narrowing it is a published-API change and belongs in a
  release with a note, not in a build tweak.)
- **A javadoc error is a build failure too.** `doclint` runs `all,-missing`: broken HTML, an
  unclosed tag or a `@link` to something that does not exist fails the build, while an undocumented
  `@param` stays a warning. New public members should still be documented — the javadoc jar is
  published and it is what a consumer sees in their IDE.
- **The build is reproducible.** `project.build.outputTimestamp` in `pom.xml` fixes every timestamp
  Maven writes into the archives, so two builds of the same source produce byte-identical jars. CI
  checks it by building twice and comparing SHA-256 hashes. If you change the version, move that
  timestamp with it.
- **`verify` is the useful target, not `package`.** It additionally builds the sources jar, the
  javadoc jar and the CycloneDX SBOM, so a javadoc or packaging error is found by you rather than
  by a release.

## Test

```bash
mvn test
```

Three suites run, and all three must be green:

| Source root | Tests | What it covers |
|---|---|---|
| `src/test/java` | 3,306 | unit tests of the object model |
| `tests/integration` | 98 | end-to-end use of the public API |
| `tests/conformance` | 232 | assertions against the produced `.pptx` package |

`tests/integration` and `tests/conformance` are outside `src/test/java` and are added as test source
roots by `build-helper-maven-plugin` in `pom.xml`. A new directory there needs a line in that
plugin's configuration or it is silently never compiled.

Run a subset with surefire's `-Dtest`:

```bash
mvn test -Dtest='*ConformanceTest,HarnessCalibrationTest'   # the conformance suite
mvn test -Dtest=SaveFormatConformanceTest                   # one class
mvn test -Dtest='SaveFormatConformanceTest#*Potx*'          # one method pattern
```

Run the whole suite before you push, rather than only the class you touched. Naming classes one at a
time means a suite added later is silently never run.

### The conformance suite

[`tests/conformance/README.md`](tests/conformance/README.md) is the full description. The short
version is that the harness gives you six types:

| Type | What it gives you |
|---|---|
| `PptxPackage` | opens a produced file as a ZIP: entry names, part bytes, part text, part DOM, and the two OPC path rules |
| `PackageAssertions` | the four package-wide rules, the registered slide count read from `<p:sldIdLst>`, XPath selection, attribute and schema child-order assertions |
| `ThirdPartyReadBack` | Apache POI as an independent reader — and note that POI *loading* a file proves less than it looks, so prefer `assertPicturesResolve` over `assertOpens` |
| `Fixtures` | decks written by Apache POI rather than by this library, with a real master, eleven layouts and a theme |
| `ZipSurgery` | copies a package with one deliberate edit, used to calibrate the rules |
| `HarnessCalibrationTest` | proves each rule passes on a good package and fails on a damaged one |

Apache POI is referenced by the test sources **only**. It must never become a compile-scope
dependency: this library writes OOXML by hand, and "no runtime dependencies" is a property it
advertises and the SBOM proves.

## What a good pull request looks like

1. **It fixes one thing.** A pull request that repairs a writer and also renames three files is two
   pull requests.
2. **It has a failing test first.** Write the test against the unfixed code, watch it fail, and put
   that failure message in the pull request description. A test that has never failed has not been
   shown to test anything.
3. **The test name is the bug report**, in a full sentence and in the words a user would use:
   `anAddedSlideMustBeRegisteredInTheSavedFile`, not `testAddSlide3`.
4. **The assertion says what was actually written when it fails.** Every assertion in the
   conformance suite carries an AssertJ `.as(...)` naming the file and, where it is short enough,
   quoting the XML. `expected: 2 but was: 1` costs the next person twenty minutes.
5. **It writes its file through the public API only.** If a case needs a package-private member to
   set up, it is testing the internals rather than the artefact. If the public API cannot express
   the case, that is itself the finding.
6. **It says what changed about the file.** "Writes `<a:rPr>` before `<a:t>`, as
   `CT_RegularTextRun` requires" is reviewable. "Fixed text runs" is not.
7. **It updates [CHANGELOG.md](CHANGELOG.md)** under `## Unreleased`, in the language a caller would
   use, whenever a caller can observe the change.
8. **It adds no compile-scope dependency to `pom.xml`.** If you believe one is genuinely necessary,
   open an issue and make the case before writing the code.
9. **It does not import from `org.aspose.slides.foss.internal`** outside that package. That is
   implementation detail and is not API.

### Behaviour that is deliberate, not a bug

Before filing a fix for one of these, please open an issue instead — they are decisions with
reasoning behind them, and the reasoning is in the source and in `CHANGELOG.md`:

- `save` raises `UnsupportedOperationException` for the 22 `SaveFormat` values it does not write,
  and leaves no file behind. It will not write a presentation package under a name claiming to be
  PDF, ODP or a macro-enabled file.
- `Picture.setImage` raises `IllegalStateException` when the picture is not attached to a package
  part. An `r:embed` names a relationship in the owning part's `.rels`; with no part there is none
  to declare.
- `Camera.setCameraType(NOT_DEFINED)` and the two `LightRig` equivalents do not round-trip.
  `CT_Camera/@prst` and `CT_LightRig/@rig` and `@dir` are required attributes, so "not defined" has
  nowhere to be recorded, and the getter returns the schema default that was written.
- Emptying a text frame leaves one empty `<a:p>` in the file, because `CT_TextBody` requires at
  least one. The collection still reports a count of zero.

Genuine known defects are listed in the README under **Limitations**; `getMasters().addClone(...)`
is the current one, and a pull request fixing it is very welcome.

## Commits and style

- Follow the layout and naming of the code you are changing. There is no separate style guide and no
  static-analysis configuration beyond warnings-as-errors.
- Write the commit subject as a short sentence saying what the change does to the software; the
  existing history is the model.
- Sign nothing off — there is no CLA and no DCO check.

## Continuous integration

Every push and every pull request builds and runs all three suites on `ubuntu-latest`,
`windows-latest` and `macos-latest`, on Java 21 and Java 25, and a separate job builds the project
twice and fails if the two jars differ. Test reports are attached to every run; the jar, sources
jar, javadoc jar and SBOM are attached to the Linux/Java 21 run. If CI is red the pull request is
not ready, including when the failure is on a platform or a JDK you did not use.

Warnings-as-errors is applied on Java 21 only. `-Xlint:all` is open-ended — a later javac can add a
category — so a tree that was clean yesterday would otherwise fail today for a reason that has
nothing to do with the change under test. The warnings are still printed on the other JDK.

## Licence

By contributing you agree that your contribution is licensed under the [MIT License](LICENSE), the
same terms as the rest of the project.
