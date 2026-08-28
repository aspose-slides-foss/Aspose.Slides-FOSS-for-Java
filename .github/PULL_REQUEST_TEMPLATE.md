<!--
Thank you for the pull request. CONTRIBUTING.md has the detail; this template is the short form.
Delete any section that genuinely does not apply, rather than leaving it blank.
-->

## What this changes

<!-- One or two sentences. If it changes what ends up in the .pptx, name the part and the element:
     "writes <a:rPr> before <a:t> in ppt/slides/slide1.xml, as CT_RegularTextRun requires". -->

Closes #

## Why

<!-- The user-visible problem. What did someone see happen, and what did they expect? -->

## How it was verified

<!-- Paste the failing test output from BEFORE the fix, and the passing run after it. A test that
     has never failed has not been shown to test anything. -->

```
```

## Checklist

- [ ] `mvn verify -Dgpg.skip=true` is green on JDK 21. Warnings are errors and a javadoc error
      fails the build, so this is pass or fail rather than a judgement call.
- [ ] All three suites pass — `src/test/java`, `tests/integration` and `tests/conformance`. Run the
      whole thing, not only the class you touched.
- [ ] **If this changes what is written to the file**, there is a test in `tests/conformance` that
      unzips the produced `.pptx` and asserts on its XML. A test that reads the value back through
      this library does not count.
- [ ] The test is named after the user-visible failure, in a full sentence, and its assertion
      carries an AssertJ `.as(...)` that says what was actually written when it fails.
- [ ] New public members carry javadoc. (`doclint` only fails on broken javadoc, not on missing
      javadoc, so this one is on you rather than on the compiler.)
- [ ] `CHANGELOG.md` is updated under `## Unreleased` if a caller can observe this change, in the
      words a caller would use.
- [ ] No new compile-scope dependency in `pom.xml`. "No runtime dependencies" is a property this
      library advertises and the SBOM proves.
- [ ] Nothing imports from `org.aspose.slides.foss.internal` outside that package.
- [ ] No build output, IDE files or test artefacts in the diff (`git status` before you push).

## Anything a reviewer should look at closely

<!-- A decision you were unsure about, a case you did not cover, a behaviour you changed on
     purpose. Say it here rather than letting it be found. -->
