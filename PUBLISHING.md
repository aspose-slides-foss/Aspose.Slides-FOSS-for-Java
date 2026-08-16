# Publishing to Maven Central

This library is published as `org.aspose:aspose-slides-foss` on Maven Central, through the
[Sonatype Central Portal](https://central.sonatype.com/). The whole release runs from
`.github/workflows/maven-central-release.yml`; this file explains what that workflow does, how to
cut a release, and what to do when one goes wrong.

## Cutting a release

1. Bump `<version>` in `pom.xml`, and bump `<project.build.outputTimestamp>` in the same file to
   the release date. The second one is what makes the build reproducible - two builds of the same
   source produce byte-identical jars - so it has to move with the version, not with the clock.
2. Add the release to `CHANGELOG.md`.
3. Commit and push to `main`.
4. Tag the commit with the version, prefixed by `v`, and push the tag:

   ```
   git tag v26.8.0
   git push origin v26.8.0
   ```

The tag push is what starts the release. Pushing a tag whose version does not match `pom.xml`
fails the job immediately, before anything is built.

To rehearse a release without publishing anything, run the workflow from the Actions tab
(`workflow_dispatch`). That runs the same build - full test suite included - and stops before
signing and deploying.

## What the release job does, and what fails it

In order, each step able to fail the whole release:

| Step | Fails when |
|---|---|
| Guard coordinates, match tag to POM | groupId, artifactId or version is wrong, the version is a pre-release, or the tag does not match `pom.xml` |
| Scan for committed key material | a private key block is committed anywhere in the tracked tree |
| Check whether the version is already on Central | Maven Central cannot be reached, or answers something other than 200 or 404 |
| Build, test, sign and deploy | any test fails, javadoc has an error, a compiler warning appears, signing fails, or the Portal does not report the version published |
| Verify the six artifacts | the jar, sources jar, javadoc jar or one of the three signatures is missing |
| Verify the working tree is clean | the build modified or created a tracked-looking file |
| Wait until the files are downloadable | `repo1.maven.org` is still not serving the pom, jar, sources jar, javadoc jar or signature after 28 minutes |
| Create the GitHub Release | the release cannot be created |

Two properties of that list are deliberate and worth keeping:

- **The tests gate the release.** The build command is a single `mvn clean deploy`, and `deploy`
  runs after `test` in the Maven lifecycle, so a failing test stops the publication. Do not add
  `-DskipTests` to it. A build nobody judged is not a release, it is a file transfer.
- **The last check is a real check.** A deploy that the Portal accepts but that never becomes
  downloadable is a failed release, and the job reports it as one. If that check is ever softened
  to a warning, the workflow goes back to reporting success for releases that did not happen.

## Re-running a release that half-failed

Re-run the failed run, or push the tag again. The job asks Maven Central whether the version is
already there:

- **not there** - it builds, signs and deploys as usual;
- **already there** - it skips the build and the deploy, and completes the remaining steps, so a
  release that published but then failed on a later step can be finished without inventing a new
  version number;
- **cannot tell** - it stops, rather than risk publishing on top of a version that may exist.

A version that is already published cannot be replaced: Maven Central is immutable. If something
is wrong with what was published, the only fix is a new version.

## Credentials

Four repository secrets are needed. Without them the release job fails at the deploy step; nothing
else in the repository depends on them.

| Secret | What it is |
|---|---|
| `MAVEN_CENTRAL_USERNAME` | Central Portal user token name |
| `MAVEN_CENTRAL_PASSWORD` | Central Portal user token password |
| `GPG_PRIVATE_KEY` | ASCII-armoured private signing key |
| `GPG_PASSPHRASE` | passphrase for that key |

The signing key is pinned in `pom.xml` as `<keyname>E176D5CBCA1DCC62</keyname>`, and that is the
key that signed the published artifacts - the signature at
`repo1.maven.org/maven2/org/aspose/aspose-slides-foss/26.7.0/aspose-slides-foss-26.7.0.jar.asc`
names it. Replacing the key means changing both the secret and that element.

## Why searching Maven Central does not find this library

Searching for `aspose-slides-foss` on the legacy search endpoint returns nothing at all:

```
$ curl -s 'https://search.maven.org/solrsearch/select?q=g:org.aspose&wt=json'
... "response":{"numFound":0,"start":0,"docs":[]} ...
```

The artifact is nevertheless published, complete and installable:

```
$ curl -sI https://repo1.maven.org/maven2/org/aspose/aspose-slides-foss/26.7.0/aspose-slides-foss-26.7.0.jar
HTTP/1.1 200 OK
$ mvn dependency:get -Dartifact=org.aspose:aspose-slides-foss:26.7.0
[INFO] Downloaded from central: .../aspose-slides-foss-26.7.0.jar (523 kB)
[INFO] BUILD SUCCESS
```

**The cause is where it was published, not whether it was published.** `search.maven.org` serves an
index built by the older OSSRH staging pipeline; this project publishes through the Central Portal
with `central-publishing-maven-plugin`, and components published that way do not enter that index.
Three observations pin it down:

- the legacy index returns `numFound: 0` for the whole `org.aspose` namespace, not merely for this
  artifact, while `repo1.maven.org/maven2/org/aspose/` lists four artifacts;
- it does not contain `org.sonatype.central:central-publishing-maven-plugin` - the Portal's own
  publishing plugin - at all, though it does contain a sibling library from 2023;
- for a project that moved to the Portal mid-life it freezes at the move: it reports
  `org.junit.jupiter:junit-jupiter-api` at 5.13.0-M3 while `repo1.maven.org` serves 6.1.3.

**The remedy is not to re-index; it is to stop relying on that index.**

- Search [central.sonatype.com](https://central.sonatype.com/artifact/org.aspose/aspose-slides-foss)
  instead. It is the Portal's own search, and it returns this component with the current version.
- Dependency *resolution* is unaffected. Maven, Gradle, sbt, Bazel and every IDE that resolves
  against `repo1.maven.org` download the artifact normally; only search interfaces built on the
  legacy index are blind to it.
- Because of that, the coordinates are given verbatim in `README.md`, so nobody has to search for
  them in the first place.

Do not try to force a re-index, and do not republish to work around it: the artifact is correct
where it matters, and republishing the same version is not possible anyway.

## Verifying a published release by hand

```bash
V=26.7.0
B=https://repo1.maven.org/maven2/org/aspose/aspose-slides-foss/$V/aspose-slides-foss-$V
for s in .pom .jar -sources.jar -javadoc.jar .jar.asc; do
  curl -sI "$B$s" | head -1
done
curl -sO "$B.jar" -O "$B.jar.asc"
gpg --verify "aspose-slides-foss-$V.jar.asc" "aspose-slides-foss-$V.jar"
```

## Publication history

| Version | Published | Coordinates |
|---|---|---|
| 26.7.0 | 2026-07-27 | `org.aspose:aspose-slides-foss:26.7.0` |
