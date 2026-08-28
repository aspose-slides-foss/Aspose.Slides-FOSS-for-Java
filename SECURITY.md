# Security policy

## Supported versions

Maven Central is immutable: a published version can never be replaced, so every fix ships as a new
version rather than as a correction to an old one.

| Version | Supported |
|---|---|
| `26.8.0` and later | yes — a fix goes into the next version. `26.8.0` is the version in this source tree and is not published yet. |
| `26.7.0` | no. It is the only version published so far, and it predates a large body of correctness and hardening work. Upgrade when `26.8.0` is out, or build from source, rather than asking for a patch release. |
| Anything built from an arbitrary commit | no — build from a release tag or from the default branch |

If you are pinning `26.7.0` today, read [CHANGELOG.md](CHANGELOG.md) before deciding to stay on it.
Among other things, the XML parsing this library does was hardened against hostile input in
`26.8.0`, and the entry there says exactly what changed.

## Reporting a vulnerability

**Do not open a public issue for a security problem, and do not attach a proof-of-concept file to
one.**

Use GitHub's private vulnerability reporting on this repository:
[**Report a vulnerability**](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java/security/advisories/new).
It opens a private advisory only the maintainers can read, and it lets you attach files and discuss
a fix before anything becomes public.

If that page is not available to you, open a public issue containing **only** the sentence "I would
like to report a security issue privately" and no details at all, and wait to be contacted.

Please include, as far as you can:

- the version or the commit you tested — `mvn -q -DforceStdout help:evaluate -Dexpression=project.version`
  prints the version, `git rev-parse HEAD` the commit;
- the JDK and the operating system (`java -version`, abbreviated, is ideal);
- a minimal program, and the `.pptx` it needs — attach the file, or say how to build it;
- what happens, and what you expected instead;
- the impact you believe it has.

You will get an acknowledgement. We cannot promise a fix deadline for a project with no paid support
contract behind it, but you will be told what is happening and when a fix lands, and you will be
credited in the advisory unless you ask not to be.

## What is in scope

This library parses untrusted input by design. A `.pptx` is a ZIP archive full of XML, and when the
file came from outside, every byte of both layers was chosen by whoever sent it. Reports about the
handling of a malicious or malformed presentation are in scope, including:

- **XML that reaches outside itself** — a `DOCTYPE`, an external entity, an XInclude, or anything
  else that makes the parser read a local file or open a network connection. Since `26.8.0` every
  parser is created by `org.aspose.slides.foss.internal.xml.SecureXml`, which rejects a `DOCTYPE`
  outright; a route that gets round it is a vulnerability.
- **A part name or relationship target that escapes the package** and causes a read or a write
  outside the directory the caller named.
- **Unbounded resource use** from a crafted archive or crafted XML — entity expansion, a
  compression ratio chosen to exhaust memory, a structure that does not terminate.
- **A crash, an unbounded allocation or an infinite loop** on a malformed file, where the caller
  cannot defend itself by catching an exception.
- Anything else that lets a presentation influence the process beyond the object model it is parsed
  into.

A useful report says which part of the package carried the payload and what the process did.

## What is out of scope

- **Missing capabilities.** `save` raising `UnsupportedOperationException` for a format it does not
  write, and everything under **Limitations** in the [README](README.md), is documented behaviour
  rather than a vulnerability.
- **Vulnerabilities in the JDK itself.** Report those to your JDK vendor. This library calls
  `javax.xml` and `java.util.zip` from the platform and has no XML or ZIP implementation of its own.
- **Apache POI.** It is a `test` scope dependency used by the conformance suite as an independent
  reader; it is not part of the published artifact and no consumer receives it. Report POI issues to
  [Apache](https://poi.apache.org/security.html).
- **The commercial Aspose.Slides product**, which is different software. Report those through
  [Aspose support](https://forum.aspose.com/c/slides/11).
- **Findings from an automated scanner with no demonstrated impact on this library.** A CVE listed
  against a transitive test-scope dependency is not a vulnerability in the jar you install; the
  CycloneDX SBOM the build produces lists what a consumer actually receives, and it is empty.
