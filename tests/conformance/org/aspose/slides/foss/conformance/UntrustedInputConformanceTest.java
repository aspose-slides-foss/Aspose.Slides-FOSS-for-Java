package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.IAutoShape;
import org.aspose.slides.foss.IShape;
import org.aspose.slides.foss.ISlide;
import org.aspose.slides.foss.Presentation;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * A {@code .pptx} that came from somewhere else is attacker-controlled in both of its layers, and
 * these tests are about the XML one.
 *
 * <p>An XML parser left at its defaults resolves a {@code <!DOCTYPE>} and the external entities
 * declared in it. A presentation part containing one line of DTD then makes the parser open a
 * local file, or a URL, and hand the contents back as slide text — so an application that opens an
 * uploaded deck and prints its text prints whatever the uploader named. Nothing in Office Open XML
 * needs a DTD, so this library refuses one outright; see
 * {@code org.aspose.slides.foss.internal.xml.SecureXml}.</p>
 *
 * <p>These tests are here rather than in the unit suite because they are about what the library
 * does with a package, and because the package they need has to be built by hand: no supported API
 * writes a DOCTYPE, which is the point.</p>
 */
class UntrustedInputConformanceTest {

    /** Text that must never reach the object model, no matter what the file asks for. */
    private static final String SECRET = "conformance-secret-must-not-be-disclosed";

    /** The slide text the entity reference replaces, so that the entity is really used. */
    private static final String PLACEHOLDER = "PLACEHOLDER";

    @TempDir
    Path tempDir;

    /**
     * Builds a deck whose first slide part declares an external entity pointing at {@code target}
     * and references it where the slide's text would be.
     */
    private Path deckWithExternalEntity(Path target) throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "source.pptx", PLACEHOLDER);
        Path poisoned = tempDir.resolve("poisoned.pptx");
        String uri = target.toUri().toString();
        Path result = ZipSurgery.copyWithReplacement(source, poisoned, "ppt/slides/slide1.xml",
                xml -> {
                    String doctype =
                            "<!DOCTYPE p:sld [<!ENTITY disclose SYSTEM \"" + uri + "\">]>";
                    int afterDeclaration = xml.indexOf("?>") + 2;
                    // The entity has to be referenced as well as declared, or the parser has no
                    // reason to fetch anything and the test would pass for the wrong reason.
                    return xml.substring(0, afterDeclaration) + doctype
                            + xml.substring(afterDeclaration).replace(PLACEHOLDER, "&disclose;");
                });
        try (PptxPackage pkg = PptxPackage.open(result)) {
            assertThat(pkg.text("ppt/slides/slide1.xml"))
                    .as("the poisoned package must declare the entity and reference it, or the "
                            + "parser has no reason to fetch anything and the test would pass for "
                            + "the wrong reason")
                    .contains("<!DOCTYPE")
                    .contains("&disclose;");
        }
        return result;
    }

    /** A part that declares a DOCTYPE is not a presentation part, and is not parsed. */
    @Test
    void aPartThatDeclaresADoctypeMustBeRefused() throws Exception {
        Path secret = tempDir.resolve("secret.txt");
        Files.writeString(secret, SECRET, StandardCharsets.UTF_8);
        Path poisoned = deckWithExternalEntity(secret);

        assertThatThrownBy(() -> {
            try (var pres = new Presentation(poisoned.toString())) {
                pres.getSlides().size();
            }
        }).as("opening %s, whose ppt/slides/slide1.xml declares an external entity", poisoned)
                .isInstanceOf(RuntimeException.class);
    }

    /**
     * And whatever the library decides to do with such a file, the contents of the named file must
     * never appear in the presentation.
     *
     * <p>Asserted separately from the refusal above so that a future decision to skip the DOCTYPE
     * rather than reject the part cannot quietly reintroduce the disclosure.</p>
     */
    @Test
    void anExternalEntityMustNotDiscloseTheFileItNames() throws Exception {
        Path secret = tempDir.resolve("secret.txt");
        Files.writeString(secret, SECRET, StandardCharsets.UTF_8);
        Path poisoned = deckWithExternalEntity(secret);

        StringBuilder everythingRead = new StringBuilder();
        try (var pres = new Presentation(poisoned.toString())) {
            for (int s = 0; s < pres.getSlides().size(); s++) {
                ISlide slide = pres.getSlides().get(s);
                for (int i = 0; i < slide.getShapes().size(); i++) {
                    IShape shape = slide.getShapes().get(i);
                    if (shape instanceof IAutoShape autoShape && autoShape.getTextFrame() != null) {
                        everythingRead.append(autoShape.getTextFrame().getText());
                    }
                }
            }
        } catch (RuntimeException expected) {
            everythingRead.append(String.valueOf(expected.getMessage()));
        }

        assertThat(everythingRead.toString())
                .as("everything %s yielded; it declared an external entity naming a file whose "
                        + "contents are %s", poisoned, SECRET)
                .doesNotContain(SECRET);
    }

    /**
     * The same refusal covers entity expansion, which needs no external reference at all: a
     * handful of nested internal entities expand to gigabytes and exhaust the heap before any
     * limit on document size applies.
     */
    @Test
    void anExpandingInternalEntityMustBeRefused() throws Exception {
        Path source = Fixtures.authoredDeck(tempDir, "source.pptx", "First");
        Path poisoned = tempDir.resolve("expanding.pptx");
        ZipSurgery.copyWithReplacement(source, poisoned, "ppt/slides/slide1.xml", xml -> {
            String doctype = "<!DOCTYPE p:sld ["
                    + "<!ENTITY a \"aaaaaaaaaa\">"
                    + "<!ENTITY b \"&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;\">"
                    + "<!ENTITY c \"&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;\">"
                    + "]>";
            int afterDeclaration = xml.indexOf("?>") + 2;
            return xml.substring(0, afterDeclaration) + doctype + xml.substring(afterDeclaration);
        });

        assertThatThrownBy(() -> {
            try (var pres = new Presentation(poisoned.toString())) {
                pres.getSlides().size();
            }
        }).as("opening %s, whose ppt/slides/slide1.xml declares nested internal entities", poisoned)
                .isInstanceOf(RuntimeException.class);
    }
}
