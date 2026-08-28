package org.aspose.slides.foss.conformance;

import org.aspose.slides.foss.Presentation;
import org.aspose.slides.foss.export.SaveFormat;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.catchThrowable;

/**
 * A file must be the format its name and its content type claim it is.
 *
 * <p>Asking for a format the library cannot write and getting a PowerPoint package under that
 * name is a lie the caller has no way to detect: the save call returns normally, the file
 * exists, and it is only the recipient who finds out. PowerPoint itself refuses several of
 * these names outright — "PowerPoint can't open this file because its file extension has
 * changed" — so the user's first signal is a support ticket.</p>
 */
class SaveFormatConformanceTest {

    /**
     * The formats that share the PPTX package shape, and the content type
     * {@code /ppt/presentation.xml} must carry for each of them.
     */
    private static final Map<SaveFormat, String> MAIN_PART_CONTENT_TYPES = Map.of(
            SaveFormat.PPTX,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            SaveFormat.POTX,
            "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
            SaveFormat.PPSX,
            "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml");

    @TempDir
    Path tempDir;

    /**
     * Every format either produces that format or refuses; none produces a different one.
     *
     * <p>The macro-enabled names (PPTM, POTM, PPSM) are in the refusing group: their package
     * shape is the same, but a file that claims to be macro-enabled and carries no VBA part is
     * another kind of mislabelling.</p>
     *
     * @param format the requested save format
     */
    @ParameterizedTest(name = "{0}")
    @EnumSource(SaveFormat.class)
    void savingInAFormatMustProduceThatFormatOrRefuse(SaveFormat format) throws Exception {
        String expectedContentType = MAIN_PART_CONTENT_TYPES.get(format);
        Path out = tempDir.resolve("deck." + format.name().toLowerCase(Locale.ROOT));

        try (var pres = new Presentation()) {
            if (expectedContentType == null) {
                Throwable thrown = catchThrowable(() -> pres.save(out.toString(), format));
                assertThat(thrown)
                        .as("saving as %s is not implemented and must say so, not write a "
                                + "PowerPoint package under a .%s name", format,
                                format.name().toLowerCase(Locale.ROOT))
                        .isInstanceOf(UnsupportedOperationException.class);
                assertThat(Files.exists(out))
                        .as("a refused save must not leave a file behind at %s", out)
                        .isFalse();
            } else {
                pres.save(out.toString(), format);
                try (PptxPackage pkg = PptxPackage.open(out)) {
                    assertThat(PackageAssertions.mainPartContentType(pkg))
                            .as("content type of /ppt/presentation.xml in %s", out)
                            .isEqualTo(expectedContentType);
                }
            }
        }
    }

    /** The formats that are written must not all be the same file under different names. */
    @Test
    void theFormatsThatAreWrittenMustDifferFromEachOther() throws Exception {
        Set<String> contentTypes = new LinkedHashSet<>();
        for (SaveFormat format : MAIN_PART_CONTENT_TYPES.keySet()) {
            Path out = tempDir.resolve("distinct." + format.name().toLowerCase(Locale.ROOT));
            try (var pres = new Presentation()) {
                pres.save(out.toString(), format);
            }
            try (PptxPackage pkg = PptxPackage.open(out)) {
                contentTypes.add(PackageAssertions.mainPartContentType(pkg));
            }
        }
        assertThat(contentTypes)
                .as("each written format must declare its own main-part content type")
                .hasSize(MAIN_PART_CONTENT_TYPES.size());
    }
}
