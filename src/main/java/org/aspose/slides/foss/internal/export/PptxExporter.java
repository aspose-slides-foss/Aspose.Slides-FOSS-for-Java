package org.aspose.slides.foss.internal.export;

import org.aspose.slides.foss.export.ISaveOptions;
import org.aspose.slides.foss.internal.pptx.ContentTypesManager;
import org.aspose.slides.foss.internal.pptx.OpcPackage;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Exporter for PPTX and related Office Open XML formats.
 *
 * <p>Supports:</p>
 * <ul>
 *   <li>PPTX — Standard PowerPoint presentation</li>
 *   <li>PPSX — PowerPoint show (opens in slideshow mode)</li>
 *   <li>POTX — PowerPoint template</li>
 * </ul>
 *
 * <p>These formats are all OPC packages that differ only in the content type of the
 * main presentation part, so the same serializer writes all three.</p>
 *
 * <p>The macro-enabled names (PPTM, PPSM, POTM) are deliberately absent. Their package
 * shape is the same, but their content types declare a VBA project that this library
 * does not write; a file claiming to be macro-enabled with no {@code ppt/vbaProject.bin}
 * is mislabelled just as surely as a PPTX named {@code .pdf}. They are rejected until
 * VBA parts are carried through.</p>
 */
public final class PptxExporter extends ExporterBase {

    /**
     * Mapping from SaveFormat values to main presentation content types.
     *
     * <p>Insertion-ordered: the key order is what a caller is shown when it asks for a
     * format that cannot be written, and an unordered map would list them differently
     * from one run to the next.</p>
     */
    private static final Map<String, String> CONTENT_TYPES = contentTypes();

    private static Map<String, String> contentTypes() {
        var types = new LinkedHashMap<String, String>();
        types.put("Pptx",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml");
        types.put("Ppsx",
                "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml");
        types.put("Potx",
                "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml");
        return Collections.unmodifiableMap(types);
    }

    /** The part whose content type identifies the format of the whole package. */
    private static final String MAIN_PART_NAME = "/ppt/presentation.xml";

    private final String targetFormat;

    /**
     * Initialize the PPTX exporter with the default format ({@code Pptx}).
     */
    public PptxExporter() {
        this("Pptx");
    }

    /**
     * Initialize the PPTX exporter for a specific target format.
     *
     * @param targetFormat the specific format to export to (e.g., "Pptx", "Potx")
     */
    public PptxExporter(String targetFormat) {
        this.targetFormat = targetFormat;
    }

    /**
     * Export the presentation to a PPTX file.
     *
     * @param opcPackage the OPC package containing the presentation
     * @param path       the output file path
     * @param options    optional save options (currently unused for PPTX)
     * @throws IOException if the file cannot be written
     */
    @Override
    public void exportToPath(OpcPackage opcPackage, String path, ISaveOptions options) throws IOException {
        String previous = applyMainPartContentType(opcPackage);
        try {
            try (var out = new FileOutputStream(path)) {
                opcPackage.save(out);
            }
        } finally {
            restoreMainPartContentType(opcPackage, previous);
        }
    }

    /**
     * Export the presentation to a stream.
     *
     * @param opcPackage the OPC package containing the presentation
     * @param stream     the output stream
     * @param options    optional save options (currently unused for PPTX)
     * @throws IOException if the stream cannot be written to
     */
    @Override
    public void exportToStream(OpcPackage opcPackage, OutputStream stream, ISaveOptions options) throws IOException {
        String previous = applyMainPartContentType(opcPackage);
        try {
            opcPackage.save(stream);
        } finally {
            restoreMainPartContentType(opcPackage, previous);
        }
    }

    /**
     * Declares {@code /ppt/presentation.xml} to be of the target format's content type.
     *
     * <p>ISO/IEC 29500-2 makes the content type the identity of a part: a package whose
     * main part is declared {@code presentationml.presentation.main+xml} is a presentation
     * whatever the file is named, and PowerPoint refuses a {@code .potx} or {@code .ppsx}
     * whose declared type disagrees with its extension.</p>
     *
     * @param opcPackage the OPC package to update
     * @return the content type that was declared before this call, or {@code null} if none
     */
    String applyMainPartContentType(OpcPackage opcPackage) {
        String contentType = CONTENT_TYPES.get(targetFormat);
        if (contentType == null) {
            throw new IllegalStateException(
                    "No main-part content type is defined for format '" + targetFormat + "'");
        }
        var contentTypes = new ContentTypesManager(opcPackage);
        String previous = contentTypes.getContentType(MAIN_PART_NAME).orElse(null);
        if (contentType.equals(previous)) {
            return previous;
        }
        contentTypes.addOverride(MAIN_PART_NAME, contentType);
        contentTypes.save();
        return previous;
    }

    /**
     * Restores the content type that {@link #applyMainPartContentType} replaced, so that
     * exporting in one format does not change how a later save of the same presentation
     * describes itself.
     *
     * @param opcPackage the OPC package to update
     * @param previous   the content type to restore, or {@code null} to leave as written
     */
    private void restoreMainPartContentType(OpcPackage opcPackage, String previous) {
        if (previous == null || previous.equals(CONTENT_TYPES.get(targetFormat))) {
            return;
        }
        var contentTypes = new ContentTypesManager(opcPackage);
        contentTypes.addOverride(MAIN_PART_NAME, previous);
        contentTypes.save();
    }

    /**
     * Get all OPC-based presentation formats supported by this exporter.
     *
     * @return list of SaveFormat value strings
     */
    @Override
    public List<String> getSupportedFormats() {
        return List.copyOf(CONTENT_TYPES.keySet());
    }

    /**
     * Get the target format this exporter is configured for.
     *
     * @return the target format string
     */
    public String getTargetFormat() {
        return targetFormat;
    }
}
