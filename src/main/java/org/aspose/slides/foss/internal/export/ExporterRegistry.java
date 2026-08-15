package org.aspose.slides.foss.internal.export;

import java.lang.reflect.InvocationTargetException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * Central registry for format exporters.
 *
 * <p>This registry maintains a mapping from SaveFormat value strings to their
 * corresponding exporter classes. New exporters can be registered dynamically,
 * making it easy to add new export formats.</p>
 *
 * <p>All methods are static — the registry is a global singleton backed by
 * a class-level map.</p>
 *
 * <p>Usage:</p>
 * <pre>{@code
 * // Register an exporter
 * ExporterRegistry.register(PptxExporter.class);
 *
 * // Get an exporter for a format
 * ExporterBase exporter = ExporterRegistry.getExporter("Pptx").orElseThrow();
 * exporter.export(opcPackage, "output.pptx", null);
 * }</pre>
 */
public final class ExporterRegistry {

    private static final Map<String, Class<? extends ExporterBase>> EXPORTERS = new LinkedHashMap<>();

    static {
        // The built-in exporters. Registration is done here rather than in a static
        // initializer on each exporter, so that a lookup can never depend on whether
        // the exporter class happens to have been loaded already.
        register(PptxExporter.class);
    }

    private ExporterRegistry() {
        // utility class
    }

    /**
     * Register an exporter class for all formats it supports.
     *
     * <p>The class must have a public no-arg constructor and must implement
     * {@link ExporterBase#getSupportedFormats()}.</p>
     *
     * @param exporterClass the exporter class to register
     */
    public static void register(Class<? extends ExporterBase> exporterClass) {
        List<String> formats = getSupportedFormatsFromClass(exporterClass);
        for (String format : formats) {
            EXPORTERS.put(format, exporterClass);
        }
    }

    /**
     * Unregister an exporter for a specific format.
     *
     * @param formatValue the SaveFormat value string to unregister
     * @return {@code true} if the format was registered and has been removed,
     *         {@code false} if it was not found
     */
    public static boolean unregister(String formatValue) {
        return EXPORTERS.remove(formatValue) != null;
    }

    /**
     * Get a new exporter instance for a specific format.
     *
     * <p>An exporter that serves several formats declares a constructor taking the format
     * value; it is preferred over the no-arg constructor so that the instance knows which
     * of its formats was asked for. Without this an exporter registered for six formats
     * would be built in its default one every time.</p>
     *
     * @param formatValue the SaveFormat value string (e.g., "Pptx", "Pdf")
     * @return an exporter instance, or empty if no exporter is registered for this format
     */
    public static Optional<ExporterBase> getExporter(String formatValue) {
        Class<? extends ExporterBase> clazz = EXPORTERS.get(formatValue);
        if (clazz == null) {
            return Optional.empty();
        }
        try {
            try {
                return Optional.of(clazz.getDeclaredConstructor(String.class).newInstance(formatValue));
            } catch (NoSuchMethodException e) {
                return Optional.of(clazz.getDeclaredConstructor().newInstance());
            }
        } catch (InstantiationException | IllegalAccessException
                 | InvocationTargetException | NoSuchMethodException e) {
            throw new IllegalStateException(
                    "Cannot instantiate exporter class " + clazz.getName(), e);
        }
    }

    /**
     * Get the exporter class for a specific format without instantiating it.
     *
     * @param formatValue the SaveFormat value string
     * @return the exporter class, or empty if not registered
     */
    public static Optional<Class<? extends ExporterBase>> getExporterClass(String formatValue) {
        return Optional.ofNullable(EXPORTERS.get(formatValue));
    }

    /**
     * Check if a format has a registered exporter.
     *
     * @param formatValue the SaveFormat value string
     * @return {@code true} if an exporter is registered for this format
     */
    public static boolean isFormatSupported(String formatValue) {
        return EXPORTERS.containsKey(formatValue);
    }

    /**
     * Get all formats that have registered exporters.
     *
     * @return list of SaveFormat value strings (insertion-ordered)
     */
    public static List<String> getSupportedFormats() {
        return List.copyOf(EXPORTERS.keySet());
    }

    /**
     * Clear all registered exporters. Intended for testing.
     */
    public static void clear() {
        EXPORTERS.clear();
    }

    /**
     * Instantiate the exporter class temporarily to query its supported formats.
     */
    private static List<String> getSupportedFormatsFromClass(Class<? extends ExporterBase> exporterClass) {
        try {
            ExporterBase instance = exporterClass.getDeclaredConstructor().newInstance();
            return instance.getSupportedFormats();
        } catch (InstantiationException | IllegalAccessException
                 | InvocationTargetException | NoSuchMethodException e) {
            throw new IllegalStateException(
                    "Cannot instantiate exporter class " + exporterClass.getName()
                            + " to query supported formats", e);
        }
    }
}
