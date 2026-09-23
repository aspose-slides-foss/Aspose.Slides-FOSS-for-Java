package org.aspose.slides.foss.conformance;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.function.UnaryOperator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * Copies a package with a deliberate edit, for calibrating the package rules.
 *
 * <p>Used by {@link HarnessCalibrationTest}, because a rule that has never been shown to fail is
 * not evidence of anything, so each rule is pointed at a package damaged in exactly the way it
 * claims to detect — and by {@link UntrustedInputConformanceTest}, which needs a package no
 * supported API can write. The part-numbering and layout-order tests use it to reproduce, in a
 * valid package, the numbering and relationship order other producers write.</p>
 */
final class ZipSurgery {

    private ZipSurgery() {
    }

    /** Copies {@code source} to {@code target}, adding a part that nothing declares. */
    static Path copyWithExtraPart(Path source, Path target, String partName, byte[] content)
            throws IOException {
        rewrite(source, target, (name, bytes) -> bytes, out -> {
            out.putNextEntry(new ZipEntry(partName));
            out.write(content);
            out.closeEntry();
        });
        return target;
    }

    /** Copies {@code source} to {@code target}, omitting one part and changing nothing else. */
    static Path copyWithoutPart(Path source, Path target, String partName) throws IOException {
        rewrite(source, target, (name, bytes) -> name.equals(partName) ? null : bytes, out -> {
        });
        return target;
    }

    /** Copies {@code source} to {@code target}, rewriting the text of one XML part. */
    static Path copyWithReplacement(Path source, Path target, String partName,
                                    UnaryOperator<String> edit) throws IOException {
        rewrite(source, target, (name, bytes) -> {
            if (!name.equals(partName)) {
                return bytes;
            }
            return edit.apply(new String(bytes, StandardCharsets.UTF_8))
                    .getBytes(StandardCharsets.UTF_8);
        }, out -> {
        });
        return target;
    }

    /**
     * Copies {@code source} to {@code target} with one part renamed, keeping the package valid.
     *
     * <p>The part's own {@code .rels} moves with it, and every relationship target and
     * content-type {@code Override} that names it is rewritten. References are matched as
     * {@code <directory>/<file name>"}, which covers the {@code ../dir/name.xml} targets and the
     * {@code /ppt/dir/name.xml} part names these fixtures use. Several producers number parts
     * independently of slide numbers, and this reproduces that without a binary fixture.</p>
     */
    static Path copyWithRenamedPart(Path source, Path target, String partName, String newPartName)
            throws IOException {
        String oldRels = PptxPackage.relsPartNameFor(partName);
        String newRels = PptxPackage.relsPartNameFor(newPartName);
        String oldRef = lastTwoSegments(partName) + "\"";
        String newRef = lastTwoSegments(newPartName) + "\"";
        try (ZipFile in = new ZipFile(source.toFile());
             OutputStream fileOut = Files.newOutputStream(target);
             ZipOutputStream out = new ZipOutputStream(fileOut)) {
            var entries = in.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                if (entry.isDirectory()) {
                    continue;
                }
                byte[] bytes;
                try (var stream = in.getInputStream(entry)) {
                    bytes = stream.readAllBytes();
                }
                String name = entry.getName();
                if (name.endsWith(".rels") || name.equals(PptxPackage.CONTENT_TYPES_PART)) {
                    bytes = new String(bytes, StandardCharsets.UTF_8).replace(oldRef, newRef)
                            .getBytes(StandardCharsets.UTF_8);
                }
                if (name.equals(partName)) {
                    name = newPartName;
                } else if (name.equals(oldRels)) {
                    name = newRels;
                }
                out.putNextEntry(new ZipEntry(name));
                out.write(bytes);
                out.closeEntry();
            }
        }
        return target;
    }

    private static String lastTwoSegments(String partName) {
        int last = partName.lastIndexOf('/');
        int previous = partName.lastIndexOf('/', last - 1);
        return partName.substring(previous + 1);
    }

    private static void rewrite(Path source, Path target, EntryEdit edit, Extra extra)
            throws IOException {
        try (ZipFile in = new ZipFile(source.toFile());
             OutputStream fileOut = Files.newOutputStream(target);
             ZipOutputStream out = new ZipOutputStream(fileOut)) {
            var entries = in.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                if (entry.isDirectory()) {
                    continue;
                }
                byte[] bytes;
                try (var stream = in.getInputStream(entry)) {
                    bytes = stream.readAllBytes();
                }
                byte[] edited = edit.apply(entry.getName(), bytes);
                if (edited == null) {
                    continue;
                }
                out.putNextEntry(new ZipEntry(entry.getName()));
                out.write(edited);
                out.closeEntry();
            }
            extra.write(out);
        }
    }

    @FunctionalInterface
    private interface EntryEdit {
        byte[] apply(String name, byte[] bytes);
    }

    @FunctionalInterface
    private interface Extra {
        void write(ZipOutputStream out) throws IOException;
    }
}
