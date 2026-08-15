package org.aspose.slides.foss.conformance;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * Read-only view of a produced {@code .pptx} as an OPC package: a ZIP container of parts.
 *
 * <p>This is the only way conformance tests are allowed to look at output. Nothing here calls
 * into the library that wrote the file, so an assertion made through this class cannot be
 * satisfied by the library simply agreeing with itself.</p>
 *
 * <p>Part names are ZIP entry names with no leading slash, e.g. {@code ppt/slides/slide1.xml}
 * and {@code [Content_Types].xml}.</p>
 */
public final class PptxPackage implements AutoCloseable {

    /** The OPC relationships namespace, used for {@code r:id}, {@code r:embed}, {@code r:link}. */
    public static final String NS_REL =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /** The namespace of {@code [Content_Types].xml}. */
    public static final String NS_CONTENT_TYPES =
            "http://schemas.openxmlformats.org/package/2006/content-types";
    /** The namespace of a {@code .rels} part. */
    public static final String NS_RELS_PART =
            "http://schemas.openxmlformats.org/package/2006/relationships";
    /** The PresentationML namespace. */
    public static final String NS_P =
            "http://schemas.openxmlformats.org/presentationml/2006/main";
    /** The DrawingML namespace. */
    public static final String NS_A =
            "http://schemas.openxmlformats.org/drawingml/2006/main";

    /** The name of the content-types part. */
    public static final String CONTENT_TYPES_PART = "[Content_Types].xml";

    private final Path path;
    private final ZipFile zip;

    private PptxPackage(Path path, ZipFile zip) {
        this.path = path;
        this.zip = zip;
    }

    /**
     * Opens a produced file as a ZIP container.
     *
     * @param pptx path to the file
     * @return the open package; close it when done
     * @throws IOException if the file is not a readable ZIP
     */
    public static PptxPackage open(Path pptx) throws IOException {
        return new PptxPackage(pptx, new ZipFile(pptx.toFile()));
    }

    /** @return the path this package was opened from */
    public Path path() {
        return path;
    }

    /**
     * @return every entry name in the package, sorted, exactly as stored
     */
    public List<String> entryNames() {
        List<String> names = new ArrayList<>();
        for (ZipEntry e : Collections.list(zip.entries())) {
            if (!e.isDirectory()) {
                names.add(e.getName());
            }
        }
        Collections.sort(names);
        return names;
    }

    /** @return every entry whose name ends in {@code .xml} or {@code .rels}, sorted */
    public List<String> xmlPartNames() {
        List<String> names = new ArrayList<>();
        for (String n : entryNames()) {
            if (n.endsWith(".xml") || n.endsWith(".rels")) {
                names.add(n);
            }
        }
        return names;
    }

    /**
     * @param partName the part name
     * @return whether the package contains that part
     */
    public boolean hasPart(String partName) {
        return zip.getEntry(partName) != null;
    }

    /**
     * @param partName the part name
     * @return the raw bytes of the part
     * @throws IOException if the part is missing or unreadable
     */
    public byte[] bytes(String partName) throws IOException {
        ZipEntry entry = zip.getEntry(partName);
        if (entry == null) {
            throw new IOException("no such part: " + partName + " in " + path
                    + " (parts: " + entryNames() + ")");
        }
        try (InputStream in = zip.getInputStream(entry)) {
            return in.readAllBytes();
        }
    }

    /**
     * @param partName the part name
     * @return the part decoded as UTF-8 text
     * @throws IOException if the part is missing or unreadable
     */
    public String text(String partName) throws IOException {
        return new String(bytes(partName), StandardCharsets.UTF_8);
    }

    /**
     * Parses a part as a namespace-aware DOM.
     *
     * @param partName the part name
     * @return the parsed document
     * @throws IOException if the part is missing, unreadable or not well-formed XML
     */
    public Document xml(String partName) throws IOException {
        byte[] raw = bytes(partName);
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            return builder.parse(new ByteArrayInputStream(raw));
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        } catch (Exception e) {
            throw new IOException("part " + partName + " of " + path + " is not well-formed XML", e);
        }
    }

    /**
     * @param partName the part name
     * @return the document element of that part
     * @throws IOException if the part is missing or not well-formed XML
     */
    public Element root(String partName) throws IOException {
        return xml(partName).getDocumentElement();
    }

    /**
     * Returns the name of the {@code .rels} part that owns the relationships of a given part.
     *
     * <p>{@code ppt/slides/slide1.xml} is owned by {@code ppt/slides/_rels/slide1.xml.rels};
     * the package root is owned by {@code _rels/.rels}.</p>
     *
     * @param partName the part name
     * @return the name of its relationships part, which need not exist
     */
    public static String relsPartNameFor(String partName) {
        int slash = partName.lastIndexOf('/');
        String dir = slash < 0 ? "" : partName.substring(0, slash + 1);
        String file = slash < 0 ? partName : partName.substring(slash + 1);
        return dir + "_rels/" + file + ".rels";
    }

    /**
     * Resolves a relationship target against the directory of the part that declares it.
     *
     * @param owningPartName the part whose {@code .rels} carries the relationship
     * @param target         the {@code Target} attribute, e.g. {@code ../media/image1.png}
     * @return the absolute part name inside the package, with no leading slash
     */
    public static String resolveTarget(String owningPartName, String target) {
        if (target.startsWith("/")) {
            return target.substring(1);
        }
        int slash = owningPartName.lastIndexOf('/');
        String dir = slash < 0 ? "" : owningPartName.substring(0, slash);
        List<String> segments = new ArrayList<>();
        if (!dir.isEmpty()) {
            Collections.addAll(segments, dir.split("/"));
        }
        for (String segment : target.split("/")) {
            if (segment.equals(".") || segment.isEmpty()) {
                continue;
            }
            if (segment.equals("..")) {
                if (!segments.isEmpty()) {
                    segments.remove(segments.size() - 1);
                }
            } else {
                segments.add(segment);
            }
        }
        return String.join("/", segments);
    }

    @Override
    public void close() throws IOException {
        zip.close();
    }
}
