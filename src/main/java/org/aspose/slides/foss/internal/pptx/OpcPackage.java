package org.aspose.slides.foss.internal.pptx;

import org.w3c.dom.Document;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import org.aspose.slides.foss.internal.xml.SecureXml;

/**
 * Minimal OPC (Open Packaging Conventions) package backed by a ZIP archive.
 *
 * <p>Stores part contents as byte arrays keyed by their URI path (without leading slash).</p>
 */
public final class OpcPackage {

    private final Map<String, byte[]> parts = new LinkedHashMap<>();

    /** Creates an empty package. */
    public OpcPackage() {
    }

    /**
     * Loads a package from a ZIP input stream.
     *
     * @param in the input stream
     * @throws IOException if an I/O error occurs
     */
    public void load(InputStream in) throws IOException {
        parts.clear();
        try (var zis = new ZipInputStream(in)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (!entry.isDirectory()) {
                    parts.put(entry.getName(), zis.readAllBytes());
                }
            }
        }
    }

    /**
     * Saves the package as a ZIP archive to the given output stream.
     *
     * @param out the output stream
     * @throws IOException if an I/O error occurs
     */
    public void save(OutputStream out) throws IOException {
        try (var zos = new ZipOutputStream(out)) {
            for (var e : parts.entrySet()) {
                zos.putNextEntry(new ZipEntry(e.getKey()));
                zos.write(e.getValue());
                zos.closeEntry();
            }
        }
    }

    /** Returns whether a part with the given URI exists. */
    public boolean hasPart(String uri) {
        return parts.containsKey(uri);
    }

    /** Returns the raw bytes of the given part, or {@code null} if not found. */
    public byte[] getPartBytes(String uri) {
        return parts.get(uri);
    }

    /** Sets the raw bytes for the given part URI. */
    public void setPartBytes(String uri, byte[] data) {
        parts.put(uri, data);
    }

    /** Removes a part by URI. */
    public void removePart(String uri) {
        parts.remove(uri);
    }

    /**
     * Removes a part the way ISO/IEC 29500-2 requires, rather than only its bytes.
     *
     * <p>Deleting a part is four operations that have to happen together: drop the
     * {@code Relationship} in the owner's {@code .rels}, drop the {@code Override} in
     * {@code [Content_Types].xml}, drop the part itself and drop the part's own
     * {@code .rels}. Doing only the last leaves a relationship and a content type
     * pointing at nothing, which strict readers — PowerPoint among them — reject.</p>
     *
     * <p>Parts that the removed part was the only reference to (its notes slide, its
     * comments, images used nowhere else) are removed the same way, recursively. A part
     * that anything else still points at is left alone.</p>
     *
     * @param partUri       the part to remove
     * @param owningPartUri the part whose relationship points at it, or {@code null} if none
     */
    public void removePartCascading(String partUri, String owningPartUri) {
        if (owningPartUri != null) {
            var ownerRels = new RelsHelper(this, owningPartUri);
            if (ownerRels.removeRelationshipsTo(partUri)) {
                ownerRels.save();
            }
        }
        removePartTree(partUri);
    }

    private void removePartTree(String partUri) {
        if (!hasPart(partUri)) {
            return;
        }
        List<String> targets = new RelsHelper(this, partUri).internalTargets();

        parts.remove(partUri);
        parts.remove(RelsHelper.getRelsPartName(partUri));
        var contentTypes = new ContentTypesManager(this);
        if (contentTypes.removeOverride("/" + partUri)) {
            contentTypes.save();
        }

        for (String target : targets) {
            if (hasPart(target) && !isReferencedByAnyPart(target)) {
                removePartTree(target);
            }
        }
    }

    /**
     * Returns whether any relationship anywhere in the package points at the given part.
     *
     * @param partUri the part to look for
     * @return {@code true} if some {@code .rels} in the package targets it
     */
    public boolean isReferencedByAnyPart(String partUri) {
        for (String name : List.copyOf(parts.keySet())) {
            if (!name.endsWith(".rels")) {
                continue;
            }
            String owner = RelsHelper.getSourcePartName(name);
            if (new RelsHelper(this, owner).internalTargets().contains(partUri)) {
                return true;
            }
        }
        return false;
    }

    /** Returns all part URIs as an unmodifiable set. */
    public Set<String> getPartNames() {
        return java.util.Collections.unmodifiableSet(parts.keySet());
    }

    /** Clears all parts, releasing memory held by byte arrays. */
    public void clear() {
        parts.clear();
    }

    /**
     * Parses the given part as an XML document.
     *
     * @param uri the part URI
     * @return the parsed document, or {@code null} if the part does not exist
     */
    public Document parseXml(String uri) {
        byte[] data = parts.get(uri);
        if (data == null) return null;
        try {
            var factory = SecureXml.documentBuilderFactory();
            factory.setNamespaceAware(true);
            return factory.newDocumentBuilder()
                    .parse(new ByteArrayInputStream(data));
        } catch (Exception e) {
            throw new IllegalStateException("Failed to parse XML part: " + uri, e);
        }
    }

    /**
     * Serializes the given XML document and stores it as a part.
     *
     * @param uri the part URI
     * @param doc the XML document
     */
    public void serializeXml(String uri, Document doc) {
        // A text body left with no paragraph is invalid however it came to be empty.
        // Repaired here, on a copy, rather than on each route that can empty one: the
        // paragraph collection reads this same tree and must still be able to report a
        // count of zero after being cleared.
        Document toWrite = TextBodies.withParagraphs(doc);
        try {
            var transformer = SecureXml.transformerFactory().newTransformer();
            // Never re-indent. The parser keeps the whitespace text nodes of the
            // document it read, so indenting on output adds a fresh layer of them
            // to what is already indented: opening a file and saving it unchanged
            // grew it every time, without bound in a loop, and made byte length
            // useless as a "did anything change" signal.
            transformer.setOutputProperty(OutputKeys.INDENT, "no");
            transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
            transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            var sw = new ByteArrayOutputStream();
            transformer.transform(new DOMSource(toWrite), new StreamResult(sw));
            parts.put(uri, sw.toByteArray());
        } catch (Exception e) {
            throw new IllegalStateException("Failed to serialize XML part: " + uri, e);
        }
    }

    /**
     * Creates a new empty XML document.
     *
     * @return a new DOM Document
     */
    public static Document newDocument() {
        try {
            return SecureXml.documentBuilderFactory()
                    .newDocumentBuilder()
                    .newDocument();
        } catch (javax.xml.parsers.ParserConfigurationException e) {
            throw new IllegalStateException("Failed to create XML document", e);
        }
    }
}
