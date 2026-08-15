package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.ImageUtils;
import org.aspose.slides.foss.internal.pptx.OpcPackage;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Represents a collection of images in a presentation.
 */
public final class ImageCollection implements IImageCollection {

    private static final Pattern IMAGE_NUMBER_PATTERN =
            Pattern.compile("ppt/media/image(\\d+)\\.");

    private final List<IPPImage> images = new ArrayList<>();
    private OpcPackage pkg;

    /**
     * Creates an image collection backed by the given OPC package.
     *
     * @param pkg the OPC package for storing image parts
     */
    public ImageCollection(OpcPackage pkg) {
        this.pkg = pkg;
    }

    /**
     * Creates an uninitialized image collection.
     *
     * <p>Call {@link #initInternal(OpcPackage)} to populate from a package.</p>
     */
    public ImageCollection() {
        // default constructor; requires initInternal to be called
    }

    /**
     * Internal initialization — scans the package for all images in {@code ppt/media/}.
     *
     * @param opcPackage the OPC package containing the presentation
     */
    public void initInternal(OpcPackage opcPackage) {
        this.pkg = opcPackage;
        this.images.clear();

        List<String> partNames = new ArrayList<>(opcPackage.getPartNames());
        java.util.Collections.sort(partNames);

        for (String partName : partNames) {
            if (partName.startsWith("ppt/media/")) {
                byte[] imageData = opcPackage.getPartBytes(partName);
                if (imageData != null && imageData.length > 0) {
                    String contentType = ImageUtils.guessContentType(imageData);
                    PPImage ppImage = new PPImage();
                    ppImage.initInternal(opcPackage, partName, imageData, contentType);
                    images.add(ppImage);
                }
            }
        }
    }

    @Override
    public IPPImage addImage(IImage image) {
        Objects.requireNonNull(image, "image");
        byte[] data;
        if (image instanceof Image img) {
            data = img.getData();
        } else {
            throw new IllegalArgumentException(
                    "Unsupported IImage implementation: " + image.getClass().getName());
        }
        return addImage(data);
    }

    @Override
    public IPPImage addImage(byte[] imageData) {
        Objects.requireNonNull(imageData, "imageData");
        // Check for duplicate images (same binary data)
        for (IPPImage existing : images) {
            if (java.util.Arrays.equals(existing.getBinaryData(), imageData)) {
                return existing;
            }
        }

        String contentType = ImageUtils.guessContentType(imageData);
        String extension = extensionForContentType(contentType);
        int imageNumber = findNextImageNumber();
        String partName = "ppt/media/image" + imageNumber + extension;

        pkg.setPartBytes(partName, imageData);
        addContentTypeOverride(partName, contentType);

        PPImage ppImage = new PPImage();
        ppImage.initInternal(pkg, partName, imageData, contentType);
        images.add(ppImage);
        return ppImage;
    }

    @Override
    public IPPImage addImage(InputStream stream) {
        Objects.requireNonNull(stream, "stream");
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            stream.transferTo(baos);
            return addImage(baos.toByteArray());
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    @Override
    public IPPImage get(int index) {
        return images.get(index);
    }

    @Override
    public int size() {
        return images.size();
    }

    @Override
    public List<IPPImage> asICollection() {
        return List.copyOf(images);
    }

    @Override
    public Iterable<IPPImage> asIEnumerable() {
        return List.copyOf(images);
    }

    @Override
    public Iterator<IPPImage> iterator() {
        return List.copyOf(images).iterator();
    }

    /**
     * Finds the next available image number in {@code ppt/media/}.
     *
     * <p>Scans all existing part names matching {@code ppt/media/imageN.*} and
     * returns the smallest positive integer not yet in use.</p>
     *
     * @return the next available image number
     */
    public int findNextImageNumber() {
        Set<Integer> existingNumbers = new HashSet<>();
        for (String partName : pkg.getPartNames()) {
            if (partName.startsWith("ppt/media/")) {
                Matcher matcher = IMAGE_NUMBER_PATTERN.matcher(partName);
                if (matcher.find()) {
                    existingNumbers.add(Integer.parseInt(matcher.group(1)));
                }
            }
        }

        int num = 1;
        while (existingNumbers.contains(num)) {
            num++;
        }
        return num;
    }

    /**
     * Returns the internal list of images.
     *
     * @return the list
     */
    List<IPPImage> getInternalList() {
        return images;
    }

    /**
     * Loads images from existing OPC package parts during deserialization.
     */
    void loadFromPackage() {
        initInternal(this.pkg);
    }

    private static String extensionForContentType(String contentType) {
        return switch (contentType) {
            case "image/png" -> ".png";
            case "image/jpeg" -> ".jpeg";
            case "image/gif" -> ".gif";
            case "image/bmp" -> ".bmp";
            case "image/tiff" -> ".tiff";
            case "image/webp" -> ".webp";
            default -> ".bin";
        };
    }

    private void addContentTypeOverride(String partName, String contentType) {
        // Extract extension and add Default entry to [Content_Types].xml
        byte[] ctData = pkg.getPartBytes("[Content_Types].xml");
        if (ctData == null) return;
        String ct = new String(ctData, java.nio.charset.StandardCharsets.UTF_8);
        String ext = partName.substring(partName.lastIndexOf('.') + 1);
        // Check if extension already registered
        if (ct.contains("Extension=\"" + ext + "\"")) return;
        // Insert a Default element before </Types>
        String defaultEntry = "<Default Extension=\"" + ext + "\" ContentType=\"" + contentType + "\"/>";
        ct = ct.replace("</Types>", defaultEntry + "</Types>");
        pkg.setPartBytes("[Content_Types].xml", ct.getBytes(java.nio.charset.StandardCharsets.UTF_8));
    }
}
