package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.OpcPackage;

/**
 * The package part a piece of content is serialized into.
 *
 * <p>Anything that refers to an image — a picture frame, a picture fill, a picture bullet —
 * does so through {@code r:embed}, and that id has to be declared in the {@code .rels} of the
 * part the reference is written in. Knowing which part that is has to travel with the object
 * that offers the image-setting call; without it the call can only either invent an id that
 * resolves to nothing or accept the image and write nothing at all.</p>
 *
 * @param pkg      the package to declare the relationship in
 * @param partName the name of the part the content is written to
 */
record PartContext(OpcPackage pkg, String partName) {

    /**
     * @return {@code true} when both the package and the part name are known
     */
    boolean isBound() {
        return pkg != null && partName != null;
    }
}
