package org.aspose.slides.foss.internal.xml;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerFactory;

/**
 * The XML factories this library uses, configured so that a document cannot reach outside itself.
 *
 * <p>A {@code .pptx} is a ZIP archive full of XML, and when the file came from somewhere else every
 * byte of both layers is written by whoever sent it. A {@code DocumentBuilderFactory} left at its
 * defaults resolves a {@code <!DOCTYPE>} and the external entities declared in it, so a part
 * containing</p>
 *
 * <pre>{@code
 * <!DOCTYPE p:sld [<!ENTITY x SYSTEM "file:///etc/passwd">]>
 * ... <a:t>&x;</a:t> ...
 * }</pre>
 *
 * <p>makes the parser read that file and place its contents in the presentation, where the calling
 * application will treat it as slide text. The same mechanism reaches network URLs, and a
 * self-referencing entity expands until the process runs out of memory.</p>
 *
 * <p>None of it is needed: ECMA-376 packages carry no DTD, so refusing a {@code DOCTYPE} outright
 * is both the strongest setting and a lossless one. A part that declares one is rejected with a
 * parse error rather than parsed with entities disabled, which is deliberate — a presentation part
 * with a DTD in it is not a presentation part.</p>
 *
 * <p>Every parser and serializer in this library is created here, so that no call site can be
 * left out by being forgotten.</p>
 */
public final class SecureXml {

    /** Xerces feature name; rejecting a DOCTYPE outright subsumes every entity setting below it. */
    private static final String DISALLOW_DOCTYPE =
            "http://apache.org/xml/features/disallow-doctype-decl";

    /** SAX feature name, kept as defence in depth for a parser that ignores the one above. */
    private static final String EXTERNAL_GENERAL_ENTITIES =
            "http://xml.org/sax/features/external-general-entities";

    /** SAX feature name, likewise. */
    private static final String EXTERNAL_PARAMETER_ENTITIES =
            "http://xml.org/sax/features/external-parameter-entities";

    /** Xerces feature name: do not fetch an external DTD subset either. */
    private static final String LOAD_EXTERNAL_DTD =
            "http://apache.org/xml/features/nonvalidating/load-external-dtd";

    private SecureXml() {
    }

    /**
     * Returns a namespace-aware {@link DocumentBuilderFactory} that resolves nothing outside the
     * document it is given.
     *
     * <p>Namespace awareness is on because every part of an Office Open XML package is
     * namespaced and the whole library selects elements by namespace URI.</p>
     *
     * @return a configured factory
     * @throws IllegalStateException if the platform's XML implementation refuses a setting that
     *                               is required for safe parsing, rather than silently parsing
     *                               unsafely
     */
    public static DocumentBuilderFactory documentBuilderFactory() {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        try {
            factory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
            factory.setFeature(DISALLOW_DOCTYPE, true);
            factory.setFeature(EXTERNAL_GENERAL_ENTITIES, false);
            factory.setFeature(EXTERNAL_PARAMETER_ENTITIES, false);
            factory.setFeature(LOAD_EXTERNAL_DTD, false);
        } catch (ParserConfigurationException e) {
            throw new IllegalStateException(
                    "This XML parser cannot be configured to reject external entities; "
                            + "refusing to parse rather than parsing unsafely", e);
        }
        factory.setXIncludeAware(false);
        factory.setExpandEntityReferences(false);
        return factory;
    }

    /**
     * Returns a {@link TransformerFactory} that will not load an external DTD or stylesheet.
     *
     * <p>This library only ever serializes a DOM it built itself, so nothing here is
     * attacker-controlled today. The properties are set anyway, because the cost is two lines and
     * the alternative is relying on that staying true.</p>
     *
     * @return a configured factory
     */
    public static TransformerFactory transformerFactory() {
        TransformerFactory factory = TransformerFactory.newInstance();
        try {
            factory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
        } catch (javax.xml.transform.TransformerConfigurationException e) {
            throw new IllegalStateException(
                    "This XML transformer cannot be configured for secure processing", e);
        }
        // An implementation that does not know these attributes throws rather than ignoring them,
        // and it cannot load what it does not implement, so not knowing them is not a failure.
        try {
            factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
            factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_STYLESHEET, "");
        } catch (IllegalArgumentException ignored) {
            // Not supported by this implementation.
        }
        return factory;
    }
}
