package org.aspose.slides.foss;

/**
 * Exception thrown when a PPT file cannot be read.
 */
public class PptReadException extends PptException {

    /**
     * The value the compiler computed for this class before it was declared, so that an
     * instance serialized by an earlier build still deserializes here.
     */
    private static final long serialVersionUID = 2666054243868649730L;

    /**
     * Creates a new {@code PptReadException} with no detail message.
     */
    public PptReadException() {
    }

    /**
     * Creates a new {@code PptReadException} with the specified detail message.
     *
     * @param message the detail message
     */
    public PptReadException(String message) {
        super(message);
    }

    /**
     * Creates a new {@code PptReadException} with the specified detail message and cause.
     *
     * @param message the detail message
     * @param cause   the cause of this exception
     */
    public PptReadException(String message, Throwable cause) {
        super(message, cause);
    }
}
