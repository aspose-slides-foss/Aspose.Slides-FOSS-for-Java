package org.aspose.slides.foss;

/**
 * Base exception for PPT-related errors.
 */
public class PptException extends RuntimeException {

    /**
     * The value the compiler computed for this class before it was declared, so that an
     * instance serialized by an earlier build still deserializes here.
     */
    private static final long serialVersionUID = 4558838558670875332L;

    /**
     * Creates a new {@code PptException} with no detail message.
     */
    public PptException() {
    }

    /**
     * Creates a new {@code PptException} with the specified detail message.
     *
     * @param message the detail message
     */
    public PptException(String message) {
        super(message);
    }

    /**
     * Creates a new {@code PptException} with the specified detail message and cause.
     *
     * @param message the detail message
     * @param cause   the cause of this exception
     */
    public PptException(String message, Throwable cause) {
        super(message, cause);
    }
}
