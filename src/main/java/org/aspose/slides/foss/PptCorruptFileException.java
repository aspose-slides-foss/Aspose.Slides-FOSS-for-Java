package org.aspose.slides.foss;

/**
 * Exception thrown when a PPT file is corrupt and cannot be processed.
 */
public final class PptCorruptFileException extends PptReadException {

    /**
     * The value the compiler computed for this class before it was declared, so that an
     * instance serialized by an earlier build still deserializes here.
     */
    private static final long serialVersionUID = -9139983423536833693L;

    /**
     * Creates a new {@code PptCorruptFileException} with no detail message.
     */
    public PptCorruptFileException() {
    }

    /**
     * Creates a new {@code PptCorruptFileException} with the specified detail message.
     *
     * @param message the detail message
     */
    public PptCorruptFileException(String message) {
        super(message);
    }

    /**
     * Creates a new {@code PptCorruptFileException} with the specified detail message and cause.
     *
     * @param message the detail message
     * @param cause   the cause of this exception
     */
    public PptCorruptFileException(String message, Throwable cause) {
        super(message, cause);
    }
}
