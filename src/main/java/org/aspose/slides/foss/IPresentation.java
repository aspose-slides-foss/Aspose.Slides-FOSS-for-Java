package org.aspose.slides.foss;

import org.aspose.slides.foss.export.ISaveOptions;
import org.aspose.slides.foss.export.SaveFormat;

import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDateTime;

/**
 * Represents a presentation document.
 */
public interface IPresentation extends IPresentationComponent, AutoCloseable {

    /**
     * Returns date and time which will substitute content of datetime fields.
     * Time of this Presentation object creation by default.
     * Read/write {@link LocalDateTime}.
     *
     * @return the current date time
     */
    LocalDateTime getCurrentDateTime();

    /**
     * Sets date and time which will substitute content of datetime fields.
     *
     * @param value the date time to set
     */
    void setCurrentDateTime(LocalDateTime value);

    /**
     * Returns a list of all slides that are defined in the presentation.
     * Read-only {@link ISlideCollection}.
     *
     * @return the slide collection
     */
    ISlideCollection getSlides();

    /**
     * Returns notes slide size object.
     * Read-only {@link INotesSize}.
     *
     * @return the notes size
     */
    INotesSize getNotesSize();

    /**
     * Returns a list of all layout slides that are defined in the presentation.
     * Read-only {@link IGlobalLayoutSlideCollection}.
     *
     * @return the layout slide collection
     */
    IGlobalLayoutSlideCollection getLayoutSlides();

    /**
     * Returns a list of all master slides that are defined in the presentation.
     * Read-only {@link IMasterSlideCollection}.
     *
     * @return the master slide collection
     */
    IMasterSlideCollection getMasters();

    /**
     * Returns the collection of comments authors.
     * Read-only {@link ICommentAuthorCollection}.
     *
     * @return the comment author collection
     */
    ICommentAuthorCollection getCommentAuthors();

    /**
     * Returns DocumentProperties object which contains standard and custom document properties.
     * Read-only {@link IDocumentProperties}.
     *
     * @return the document properties
     */
    IDocumentProperties getDocumentProperties();

    /**
     * Returns the collection of all images in the presentation.
     * Read-only {@link IImageCollection}.
     *
     * @return the image collection
     */
    IImageCollection getImages();

    /**
     * Returns information about from which format presentation was loaded.
     * Read-only {@link SourceFormat}.
     *
     * @return the source format
     */
    SourceFormat getSourceFormat();

    /**
     * Represents the first slide number in the presentation.
     * Read/write {@code int}.
     *
     * @return the first slide number
     */
    int getFirstSlideNumber();

    /**
     * Sets the first slide number in the presentation.
     *
     * @param value the first slide number
     */
    void setFirstSlideNumber(int value);

    /**
     * Returns the base {@link IPresentationComponent} interface.
     * Read-only.
     *
     * @return this instance as {@link IPresentationComponent}
     */
    IPresentationComponent asIPresentationComponent();

    /**
     * Saves the presentation to the given file path.
     *
     * @param path the file path
     * @throws IOException if an I/O error occurs
     */
    void save(String path) throws IOException;

    /**
     * Saves the presentation to the given output stream.
     *
     * @param stream the output stream
     * @throws IOException if an I/O error occurs
     */
    void save(OutputStream stream) throws IOException;

    /**
     * Saves the presentation to the given file path in the specified format.
     *
     * @param path   the file path
     * @param format the save format
     * @throws IOException if an I/O error occurs
     */
    void save(String path, SaveFormat format) throws IOException;

    /**
     * Saves the presentation to the given output stream in the specified format.
     *
     * @param stream the output stream
     * @param format the save format
     * @throws IOException if an I/O error occurs
     */
    void save(OutputStream stream, SaveFormat format) throws IOException;

    /**
     * Saves the presentation to the given file path with the specified options.
     *
     * @param path    the file path
     * @param format  the save format
     * @param options the save options
     * @throws IOException if an I/O error occurs
     */
    void save(String path, SaveFormat format, ISaveOptions options) throws IOException;

    /**
     * Saves the presentation to the given output stream with the specified options.
     *
     * @param stream  the output stream
     * @param format  the save format
     * @param options the save options
     * @throws IOException if an I/O error occurs
     */
    void save(OutputStream stream, SaveFormat format, ISaveOptions options) throws IOException;

    /**
     * Saves the specified slides of the presentation to the given file path.
     *
     * <p>The indices are zero-based positions in {@link #getSlides()}. Repeats are
     * ignored and the slides kept stay in document order, not the order they were
     * named in. An empty array is rejected, because saving no slides is not a thing
     * that can be done; {@code null} means every slide, which is what the overloads
     * without a slide array do.</p>
     *
     * @param path   the file path
     * @param slides the slide indices to save, or {@code null} for every slide
     * @param format the save format
     * @throws IOException if an I/O error occurs
     */
    void save(String path, int[] slides, SaveFormat format) throws IOException;

    /**
     * Saves the specified slides of the presentation to the given file path with options.
     *
     * <p>The indices are zero-based positions in {@link #getSlides()}. Repeats are
     * ignored and the slides kept stay in document order, not the order they were
     * named in. An empty array is rejected, because saving no slides is not a thing
     * that can be done; {@code null} means every slide, which is what the overloads
     * without a slide array do.</p>
     *
     * @param path    the file path
     * @param slides  the slide indices to save, or {@code null} for every slide
     * @param format  the save format
     * @param options the save options
     * @throws IOException if an I/O error occurs
     */
    void save(String path, int[] slides, SaveFormat format, ISaveOptions options) throws IOException;

    /**
     * Saves the specified slides of the presentation to the given output stream.
     *
     * <p>The indices are zero-based positions in {@link #getSlides()}. Repeats are
     * ignored and the slides kept stay in document order, not the order they were
     * named in. An empty array is rejected, because saving no slides is not a thing
     * that can be done; {@code null} means every slide, which is what the overloads
     * without a slide array do.</p>
     *
     * @param stream the output stream
     * @param slides the slide indices to save, or {@code null} for every slide
     * @param format the save format
     * @throws IOException if an I/O error occurs
     */
    void save(OutputStream stream, int[] slides, SaveFormat format) throws IOException;

    /**
     * Saves the specified slides of the presentation to the given output stream with options.
     *
     * <p>The indices are zero-based positions in {@link #getSlides()}. Repeats are
     * ignored and the slides kept stay in document order, not the order they were
     * named in. An empty array is rejected, because saving no slides is not a thing
     * that can be done; {@code null} means every slide, which is what the overloads
     * without a slide array do.</p>
     *
     * @param stream  the output stream
     * @param slides  the slide indices to save, or {@code null} for every slide
     * @param format  the save format
     * @param options the save options
     * @throws IOException if an I/O error occurs
     */
    void save(OutputStream stream, int[] slides, SaveFormat format, ISaveOptions options) throws IOException;

    /**
     * Saves the presentation using the specified save options.
     *
     * @param options the save options
     * @throws IOException if an I/O error occurs
     */
    void save(ISaveOptions options) throws IOException;

    /** Disposes of resources held by this presentation. */
    void dispose();
}
