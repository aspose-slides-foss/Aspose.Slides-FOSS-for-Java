package org.aspose.slides.foss;

import org.aspose.slides.foss.drawing.PointF;

import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;

/**
 * Represents a collection of comments belonging to a single author.
 */
public interface ICommentCollection {

    /**
     * Adds a comment to the collection.
     *
     * @param text        the comment text
     * @param slide       the target slide
     * @param position    the position on the slide, in centimetres from its top-left corner
     * @param createdTime the creation date/time
     * @return the new comment
     */
    IComment addComment(String text, ISlide slide, PointF position, LocalDateTime createdTime);

    /**
     * Inserts a comment at the given index.
     *
     * @param index       the insertion index
     * @param text        the comment text
     * @param slide       the target slide
     * @param position    the position on the slide, in centimetres from its top-left corner
     * @param createdTime the creation date/time
     * @return the new comment
     */
    IComment insertComment(int index, String text, ISlide slide, PointF position, LocalDateTime createdTime);

    /**
     * Removes the comment at the given index.
     *
     * @param index the index to remove
     */
    void removeAt(int index);

    /**
     * Removes the given comment (and its replies) from the collection.
     *
     * @param comment the comment to remove
     */
    void remove(IComment comment);

    /** Removes all comments from this collection. */
    void clear();

    /**
     * Returns the comment at the given index.
     *
     * @param index the zero-based index
     * @return the comment
     */
    IComment get(int index);

    /** Returns the number of comments. */
    int size();

    /**
     * Returns all comments as an array.
     *
     * @return an array of all comments
     */
    IComment[] toArray();

    /**
     * Returns a subset of comments as an array.
     *
     * @param startIndex the starting index
     * @param count      the number of comments to return
     * @return an array of comments
     */
    IComment[] toArray(int startIndex, int count);

    /**
     * Returns this collection as a list.
     *
     * @return a list view of the comments
     */
    List<IComment> asICollection();

    /**
     * Returns this collection as an iterable.
     *
     * @return an iterable over the comments
     */
    Iterable<IComment> asIEnumerable();

    /**
     * Finds a comment by its unique index (per author).
     *
     * @param idx the comment index to search for
     * @return an {@link Optional} containing the matching comment, or empty if not found
     */
    Optional<IComment> findCommentByIdx(int idx);
}
