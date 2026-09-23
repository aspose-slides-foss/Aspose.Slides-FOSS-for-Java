package org.aspose.slides.foss;

import org.aspose.slides.foss.drawing.PointF;

import java.time.LocalDateTime;
import java.util.Optional;

/**
 * Represents a comment on a presentation slide.
 */
public interface IComment {

    /** Returns the comment text. */
    String getText();

    /** Sets the comment text. */
    void setText(String value);

    /** Returns the author of this comment. */
    ICommentAuthor getAuthor();

    /** Returns the slide this comment belongs to. */
    ISlide getSlide();

    /**
     * Returns the position of the comment.
     *
     * @return the position, in centimetres from the top-left corner of the slide
     */
    PointF getPosition();

    /** Returns the creation date/time. */
    LocalDateTime getCreatedTime();

    /**
     * Returns the parent comment, if any.
     *
     * @return an {@link Optional} containing the parent comment, or empty if this is a top-level comment
     */
    Optional<IComment> getParentComment();

    /**
     * Sets the parent comment.
     *
     * @param parent the parent comment, or {@code null} to make this a top-level comment
     */
    void setParentComment(IComment parent);

    /**
     * Removes this comment and all its replies from the parent collection.
     */
    void remove();
}
