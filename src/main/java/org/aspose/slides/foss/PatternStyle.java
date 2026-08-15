package org.aspose.slides.foss;

/**
 * Represents the pattern style.
 */
public enum PatternStyle {
    /** Not defined. */
    NOT_DEFINED("NotDefined", null),
    /** Unknown. */
    UNKNOWN("Unknown", null),
    /** 5 percent. */
    PERCENT05("Percent05", "pct5"),
    /** 10 percent. */
    PERCENT10("Percent10", "pct10"),
    /** 20 percent. */
    PERCENT20("Percent20", "pct20"),
    /** 25 percent. */
    PERCENT25("Percent25", "pct25"),
    /** 30 percent. */
    PERCENT30("Percent30", "pct30"),
    /** 40 percent. */
    PERCENT40("Percent40", "pct40"),
    /** 50 percent. */
    PERCENT50("Percent50", "pct50"),
    /** 60 percent. */
    PERCENT60("Percent60", "pct60"),
    /** 70 percent. */
    PERCENT70("Percent70", "pct70"),
    /** 75 percent. */
    PERCENT75("Percent75", "pct75"),
    /** 80 percent. */
    PERCENT80("Percent80", "pct80"),
    /** 90 percent. */
    PERCENT90("Percent90", "pct90"),
    /** Dark Horizontal. */
    DARK_HORIZONTAL("DarkHorizontal", "dkHorz"),
    /** Dark Vertical. */
    DARK_VERTICAL("DarkVertical", "dkVert"),
    /** Dark Downward Diagonal. */
    DARK_DOWNWARD_DIAGONAL("DarkDownwardDiagonal", "dkDnDiag"),
    /** Dark Upward Diagonal. */
    DARK_UPWARD_DIAGONAL("DarkUpwardDiagonal", "dkUpDiag"),
    /** Small Checker Board. */
    SMALL_CHECKER_BOARD("SmallCheckerBoard", "smCheck"),
    /** Trellis. */
    TRELLIS("Trellis", "trellis"),
    /** Light Horizontal. */
    LIGHT_HORIZONTAL("LightHorizontal", "ltHorz"),
    /** Light Vertical. */
    LIGHT_VERTICAL("LightVertical", "ltVert"),
    /** Light Downward Diagonal. */
    LIGHT_DOWNWARD_DIAGONAL("LightDownwardDiagonal", "ltDnDiag"),
    /** Light Upward Diagonal. */
    LIGHT_UPWARD_DIAGONAL("LightUpwardDiagonal", "ltUpDiag"),
    /** Small Grid. */
    SMALL_GRID("SmallGrid", "smGrid"),
    /** Dotted Diamond. */
    DOTTED_DIAMOND("DottedDiamond", "dotDmnd"),
    /** Wide Downward Diagonal. */
    WIDE_DOWNWARD_DIAGONAL("WideDownwardDiagonal", "wdDnDiag"),
    /** Wide Upward Diagonal. */
    WIDE_UPWARD_DIAGONAL("WideUpwardDiagonal", "wdUpDiag"),
    /** Dashed Downward Diagonal. */
    DASHED_DOWNWARD_DIAGONAL("DashedDownwardDiagonal", "dashDnDiag"),
    /** Dashed Upward Diagonal. */
    DASHED_UPWARD_DIAGONAL("DashedUpwardDiagonal", "dashUpDiag"),
    /** Narrow Vertical. */
    NARROW_VERTICAL("NarrowVertical", "narVert"),
    /** Narrow Horizontal. */
    NARROW_HORIZONTAL("NarrowHorizontal", "narHorz"),
    /** Dashed Vertical. */
    DASHED_VERTICAL("DashedVertical", "dashVert"),
    /** Dashed Horizontal. */
    DASHED_HORIZONTAL("DashedHorizontal", "dashHorz"),
    /** Large Confetti. */
    LARGE_CONFETTI("LargeConfetti", "lgConfetti"),
    /** Large Grid. */
    LARGE_GRID("LargeGrid", "lgGrid"),
    /** Horizontal Brick. */
    HORIZONTAL_BRICK("HorizontalBrick", "horzBrick"),
    /** Large Checker Board. */
    LARGE_CHECKER_BOARD("LargeCheckerBoard", "lgCheck"),
    /** Small Confetti. */
    SMALL_CONFETTI("SmallConfetti", "smConfetti"),
    /** Zigzag. */
    ZIGZAG("Zigzag", "zigZag"),
    /** Solid Diamond. */
    SOLID_DIAMOND("SolidDiamond", "solidDmnd"),
    /** Diagonal Brick. */
    DIAGONAL_BRICK("DiagonalBrick", "diagBrick"),
    /** Outlined Diamond. */
    OUTLINED_DIAMOND("OutlinedDiamond", "openDmnd"),
    /** Plaid. */
    PLAID("Plaid", "plaid"),
    /** Sphere. */
    SPHERE("Sphere", "sphere"),
    /** Weave. */
    WEAVE("Weave", "weave"),
    /** Dotted Grid. */
    DOTTED_GRID("DottedGrid", "dotGrid"),
    /** Divot. */
    DIVOT("Divot", "divot"),
    /** Shingle. */
    SHINGLE("Shingle", "shingle"),
    /** Wave. */
    WAVE("Wave", "wave"),
    /** Horizontal. */
    HORIZONTAL("Horizontal", "horz"),
    /** Vertical. */
    VERTICAL("Vertical", "vert"),
    /** Cross. */
    CROSS("Cross", "cross"),
    /** Downward Diagonal. */
    DOWNWARD_DIAGONAL("DownwardDiagonal", "dnDiag"),
    /** Upward Diagonal. */
    UPWARD_DIAGONAL("UpwardDiagonal", "upDiag"),
    /** Diagonal Cross. */
    DIAGONAL_CROSS("DiagonalCross", "diagCross");

    private final String value;
    private final String ooxml;

    PatternStyle(String value, String ooxml) {
        this.value = value;
        this.ooxml = ooxml;
    }

    /**
     * Returns the string value of this constant.
     *
     * @return the string value
     */
    public String getValue() {
        return value;
    }

    /**
     * Returns the {@code ST_PresetPatternVal} token for this pattern.
     *
     * <p>{@code ST_PresetPatternVal} (ECMA-376 §20.1.10.50) is a closed list of 54 tokens
     * that does not follow from the constant name: {@code DIAGONAL_BRICK} is
     * {@code diagBrick}, {@code PERCENT05} is {@code pct5}, {@code ZIGZAG} is
     * {@code zigZag}. Only these tokens may be written as {@code a:pattFill/@prst}.</p>
     *
     * @return the OOXML token, or {@code null} for the constants that name no pattern
     *         ({@link #NOT_DEFINED} and {@link #UNKNOWN})
     */
    public String getOoxml() {
        return ooxml;
    }

    /**
     * Returns the pattern style for an OOXML {@code prst} token.
     *
     * @param token a {@code ST_PresetPatternVal} value, or {@code null}
     * @return the matching constant, or {@link #NOT_DEFINED} if the token names no pattern
     */
    public static PatternStyle fromOoxml(String token) {
        if (token == null || token.isEmpty()) {
            return NOT_DEFINED;
        }
        for (PatternStyle style : values()) {
            if (token.equals(style.ooxml)) {
                return style;
            }
        }
        return NOT_DEFINED;
    }
}
