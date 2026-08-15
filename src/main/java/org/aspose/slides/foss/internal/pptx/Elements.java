package org.aspose.slides.foss.internal.pptx;

/**
 * Common PPTX element names with full namespace qualification.
 *
 * <p>Each constant is a fully-qualified element name in the form
 * {@code {namespace-uri}localName}, suitable for XML element lookups.</p>
 */
public final class Elements {

    private Elements() {
        // utility class
    }

    // ── Presentation elements ──────────────────────────────────────────

    /** {@code <p:presentation>} */
    public static final String PRESENTATION = Ns.P + "presentation";

    /** {@code <p:sldIdLst>} */
    public static final String SLD_ID_LST = Ns.P + "sldIdLst";

    /** {@code <p:sldId>} */
    public static final String SLD_ID = Ns.P + "sldId";

    /** {@code <p:sldMasterIdLst>} */
    public static final String SLD_MASTER_ID_LST = Ns.P + "sldMasterIdLst";

    /** {@code <p:sldMasterId>} */
    public static final String SLD_MASTER_ID = Ns.P + "sldMasterId";

    /** {@code <p:sldSz>} */
    public static final String SLD_SZ = Ns.P + "sldSz";

    /** {@code <p:notesSz>} */
    public static final String NOTES_SZ = Ns.P + "notesSz";

    // ── Slide elements ─────────────────────────────────────────────────

    /** {@code <p:sld>} */
    public static final String SLD = Ns.P + "sld";

    /** {@code <p:cSld>} */
    public static final String C_SLD = Ns.P + "cSld";

    /** {@code <p:spTree>} */
    public static final String SP_TREE = Ns.P + "spTree";

    // ── Shape elements ─────────────────────────────────────────────────

    /** {@code <p:sp>} */
    public static final String SP = Ns.P + "sp";

    /** {@code <p:nvSpPr>} */
    public static final String NV_SP_PR = Ns.P + "nvSpPr";

    /** {@code <p:cNvPr>} */
    public static final String C_NV_PR = Ns.P + "cNvPr";

    /** {@code <p:spPr>} */
    public static final String SP_PR = Ns.P + "spPr";

    /** {@code <p:txBody>} */
    public static final String TX_BODY = Ns.P + "txBody";

    // ── DrawingML elements ─────────────────────────────────────────────

    /** {@code <a:p>} — paragraph. */
    public static final String A_P = Ns.A + "p";

    /** {@code <a:r>} — run. */
    public static final String A_R = Ns.A + "r";

    /** {@code <a:t>} — text. */
    public static final String A_T = Ns.A + "t";

    /** {@code <a:xfrm>} — transform. */
    public static final String A_XFRM = Ns.A + "xfrm";

    /** {@code <a:off>} — offset. */
    public static final String A_OFF = Ns.A + "off";

    /** {@code <a:ext>} — extents. */
    public static final String A_EXT = Ns.A + "ext";

    // ── Fill elements ──────────────────────────────────────────────────

    /** {@code <a:noFill>} */
    public static final String A_NO_FILL = Ns.A + "noFill";

    /** {@code <a:solidFill>} */
    public static final String A_SOLID_FILL = Ns.A + "solidFill";

    /** {@code <a:gradFill>} */
    public static final String A_GRAD_FILL = Ns.A + "gradFill";

    /** {@code <a:pattFill>} */
    public static final String A_PATT_FILL = Ns.A + "pattFill";

    /** {@code <a:blipFill>} */
    public static final String A_BLIP_FILL = Ns.A + "blipFill";

    /** {@code <a:grpFill>} */
    public static final String A_GRP_FILL = Ns.A + "grpFill";

    // ── Color elements ─────────────────────────────────────────────────

    /** {@code <a:srgbClr>} */
    public static final String A_SRGB_CLR = Ns.A + "srgbClr";

    /** {@code <a:schemeClr>} */
    public static final String A_SCHEME_CLR = Ns.A + "schemeClr";

    /** {@code <a:prstClr>} */
    public static final String A_PRST_CLR = Ns.A + "prstClr";

    /** {@code <a:sysClr>} */
    public static final String A_SYS_CLR = Ns.A + "sysClr";

    /** {@code <a:hlsClr>} */
    public static final String A_HLS_CLR = Ns.A + "hlsClr";

    /** {@code <a:scrgbClr>} */
    public static final String A_SCR_GB_CLR = Ns.A + "scrgbClr";

    // ── Gradient elements ──────────────────────────────────────────────

    /** {@code <a:gsLst>} */
    public static final String A_GS_LST = Ns.A + "gsLst";

    /** {@code <a:gs>} */
    public static final String A_GS = Ns.A + "gs";

    /** {@code <a:lin>} */
    public static final String A_LIN = Ns.A + "lin";

    /** {@code <a:path>} */
    public static final String A_PATH = Ns.A + "path";

    /** {@code <a:tileRect>} */
    public static final String A_TILE_RECT = Ns.A + "tileRect";

    // ── Pattern elements ───────────────────────────────────────────────

    /** {@code <a:fgClr>} */
    public static final String A_FG_CLR = Ns.A + "fgClr";

    /** {@code <a:bgClr>} */
    public static final String A_BG_CLR = Ns.A + "bgClr";

    // ── Text elements ──────────────────────────────────────────────────

    /** {@code <a:bodyPr>} */
    public static final String A_BODY_PR = Ns.A + "bodyPr";

    /** {@code <a:lstStyle>} */
    public static final String A_LST_STYLE = Ns.A + "lstStyle";

    /** {@code <a:rPr>} */
    public static final String A_R_PR = Ns.A + "rPr";

    /** {@code <a:pPr>} */
    public static final String A_P_PR = Ns.A + "pPr";

    /** {@code <a:endParaRPr>} */
    public static final String A_END_PARA_RPR = Ns.A + "endParaRPr";

    /** {@code <a:highlight>} */
    public static final String A_HIGHLIGHT = Ns.A + "highlight";

    /** {@code <a:uLnTx>} */
    public static final String A_U_LN_TX = Ns.A + "uLnTx";

    /** {@code <a:uLn>} */
    public static final String A_U_LN = Ns.A + "uLn";

    /** {@code <a:uFillTx>} */
    public static final String A_U_FILL_TX = Ns.A + "uFillTx";

    /** {@code <a:uFill>} */
    public static final String A_U_FILL = Ns.A + "uFill";

    /** {@code <a:defRPr>} */
    public static final String A_DEF_R_PR = Ns.A + "defRPr";

    /** {@code <a:tabLst>} */
    public static final String A_TAB_LST = Ns.A + "tabLst";

    /** {@code <a:lnSpc>} */
    public static final String A_LN_SPC = Ns.A + "lnSpc";

    /** {@code <a:spcBef>} */
    public static final String A_SPC_BEF = Ns.A + "spcBef";

    /** {@code <a:spcAft>} */
    public static final String A_SPC_AFT = Ns.A + "spcAft";

    /** {@code <a:spcPct>} */
    public static final String A_SPC_PCT = Ns.A + "spcPct";

    /** {@code <a:spcPts>} */
    public static final String A_SPC_PTS = Ns.A + "spcPts";

    /** {@code <a:buNone>} */
    public static final String A_BU_NONE = Ns.A + "buNone";

    /** {@code <a:buChar>} */
    public static final String A_BU_CHAR = Ns.A + "buChar";

    /** {@code <a:buAutoNum>} */
    public static final String A_BU_AUTO_NUM = Ns.A + "buAutoNum";

    /** {@code <a:buFont>} */
    public static final String A_BU_FONT = Ns.A + "buFont";

    /** {@code <a:buSzPct>} */
    public static final String A_BU_SZ_PCT = Ns.A + "buSzPct";

    /** {@code <a:buSzPts>} */
    public static final String A_BU_SZ_PTS = Ns.A + "buSzPts";

    /** {@code <a:buClr>} */
    public static final String A_BU_CLR = Ns.A + "buClr";

    /** {@code <a:buClrTx>} */
    public static final String A_BU_CLR_TX = Ns.A + "buClrTx";

    /** {@code <a:buFontTx>} */
    public static final String A_BU_FONT_TX = Ns.A + "buFontTx";

    /** {@code <a:buSzTx>} */
    public static final String A_BU_SZ_TX = Ns.A + "buSzTx";

    /** {@code <a:buBlip>} */
    public static final String A_BU_BLIP = Ns.A + "buBlip";

    /** {@code <a:latin>} */
    public static final String A_LATIN = Ns.A + "latin";

    /** {@code <a:ea>} */
    public static final String A_EA = Ns.A + "ea";

    /** {@code <a:cs>} */
    public static final String A_CS = Ns.A + "cs";

    /** {@code <a:sym>} */
    public static final String A_SYM = Ns.A + "sym";

    // ── Autofit elements ───────────────────────────────────────────────

    /** {@code <a:noAutofit>} */
    public static final String A_NO_AUTOFIT = Ns.A + "noAutofit";

    /** {@code <a:spAutoFit>} */
    public static final String A_SP_AUTO_FIT = Ns.A + "spAutoFit";

    /** {@code <a:normAutofit>} */
    public static final String A_NORM_AUTOFIT = Ns.A + "normAutofit";

    // ── Text warp element ──────────────────────────────────────────────

    /** {@code <a:prstTxWarp>} */
    public static final String A_PRST_TX_WARP = Ns.A + "prstTxWarp";

    // ── Line elements ──────────────────────────────────────────────────

    /** {@code <a:ln>} */
    public static final String A_LN = Ns.A + "ln";

    /** {@code <a:prstDash>} */
    public static final String A_PRST_DASH = Ns.A + "prstDash";

    /** {@code <a:custDash>} */
    public static final String A_CUST_DASH = Ns.A + "custDash";

    /** {@code <a:round>} */
    public static final String A_ROUND = Ns.A + "round";

    /** {@code <a:bevel>} */
    public static final String A_BEVEL = Ns.A + "bevel";

    /** {@code <a:miter>} */
    public static final String A_MITER = Ns.A + "miter";

    /** {@code <a:headEnd>} */
    public static final String A_HEAD_END = Ns.A + "headEnd";

    /** {@code <a:tailEnd>} */
    public static final String A_TAIL_END = Ns.A + "tailEnd";

    // ── 3D elements ────────────────────────────────────────────────────

    /** {@code <a:scene3d>} */
    public static final String A_SCENE_3D = Ns.A + "scene3d";

    /** {@code <a:sp3d>} */
    public static final String A_SP_3D = Ns.A + "sp3d";

    /** {@code <a:camera>} */
    public static final String A_CAMERA = Ns.A + "camera";

    /** {@code <a:lightRig>} */
    public static final String A_LIGHT_RIG = Ns.A + "lightRig";

    /** {@code <a:bevelT>} */
    public static final String A_BEVEL_T = Ns.A + "bevelT";

    /** {@code <a:bevelB>} */
    public static final String A_BEVEL_B = Ns.A + "bevelB";

    /** {@code <a:contourClr>} */
    public static final String A_CONTOUR_CLR = Ns.A + "contourClr";

    /** {@code <a:extrusionClr>} */
    public static final String A_EXTRUSION_CLR = Ns.A + "extrusionClr";

    /** {@code <a:rot>} */
    public static final String A_ROT = Ns.A + "rot";

    /** {@code <a:effectLst>} */
    public static final String A_EFFECT_LST = Ns.A + "effectLst";

    /** {@code <a:effectDag>} */
    public static final String A_EFFECT_DAG = Ns.A + "effectDag";

    /** {@code <a:extLst>} */
    public static final String A_EXT_LST = Ns.A + "extLst";

    // ── Effect elements (children of effectLst, OOXML order) ───────────

    /** {@code <a:blur>} */
    public static final String A_BLUR = Ns.A + "blur";

    /** {@code <a:fillOverlay>} */
    public static final String A_FILL_OVERLAY = Ns.A + "fillOverlay";

    /** {@code <a:glow>} */
    public static final String A_GLOW = Ns.A + "glow";

    /** {@code <a:innerShdw>} */
    public static final String A_INNER_SHDW = Ns.A + "innerShdw";

    /** {@code <a:outerShdw>} */
    public static final String A_OUTER_SHDW = Ns.A + "outerShdw";

    /** {@code <a:prstShdw>} */
    public static final String A_PRST_SHDW = Ns.A + "prstShdw";

    /** {@code <a:reflection>} */
    public static final String A_REFLECTION = Ns.A + "reflection";

    /** {@code <a:softEdge>} */
    public static final String A_SOFT_EDGE = Ns.A + "softEdge";

    // ── Table elements ─────────────────────────────────────────────────

    /** {@code <a:graphic>} */
    public static final String A_GRAPHIC = Ns.A + "graphic";

    /** {@code <a:graphicData>} */
    public static final String A_GRAPHIC_DATA = Ns.A + "graphicData";

    /** {@code <a:tbl>} */
    public static final String A_TBL = Ns.A + "tbl";

    /** {@code <a:tblPr>} */
    public static final String A_TBL_PR = Ns.A + "tblPr";

    /** {@code <a:tblGrid>} */
    public static final String A_TBL_GRID = Ns.A + "tblGrid";

    /** {@code <a:gridCol>} */
    public static final String A_GRID_COL = Ns.A + "gridCol";

    /** {@code <a:tr>} */
    public static final String A_TR = Ns.A + "tr";

    /** {@code <a:tc>} */
    public static final String A_TC = Ns.A + "tc";

    /** {@code <a:tcPr>} */
    public static final String A_TC_PR = Ns.A + "tcPr";

    /** {@code <a:tblStyle>} */
    public static final String A_TBL_STYLE = Ns.A + "tblStyle";

    /** {@code <a:tableStyleId>} */
    public static final String A_TABLE_STYLE_ID = Ns.A + "tableStyleId";

    // ── Table border elements ──────────────────────────────────────────

    /** {@code <a:lnL>} */
    public static final String A_LN_L = Ns.A + "lnL";

    /** {@code <a:lnR>} */
    public static final String A_LN_R = Ns.A + "lnR";

    /** {@code <a:lnT>} */
    public static final String A_LN_T = Ns.A + "lnT";

    /** {@code <a:lnB>} */
    public static final String A_LN_B = Ns.A + "lnB";

    /** {@code <a:lnTlToBr>} */
    public static final String A_LN_TL_TO_BR = Ns.A + "lnTlToBr";

    /** {@code <a:lnBlToTr>} */
    public static final String A_LN_BL_TO_TR = Ns.A + "lnBlToTr";

    // ── GraphicFrame elements ──────────────────────────────────────────

    /** {@code <p:graphicFrame>} */
    public static final String P_GRAPHIC_FRAME = Ns.P + "graphicFrame";

    /** {@code <p:nvGraphicFramePr>} */
    public static final String P_NV_GRAPHIC_FRAME_PR = Ns.P + "nvGraphicFramePr";

    /** {@code <p:cNvGraphicFramePr>} */
    public static final String P_C_NV_GRAPHIC_FRAME_PR = Ns.P + "cNvGraphicFramePr";

    /** {@code <a:graphicFrameLocks>} */
    public static final String A_GRAPHIC_FRAME_LOCKS = Ns.A + "graphicFrameLocks";

    /** {@code <p:nvPr>} */
    public static final String P_NV_PR = Ns.P + "nvPr";

    /** {@code <p:xfrm>} */
    public static final String P_XFRM = Ns.P + "xfrm";

    // ── Table URI ──────────────────────────────────────────────────────

    /** URI for DrawingML table graphic data. */
    public static final String TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

    // ── DrawingML txBody (for table cells) ─────────────────────────────

    /** {@code <a:txBody>} */
    public static final String A_TX_BODY = Ns.A + "txBody";
}
