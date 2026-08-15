package org.aspose.slides.foss;

import org.aspose.slides.foss.drawing.PointF;
import org.aspose.slides.foss.drawing.SizeF;
import org.aspose.slides.foss.export.ISaveOptions;
import org.aspose.slides.foss.export.SaveFormat;
import org.aspose.slides.foss.internal.export.ExporterBase;
import org.aspose.slides.foss.internal.export.ExporterRegistry;
import org.aspose.slides.foss.internal.pptx.OpcPackage;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import org.aspose.slides.foss.internal.pptx.LayoutSlidePart;
import org.aspose.slides.foss.internal.pptx.MasterSlidePart;
import org.aspose.slides.foss.internal.pptx.PresentationPart;
import org.aspose.slides.foss.internal.pptx.RelsHelper;
import org.aspose.slides.foss.internal.pptx.ThreadedCommentsPart;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/**
 * Represents a PowerPoint presentation.
 *
 * <p>Implements {@link AutoCloseable} for use with try-with-resources.</p>
 */
public final class Presentation implements IPresentation {

    private static final String NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static final String NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

    private final OpcPackage pkg;
    private final DocumentProperties documentProperties;
    private final CommentAuthorCollection commentAuthors;
    private final SlideCollection slides;
    private final ImageCollection imageCollection;
    private final NotesSize notesSize;
    private final MasterSlideCollection masters;
    private final GlobalLayoutSlideCollection globalLayoutSlides;
    private LocalDateTime currentDateTime;
    private SourceFormat sourceFormat;
    private int firstSlideNumber = 1;

    // Lazily initialized maps for master/layout slide resolution
    private Map<String, MasterSlide> masterSlidesMap;
    private Map<String, LayoutSlide> layoutSlidesMap;
    private List<IMasterSlide> parsedMasters;
    private List<ILayoutSlide> parsedLayouts;

    /**
     * Creates a new blank presentation with one slide.
     */
    public Presentation() {
        pkg = new OpcPackage();
        initBlankPackage();
        documentProperties = new DocumentProperties();
        documentProperties.initInternal(pkg, this);
        commentAuthors = new CommentAuthorCollection();
        imageCollection = new ImageCollection(pkg);
        masters = new MasterSlideCollection();
        globalLayoutSlides = new GlobalLayoutSlideCollection();
        initDefaultMasterAndLayout();
        slides = new SlideCollection();
        slides.initInternal(this, pkg, new PresentationPart(pkg), this::resolveLayoutSlide);
        for (Slide slide : slides.getInternalList()) {
            if (slide.getLayoutSlide() == null) {
                slide.setLayoutSlide(globalLayoutSlides.get(0));
            }
        }
        notesSize = new NotesSize(new SizeF(540f, 720f));
        currentDateTime = LocalDateTime.now();
        sourceFormat = SourceFormat.PPTX;
    }

    /**
     * Opens a presentation from the given file path.
     *
     * @param path the file path
     * @throws IOException if an I/O error occurs
     */
    public Presentation(String path) throws IOException {
        this(Objects.requireNonNull(path, "path"),
                Files.newInputStream(Path.of(path)));
    }

    private Presentation(String path, InputStream in) throws IOException {
        this(in);
        detectSourceFormat(path);
    }

    /**
     * Opens a presentation from the given input stream.
     *
     * @param in the input stream
     * @throws IOException if an I/O error occurs
     */
    public Presentation(InputStream in) throws IOException {
        Objects.requireNonNull(in, "in");
        pkg = new OpcPackage();
        try (in) {
            pkg.load(in);
        }
        documentProperties = new DocumentProperties();
        documentProperties.initInternal(pkg, this);
        commentAuthors = new CommentAuthorCollection();
        imageCollection = new ImageCollection(pkg);
        imageCollection.loadFromPackage();
        masters = new MasterSlideCollection();
        globalLayoutSlides = new GlobalLayoutSlideCollection();
        if (!loadMastersAndLayouts()) {
            initDefaultMasterAndLayout();
        }
        slides = new SlideCollection();
        loadSlides();
        notesSize = new NotesSize(new SizeF(540f, 720f));
        loadCommentAuthors();
        loadComments();
        loadFirstSlideNumber();
        currentDateTime = LocalDateTime.now();
        sourceFormat = SourceFormat.PPTX;
    }

    /**
     * Populates {@link #getMasters()} and {@link #getLayoutSlides()} from the
     * masters and layouts the package actually contains.
     *
     * <p>Without this the load path fabricated one synthetic master and one
     * synthetic layout named "Blank" and published those instead, so
     * {@code addEmptySlide(pres.getLayoutSlides().get(0))} — the documented
     * form — always passed a layout that is not in the document, and nothing
     * could inherit placeholders or a theme from it. The parser that reads the
     * real ones already existed; only {@code Slide.resolveLayoutSlide} reached
     * it.</p>
     *
     * @return {@code true} if the package declared at least one master, so that
     *         the blank-presentation fallback is used only when it has none
     */
    private boolean loadMastersAndLayouts() {
        ensureLayoutSlidesParsed();
        if (parsedMasters.isEmpty()) {
            return false;
        }
        for (IMasterSlide master : parsedMasters) {
            masters.add(master);
        }
        globalLayoutSlides.initInternal(new ArrayList<>(parsedLayouts));
        return true;
    }

    private void initDefaultMasterAndLayout() {
        var defaultLayout = new LayoutSlide();
        defaultLayout.setLayoutType(SlideLayoutType.BLANK);
        defaultLayout.setName("Blank");

        var masterLayoutCollection = new MasterLayoutSlideCollection();
        masterLayoutCollection.getInternalList().add(defaultLayout);

        var masterSlide = new MasterSlide(masterLayoutCollection);
        defaultLayout.setMasterSlide(masterSlide);

        masters.add(masterSlide);
        globalLayoutSlides.add(defaultLayout);
    }

    // ---- Blank package initialization ----

    private void initBlankPackage() {
        var utf8 = java.nio.charset.StandardCharsets.UTF_8;

        pkg.setPartBytes("[Content_Types].xml", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>" +
                "<Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>" +
                "<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>" +
                "<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>" +
                "<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" +
                "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" +
                "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                "</Types>"
        ).getBytes(utf8));

        pkg.setPartBytes("_rels/.rels", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>" +
                "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
                "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                "</Relationships>"
        ).getBytes(utf8));

        pkg.setPartBytes("ppt/presentation.xml", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>" +
                "<p:sldIdLst><p:sldId id=\"256\" r:id=\"rId3\"/></p:sldIdLst>" +
                "<p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"/>" +
                "<p:notesSz cx=\"6858000\" cy=\"9144000\"/>" +
                "</p:presentation>"
        ).getBytes(utf8));

        pkg.setPartBytes("ppt/_rels/presentation.xml.rels", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>" +
                "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>" +
                "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>" +
                "</Relationships>"
        ).getBytes(utf8));

        pkg.setPartBytes("ppt/slides/slide1.xml", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:cSld><p:spTree>" +
                "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
                "<p:grpSpPr/></p:spTree></p:cSld></p:sld>"
        ).getBytes(utf8));

        pkg.setPartBytes("ppt/slides/_rels/slide1.xml.rels", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
                "</Relationships>"
        ).getBytes(utf8));

        pkg.setPartBytes("ppt/slideMasters/slideMaster1.xml", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:sldMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:cSld><p:spTree>" +
                "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
                "<p:grpSpPr/></p:spTree></p:cSld>" +
                "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" " +
                "accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" " +
                "accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>" +
                "<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst>" +
                "</p:sldMaster>"
        ).getBytes(utf8));

        pkg.setPartBytes("ppt/slideMasters/_rels/slideMaster1.xml.rels", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
                "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>" +
                "</Relationships>"
        ).getBytes(utf8));

        pkg.setPartBytes("ppt/slideLayouts/slideLayout1.xml", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:sldLayout xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                "type=\"blank\" preserve=\"1\">" +
                "<p:cSld name=\"Blank\"><p:spTree>" +
                "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
                "<p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"
        ).getBytes(utf8));

        pkg.setPartBytes("ppt/slideLayouts/_rels/slideLayout1.xml.rels", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/>" +
                "</Relationships>"
        ).getBytes(utf8));

        pkg.setPartBytes("ppt/theme/theme1.xml", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">" +
                "<a:themeElements>" +
                "<a:clrScheme name=\"Office\">" +
                "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>" +
                "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>" +
                "<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2>" +
                "<a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>" +
                "<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1>" +
                "<a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>" +
                "<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3>" +
                "<a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>" +
                "<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5>" +
                "<a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>" +
                "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>" +
                "<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>" +
                "</a:clrScheme>" +
                "<a:fontScheme name=\"Office\">" +
                "<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
                "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>" +
                "</a:fontScheme>" +
                "<a:fmtScheme name=\"Office\">" +
                "<a:fillStyleLst>" +
                "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                "</a:fillStyleLst>" +
                "<a:lnStyleLst>" +
                "<a:ln w=\"6350\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>" +
                "<a:ln w=\"12700\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>" +
                "<a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>" +
                "</a:lnStyleLst>" +
                "<a:effectStyleLst>" +
                "<a:effectStyle><a:effectLst/></a:effectStyle>" +
                "<a:effectStyle><a:effectLst/></a:effectStyle>" +
                "<a:effectStyle><a:effectLst/></a:effectStyle>" +
                "</a:effectStyleLst>" +
                "<a:bgFillStyleLst>" +
                "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                "</a:bgFillStyleLst>" +
                "</a:fmtScheme>" +
                "</a:themeElements>" +
                "</a:theme>"
        ).getBytes(utf8));

        pkg.setPartBytes("docProps/core.xml", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" " +
                "xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" " +
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>"
        ).getBytes(utf8));

        pkg.setPartBytes("docProps/app.xml", (
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"/>"
        ).getBytes(utf8));
    }

    // ---- Load helpers ----

    /**
     * Builds the slide collection from {@code ppt/presentation.xml}.
     *
     * <p>The slides of a presentation are the entries of {@code p:sldIdLst}, in that order,
     * each resolved through a relationship. Counting {@code ppt/slides/slideN.xml} parts
     * instead reports slides that no reader can see, and reports them in file-name order
     * rather than presentation order.</p>
     */
    private void loadSlides() {
        if (pkg.hasPart(PresentationPart.PART_NAME)) {
            slides.initInternal(this, pkg, new PresentationPart(pkg), this::resolveLayoutSlide);
        } else {
            // Not a presentation package; leave the collection empty rather than invent a slide.
            slides.setPresentation(this);
        }
    }

    private void loadFirstSlideNumber() {
        var presPart = new PresentationPart(pkg);
        this.firstSlideNumber = presPart.getFirstSlideNumber();
    }

    private void loadCommentAuthors() {
        Document doc = pkg.parseXml("ppt/commentAuthors.xml");
        if (doc == null) return;
        NodeList authorNodes = doc.getElementsByTagNameNS(NS_P, "cmAuthor");
        int maxId = -1;
        for (int i = 0; i < authorNodes.getLength(); i++) {
            Element el = (Element) authorNodes.item(i);
            String name = el.getAttribute("name");
            String initials = el.getAttribute("initials");
            int id = Integer.parseInt(el.getAttribute("id"));
            if (id > maxId) maxId = id;
            var author = new CommentAuthor(name, initials, id);
            commentAuthors.getInternalList().add(author);
        }
        commentAuthors.setNextId(maxId + 1);
    }

    private void loadComments() {
        for (int slideIdx = 0; slideIdx < slides.size(); slideIdx++) {
            Slide slide = slides.getInternalList().get(slideIdx);
            int slideNumber = slideFileNumberOf(slide.getSlidePartUri(), slideIdx + 1);
            String partUri = "ppt/comments/comment" + slideNumber + ".xml";
            Document doc = pkg.parseXml(partUri);
            if (doc == null) continue;
            var slideComments = new ArrayList<Comment>();
            NodeList cmNodes = doc.getElementsByTagNameNS(NS_P, "cm");
            for (int i = 0; i < cmNodes.getLength(); i++) {
                Element cmEl = (Element) cmNodes.item(i);
                int authorId = Integer.parseInt(cmEl.getAttribute("authorId"));
                String dtStr = cmEl.getAttribute("dt");
                LocalDateTime dt = parseDateTime(dtStr);

                // Parse position (stored as EMUs in XML, exposed as centimeters)
                float x = 0, y = 0;
                NodeList posNodes = cmEl.getElementsByTagNameNS(NS_P, "pos");
                if (posNodes.getLength() > 0) {
                    Element posEl = (Element) posNodes.item(0);
                    x = Integer.parseInt(posEl.getAttribute("x")) / (float) CM_TO_EMU;
                    y = Integer.parseInt(posEl.getAttribute("y")) / (float) CM_TO_EMU;
                }

                // Parse text
                String text = "";
                NodeList textNodes = cmEl.getElementsByTagNameNS(NS_P, "text");
                if (textNodes.getLength() > 0) {
                    text = textNodes.item(0).getTextContent();
                }

                // Find the author
                CommentAuthor author = findAuthorById(authorId);
                if (author != null) {
                    slideComments.add((Comment) author.getCommentCollection()
                            .addComment(text, slide, new PointF(x, y), dt));
                }
            }
            restoreCommentThreads(slideNumber, slideComments);
        }
    }

    /**
     * Restores the reply relationships recorded in a slide's threaded-comment part.
     *
     * <p>The classic {@code <p:cm>} carries no parent, so without this a deck
     * saved with threads reads back as unrelated comments — the same loss the
     * save path used to have.</p>
     *
     * @param slideNumber   the slide part number
     * @param slideComments the slide's comments, in the order the classic part lists them
     */
    private void restoreCommentThreads(int slideNumber, List<Comment> slideComments) {
        List<Integer> parents = ThreadedCommentsPart.readParentIndices(pkg, slideNumber);
        if (parents.size() != slideComments.size()) {
            return;
        }
        for (int i = 0; i < parents.size(); i++) {
            int parentIndex = parents.get(i);
            if (parentIndex >= 0 && parentIndex < slideComments.size() && parentIndex != i) {
                slideComments.get(i).setParentComment(slideComments.get(parentIndex));
            }
        }
    }

    private CommentAuthor findAuthorById(int id) {
        for (CommentAuthor a : commentAuthors.getInternalList()) {
            if (a.getId() == id) return a;
        }
        return null;
    }

    private static LocalDateTime parseDateTime(String s) {
        if (s == null || s.isBlank()) return LocalDateTime.now();
        try {
            return LocalDateTime.parse(s, DateTimeFormatter.ISO_DATE_TIME);
        } catch (DateTimeParseException e) {
            return LocalDateTime.now();
        }
    }

    // ---- Source format detection ----

    /**
     * Detects source format from a file extension.
     *
     * @param path the file path
     */
    private void detectSourceFormat(String path) {
        String lower = path.toLowerCase(java.util.Locale.ROOT);
        if (lower.endsWith(".pptx") || lower.endsWith(".pptm")
                || lower.endsWith(".ppsx") || lower.endsWith(".potx")) {
            this.sourceFormat = SourceFormat.PPTX;
        } else if (lower.endsWith(".ppt")) {
            this.sourceFormat = SourceFormat.PPT;
        } else if (lower.endsWith(".odp")) {
            this.sourceFormat = SourceFormat.ODP;
        }
    }

    // ---- Lazy master/layout slide parsing ----

    /**
     * Parses all master slides and layout slides from the package on first access.
     *
     * <p>This method is idempotent; subsequent calls return immediately once
     * the master and layout slide maps have been populated.</p>
     */
    void ensureLayoutSlidesParsed() {
        if (layoutSlidesMap != null) {
            return;
        }

        masterSlidesMap = new HashMap<>();
        layoutSlidesMap = new HashMap<>();
        parsedMasters = new ArrayList<>();
        parsedLayouts = new ArrayList<>();

        var presRels = new RelsHelper(pkg, PresentationPart.PART_NAME);
        var presPart = new PresentationPart(pkg);

        for (var masterRef : presPart.masterReferences()) {
            // Find the relationship by ID
            RelsHelper.RelEntry rel = null;
            for (var entry : presRels.getAllRelationships()) {
                if (entry.id().equals(masterRef.getRelId())) {
                    rel = entry;
                    break;
                }
            }
            if (rel == null) {
                continue;
            }

            // Resolve the master part name
            String target = rel.target();
            String masterPartName;
            if (target.startsWith("/")) {
                masterPartName = target.replaceFirst("^/+", "");
            } else {
                masterPartName = "ppt/" + target;
            }

            var masterPart = new MasterSlidePart(pkg, masterPartName);

            // Parse layout slides for this master
            List<ILayoutSlide> masterLayouts = new ArrayList<>();
            for (String layoutPartName : masterPart.getLayoutPartNames()) {
                if (layoutSlidesMap.containsKey(layoutPartName)) {
                    masterLayouts.add(layoutSlidesMap.get(layoutPartName));
                    continue;
                }

                var layoutPart = new LayoutSlidePart(pkg, layoutPartName);
                var layoutSlide = new LayoutSlide();
                layoutSlide.initInternal(
                        this, pkg, layoutPartName, layoutPart, this::resolveMasterSlide);
                layoutSlidesMap.put(layoutPartName, layoutSlide);
                parsedLayouts.add(layoutSlide);
                masterLayouts.add(layoutSlide);
            }

            // Create master slide object
            var masterSlide = new MasterSlide();
            masterSlide.initInternal(this, pkg, masterPartName, masterPart, masterLayouts);
            masterSlidesMap.put(masterPartName, masterSlide);
            parsedMasters.add(masterSlide);
        }
    }

    /**
     * Resolves a layout slide part name to a {@link LayoutSlide} object.
     *
     * @param partName the layout slide part name
     * @return the resolved layout slide, or {@code null} if not found
     */
    LayoutSlide resolveLayoutSlide(String partName) {
        ensureLayoutSlidesParsed();
        return layoutSlidesMap.get(partName);
    }

    /**
     * Resolves a master slide part name to a {@link MasterSlide} object.
     *
     * @param partName the master slide part name
     * @return the resolved master slide, or {@code null} if not found
     */
    MasterSlide resolveMasterSlide(String partName) {
        ensureLayoutSlidesParsed();
        return masterSlidesMap.get(partName);
    }

    // ---- Save ----

    @Override
    public void save(String path) throws IOException {
        Objects.requireNonNull(path, "path");
        try (var out = Files.newOutputStream(Path.of(path))) {
            save(out);
        }
    }

    @Override
    public void save(OutputStream stream) throws IOException {
        Objects.requireNonNull(stream, "stream");
        flushToPackage();
        pkg.save(stream);
    }

    /**
     * Writes every pending in-memory change into the OPC package.
     *
     * <p>Called before each save so that the package is complete whichever exporter
     * then serializes it.</p>
     */
    private void flushToPackage() {
        for (Slide slide : slides.getInternalList()) {
            slide.flush();
        }
        saveCommentAuthors();
        saveComments();
        documentProperties.save();
        saveFirstSlideNumber();
    }

    /**
     * Returns the exporter for a save format, or fails if there is none.
     *
     * <p>A format with no exporter is not written under another format's bytes: the caller
     * is told, before anything is created, that the conversion did not happen.</p>
     *
     * @param format the requested save format
     * @return the exporter for that format
     * @throws UnsupportedOperationException if no exporter can write the format
     */
    private static ExporterBase exporterFor(SaveFormat format) {
        Objects.requireNonNull(format, "format");
        return ExporterRegistry.getExporter(format.getValue())
                .orElseThrow(() -> new UnsupportedOperationException(
                        "Export format '" + format.getValue() + "' is not supported"));
    }

    private void saveFirstSlideNumber() {
        var presPart = new PresentationPart(pkg);
        presPart.setFirstSlideNumber(firstSlideNumber);
        presPart.save();
    }

    private void saveCommentAuthors() {
        if (commentAuthors.isEmpty()) return;
        Document doc = OpcPackage.newDocument();
        Element root = doc.createElementNS(NS_P, "p:cmAuthorLst");
        doc.appendChild(root);
        for (CommentAuthor author : commentAuthors.getInternalList()) {
            Element el = doc.createElementNS(NS_P, "p:cmAuthor");
            el.setAttribute("id", String.valueOf(author.getId()));
            el.setAttribute("name", author.getName());
            el.setAttribute("initials", author.getInitials());
            el.setAttribute("lastIdx", String.valueOf(author.getCommentCollection().size()));
            el.setAttribute("clrIdx", String.valueOf(author.getId()));
            root.appendChild(el);
        }
        pkg.serializeXml("ppt/commentAuthors.xml", doc);

        // Register content type
        var ctm = new org.aspose.slides.foss.internal.pptx.ContentTypesManager(pkg);
        ctm.addOverride("/ppt/commentAuthors.xml",
                org.aspose.slides.foss.internal.pptx.ContentTypesManager.CONTENT_TYPES.get("commentAuthors"));
        ctm.save();

        // Add relationship if not present
        addPresentationRelationship(
                "commentAuthors",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
                "commentAuthors.xml"
        );
    }

    /** Centimeters-to-EMU conversion factor for comment positions. */
    private static final int CM_TO_EMU = 360000;

    private void saveComments() {
        for (int slideIdx = 0; slideIdx < slides.size(); slideIdx++) {
            Slide slide = slides.getInternalList().get(slideIdx);
            // Collect all comments for this slide
            var slideComments = new java.util.ArrayList<Comment>();
            for (CommentAuthor author : commentAuthors.getInternalList()) {
                for (Comment c : author.getCommentCollection().getInternalList()) {
                    if (c.getSlide() == slide) {
                        slideComments.add(c);
                    }
                }
            }
            // Name the comments part after the slide part that owns it, not after the
            // slide's position: the two differ once slides have been added or removed.
            String slidePartUri = slide.getSlidePartUri();
            int slideNumber = slideFileNumberOf(slidePartUri, slideIdx + 1);
            String partUri = "ppt/comments/comment" + slideNumber + ".xml";
            if (slideComments.isEmpty()) {
                // The relationship and the content-type Override go with the part.
                pkg.removePartCascading(partUri, slidePartUri);
                ThreadedCommentsPart.delete(pkg, slidePartUri, slideNumber);
                continue;
            }
            Document doc = OpcPackage.newDocument();
            Element root = doc.createElementNS(NS_P, "p:cmLst");
            doc.appendChild(root);
            int idx = 1;
            for (Comment c : slideComments) {
                Element cmEl = doc.createElementNS(NS_P, "p:cm");
                CommentAuthor author = (CommentAuthor) c.getAuthor();
                cmEl.setAttribute("authorId", String.valueOf(author.getId()));
                cmEl.setAttribute("idx", String.valueOf(idx++));
                if (c.getCreatedTime() != null) {
                    cmEl.setAttribute("dt", c.getCreatedTime().format(DateTimeFormatter.ISO_DATE_TIME));
                }
                Element posEl = doc.createElementNS(NS_P, "p:pos");
                posEl.setAttribute("x", String.valueOf(Math.round(c.getPosition().getX() * CM_TO_EMU)));
                posEl.setAttribute("y", String.valueOf(Math.round(c.getPosition().getY() * CM_TO_EMU)));
                cmEl.appendChild(posEl);
                Element textEl = doc.createElementNS(NS_P, "p:text");
                textEl.setTextContent(c.getText());
                cmEl.appendChild(textEl);
                root.appendChild(cmEl);
            }
            pkg.serializeXml(partUri, doc);

            // Register content type
            var ctm = new org.aspose.slides.foss.internal.pptx.ContentTypesManager(pkg);
            ctm.addOverride("/" + partUri,
                    org.aspose.slides.foss.internal.pptx.ContentTypesManager.CONTENT_TYPES.get("comments"));
            ctm.save();

            // Add slide relationship to comments
            addSlideCommentRelationship(slideNumber);

            // A reply has nowhere to live on the classic <p:cm>, so the threads
            // are written to the modern part beside it.
            saveThreadedComments(slidePartUri, slideNumber, slideComments);
        }
        saveCommentAuthorsForThreads();
    }

    /**
     * Writes the modern threaded-comment part for one slide.
     *
     * <p>{@code CT_Comment} has no attribute for a parent comment, so a reply
     * recorded only on {@code <p:cm>} is lost on save — which is what used to
     * happen: {@code setParentComment} assigned a field and reached no part of
     * the package, and a reviewer's discussion opened as unrelated comments.</p>
     *
     * @param slidePartUri  the part name of the slide
     * @param slideNumber   the slide part number
     * @param slideComments the slide's comments, in the order the classic part
     *                      writes them
     */
    private void saveThreadedComments(String slidePartUri, int slideNumber,
                                      List<Comment> slideComments) {
        var entries = new ArrayList<ThreadedCommentsPart.Entry>(slideComments.size());
        for (Comment comment : slideComments) {
            var author = (CommentAuthor) comment.getAuthor();
            int parentIndex = comment.getParentComment()
                    .map(parent -> slideComments.indexOf((Comment) parent))
                    .orElse(-1);
            entries.add(new ThreadedCommentsPart.Entry(
                    ThreadedCommentsPart.authorGuid(author.getName(), author.getInitials()),
                    comment.getCreatedTime(),
                    comment.getText(),
                    Math.round(comment.getPosition().getX() * CM_TO_EMU),
                    Math.round(comment.getPosition().getY() * CM_TO_EMU),
                    parentIndex));
        }
        ThreadedCommentsPart.write(pkg, slidePartUri, slideNumber, entries);
    }

    /** Writes {@code ppt/authors.xml}, which the threaded-comment parts refer to. */
    private void saveCommentAuthorsForThreads() {
        boolean anyThreadedPart = pkg.getPartNames().stream()
                .anyMatch(name -> name.startsWith("ppt/threadedComments/"));
        if (!anyThreadedPart) {
            return;
        }
        var authors = new ArrayList<String[]>();
        for (CommentAuthor author : commentAuthors.getInternalList()) {
            authors.add(new String[]{author.getName(), author.getInitials()});
        }
        ThreadedCommentsPart.writeAuthors(pkg, authors);
    }

    /**
     * Extracts the {@code N} of {@code ppt/slides/slideN.xml}.
     *
     * @param slidePartUri the slide's part name, which may be {@code null}
     * @param fallback     the number to use when the part name does not carry one
     * @return the slide part number
     */
    private static int slideFileNumberOf(String slidePartUri, int fallback) {
        if (slidePartUri == null) {
            return fallback;
        }
        var matcher = java.util.regex.Pattern
                .compile("ppt/slides/slide(\\d+)\\.xml").matcher(slidePartUri);
        return matcher.matches() ? Integer.parseInt(matcher.group(1)) : fallback;
    }

    private void addPresentationRelationship(String idSuffix, String type, String target) {
        String relsUri = "ppt/_rels/presentation.xml.rels";
        Document doc = pkg.parseXml(relsUri);
        if (doc == null) return;
        Element root = doc.getDocumentElement();
        // Check if relationship already exists
        NodeList rels = root.getElementsByTagName("Relationship");
        for (int i = 0; i < rels.getLength(); i++) {
            Element rel = (Element) rels.item(i);
            if (type.equals(rel.getAttribute("Type"))) return;
        }
        Element rel = doc.createElementNS(REL_NS, "Relationship");
        rel.setAttribute("Id", "rId_" + idSuffix);
        rel.setAttribute("Type", type);
        rel.setAttribute("Target", target);
        root.appendChild(rel);
        pkg.serializeXml(relsUri, doc);
    }

    private void addSlideCommentRelationship(int slideNumber) {
        String relsUri = "ppt/slides/_rels/slide" + slideNumber + ".xml.rels";
        String type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        Document doc = pkg.parseXml(relsUri);
        if (doc == null) {
            doc = OpcPackage.newDocument();
            Element root = doc.createElementNS(REL_NS, "Relationships");
            doc.appendChild(root);
            Element rel = doc.createElementNS(REL_NS, "Relationship");
            rel.setAttribute("Id", "rId_comments");
            rel.setAttribute("Type", type);
            rel.setAttribute("Target", "../comments/comment" + slideNumber + ".xml");
            root.appendChild(rel);
            pkg.serializeXml(relsUri, doc);
            return;
        }
        Element root = doc.getDocumentElement();
        NodeList rels = root.getElementsByTagName("Relationship");
        for (int i = 0; i < rels.getLength(); i++) {
            Element rel = (Element) rels.item(i);
            if (type.equals(rel.getAttribute("Type"))) return;
        }
        Element rel = doc.createElementNS(REL_NS, "Relationship");
        rel.setAttribute("Id", "rId_comments");
        rel.setAttribute("Type", type);
        rel.setAttribute("Target", "../comments/comment" + slideNumber + ".xml");
        root.appendChild(rel);
        pkg.serializeXml(relsUri, doc);
    }

    // ---- Public API ----

    @Override
    public IPresentation getPresentation() {
        return this;
    }

    @Override
    public LocalDateTime getCurrentDateTime() {
        return currentDateTime;
    }

    @Override
    public void setCurrentDateTime(LocalDateTime value) {
        this.currentDateTime = value;
    }

    @Override
    public IDocumentProperties getDocumentProperties() {
        return documentProperties;
    }

    @Override
    public ICommentAuthorCollection getCommentAuthors() {
        return commentAuthors;
    }

    @Override
    public ISlideCollection getSlides() {
        return slides;
    }

    @Override
    public INotesSize getNotesSize() {
        return notesSize;
    }

    @Override
    public IGlobalLayoutSlideCollection getLayoutSlides() {
        return globalLayoutSlides;
    }

    @Override
    public IMasterSlideCollection getMasters() {
        return masters;
    }

    @Override
    public IImageCollection getImages() {
        return imageCollection;
    }

    @Override
    public SourceFormat getSourceFormat() {
        return sourceFormat;
    }

    @Override
    public int getFirstSlideNumber() {
        return firstSlideNumber;
    }

    @Override
    public void setFirstSlideNumber(int value) {
        this.firstSlideNumber = value;
    }

    @Override
    public IPresentationComponent asIPresentationComponent() {
        return this;
    }

    @Override
    public void save(String path, SaveFormat format) throws IOException {
        save(path, format, null);
    }

    @Override
    public void save(OutputStream stream, SaveFormat format) throws IOException {
        save(stream, format, null);
    }

    @Override
    public void save(String path, SaveFormat format, ISaveOptions options) throws IOException {
        Objects.requireNonNull(path, "path");
        ExporterBase exporter = exporterFor(format);
        flushToPackage();
        exporter.exportToPath(pkg, path, options);
    }

    @Override
    public void save(OutputStream stream, SaveFormat format, ISaveOptions options) throws IOException {
        Objects.requireNonNull(stream, "stream");
        ExporterBase exporter = exporterFor(format);
        flushToPackage();
        exporter.exportToStream(pkg, stream, options);
    }

    @Override
    public void save(String path, int[] slides, SaveFormat format) throws IOException {
        save(path, slides, format, null);
    }

    @Override
    public void save(String path, int[] slides, SaveFormat format, ISaveOptions options) throws IOException {
        Objects.requireNonNull(path, "path");
        if (slides == null) {
            save(path, format, options);
            return;
        }
        try (Presentation subset = subsetOf(slides)) {
            subset.save(path, format, options);
        }
    }

    @Override
    public void save(OutputStream stream, int[] slides, SaveFormat format) throws IOException {
        save(stream, slides, format, null);
    }

    @Override
    public void save(OutputStream stream, int[] slides, SaveFormat format, ISaveOptions options) throws IOException {
        Objects.requireNonNull(stream, "stream");
        if (slides == null) {
            save(stream, format, options);
            return;
        }
        try (Presentation subset = subsetOf(slides)) {
            subset.save(stream, format, options);
        }
    }

    /**
     * Returns a presentation holding only the requested slides, in document order.
     *
     * <p>The subset overloads used to discard this argument and write the whole
     * deck, reporting success — a caller who asked for one slide got all of them
     * with no signal that anything had happened.</p>
     *
     * <p>Indices are positions in {@link #getSlides()}, zero-based. Repeats are
     * ignored; the slides that are kept stay in the order the document has them,
     * not the order they were named in.</p>
     *
     * @param slideIndices the positions of the slides to keep
     * @return a new presentation containing only those slides; the caller closes it
     * @throws IndexOutOfBoundsException if an index is not a slide of this presentation
     * @throws IllegalArgumentException  if no slides are requested
     * @throws IOException               if the intermediate package cannot be written or read
     */
    private Presentation subsetOf(int[] slideIndices) throws IOException {
        int total = slides.size();
        if (slideIndices.length == 0) {
            throw new IllegalArgumentException(
                    "No slides were requested; a subset save must name at least one slide");
        }
        var keep = new java.util.LinkedHashSet<Integer>();
        for (int index : slideIndices) {
            if (index < 0 || index >= total) {
                throw new IndexOutOfBoundsException("Slide index " + index
                        + " is out of range; this presentation has " + total + " slides");
            }
            keep.add(index);
        }

        flushToPackage();
        var buffer = new java.io.ByteArrayOutputStream();
        pkg.save(buffer);
        var subset = new Presentation(new java.io.ByteArrayInputStream(buffer.toByteArray()));
        try {
            for (int index = total - 1; index >= 0; index--) {
                if (!keep.contains(index)) {
                    subset.getSlides().removeAt(index);
                }
            }
        } catch (RuntimeException e) {
            subset.close();
            throw e;
        }
        return subset;
    }

    @Override
    public void save(ISaveOptions options) throws IOException {
        throw new UnsupportedOperationException("save(ISaveOptions) requires an output path or stream");
    }

    /**
     * Returns the underlying OPC package (for internal use).
     *
     * @return the OPC package
     */
    OpcPackage getPackage() {
        return pkg;
    }

    /**
     * Returns the concrete slide collection (for internal use).
     *
     * @return the slide collection
     */
    SlideCollection getSlideCollection() {
        return slides;
    }

    /**
     * Returns the concrete comment author collection (for internal use).
     *
     * @return the comment author collection
     */
    CommentAuthorCollection getCommentAuthorCollection() {
        return commentAuthors;
    }

    @Override
    public void dispose() {
        masterSlidesMap = null;
        layoutSlidesMap = null;
        parsedMasters = null;
        parsedLayouts = null;
        pkg.clear();
    }

    @Override
    public void close() {
        dispose();
    }
}
