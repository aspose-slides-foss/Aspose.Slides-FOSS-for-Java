package org.aspose.slides.foss;

import org.aspose.slides.foss.internal.pptx.AppPropertiesPart;
import org.aspose.slides.foss.internal.pptx.CorePropertiesPart;
import org.aspose.slides.foss.internal.pptx.CustomPropertiesPart;
import org.aspose.slides.foss.internal.pptx.HeadingPairData;
import org.aspose.slides.foss.internal.pptx.OpcPackage;

import java.lang.ref.WeakReference;
import java.time.Duration;
import java.util.Objects;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZoneOffset;
import java.util.ArrayList;
import java.util.List;

/**
 * Represents properties of a presentation.
 *
 * <p>Lazily loads core, app, and custom property parts from the underlying OPC package.</p>
 */
public final class DocumentProperties implements IDocumentProperties {

    private OpcPackage pkg;
    private WeakReference<Presentation> presentationRef;
    private CorePropertiesPart corePart;
    private AppPropertiesPart appPart;
    private CustomPropertiesPart customPart;

    /**
     * Initializes internal state with the given OPC package.
     *
     * @param pkg             the OPC package
     * @param presentationRef optional weak reference to the owning presentation
     */
    public void initInternal(OpcPackage pkg, Presentation presentationRef) {
        this.pkg = pkg;
        this.presentationRef = presentationRef != null
                ? new WeakReference<>(presentationRef) : null;
        this.corePart = null;
        this.appPart = null;
        this.customPart = null;
    }

    private CorePropertiesPart ensureCore() {
        if (corePart == null) {
            corePart = new CorePropertiesPart(pkg);
        }
        return corePart;
    }

    private AppPropertiesPart ensureApp() {
        if (appPart == null) {
            appPart = new AppPropertiesPart(pkg);
        }
        return appPart;
    }

    private CustomPropertiesPart ensureCustom() {
        if (customPart == null) {
            customPart = new CustomPropertiesPart(pkg);
        }
        return customPart;
    }

    // ---- Core properties ----

    @Override
    public String getTitle() {
        String v = ensureCore().getTitle();
        return v != null ? v : "";
    }

    @Override
    public void setTitle(String value) {
        var core = ensureCore();
        core.setTitle(value);
        core.markDirty();
    }

    @Override
    public String getSubject() {
        String v = ensureCore().getSubject();
        return v != null ? v : "";
    }

    @Override
    public void setSubject(String value) {
        var core = ensureCore();
        core.setSubject(value);
        core.markDirty();
    }

    @Override
    public String getAuthor() {
        String v = ensureCore().getCreator();
        return v != null ? v : "";
    }

    @Override
    public void setAuthor(String value) {
        var core = ensureCore();
        core.setCreator(value);
        core.markDirty();
    }

    @Override
    public String getKeywords() {
        String v = ensureCore().getKeywords();
        return v != null ? v : "";
    }

    @Override
    public void setKeywords(String value) {
        var core = ensureCore();
        core.setKeywords(value);
        core.markDirty();
    }

    @Override
    public String getComments() {
        String v = ensureCore().getDescription();
        return v != null ? v : "";
    }

    @Override
    public void setComments(String value) {
        var core = ensureCore();
        core.setDescription(value);
        core.markDirty();
    }

    @Override
    public String getCategory() {
        String v = ensureCore().getCategory();
        return v != null ? v : "";
    }

    @Override
    public void setCategory(String value) {
        var core = ensureCore();
        core.setCategory(value);
        core.markDirty();
    }

    @Override
    public String getContentStatus() {
        String v = ensureCore().getContentStatus();
        return v != null ? v : "";
    }

    @Override
    public void setContentStatus(String value) {
        var core = ensureCore();
        core.setContentStatus(value);
        core.markDirty();
    }

    @Override
    public String getContentType() {
        String v = ensureCore().getContentType();
        return v != null ? v : "";
    }

    @Override
    public void setContentType(String value) {
        var core = ensureCore();
        core.setContentType(value);
        core.markDirty();
    }

    @Override
    public String getLastSavedBy() {
        String v = ensureCore().getLastModifiedBy();
        return v != null ? v : "";
    }

    @Override
    public void setLastSavedBy(String value) {
        var core = ensureCore();
        core.setLastModifiedBy(value);
        core.markDirty();
    }

    @Override
    public int getRevisionNumber() {
        String rev = ensureCore().getRevision();
        if (rev != null) {
            try {
                return Integer.parseInt(rev.strip());
            } catch (NumberFormatException ignored) {
            }
        }
        return 0;
    }

    @Override
    public void setRevisionNumber(int value) {
        var core = ensureCore();
        core.setRevision(String.valueOf(value));
        core.markDirty();
    }

    @Override
    public LocalDateTime getCreatedTime() {
        OffsetDateTime odt = ensureCore().getCreated();
        return odt != null ? odt.toLocalDateTime() : null;
    }

    @Override
    public void setCreatedTime(LocalDateTime value) {
        var core = ensureCore();
        core.setCreated(value != null ? value.atOffset(ZoneOffset.UTC) : null);
        core.markDirty();
    }

    @Override
    public LocalDateTime getLastSavedTime() {
        OffsetDateTime odt = ensureCore().getModified();
        return odt != null ? odt.toLocalDateTime() : null;
    }

    @Override
    public void setLastSavedTime(LocalDateTime value) {
        var core = ensureCore();
        core.setModified(value != null ? value.atOffset(ZoneOffset.UTC) : null);
        core.markDirty();
    }

    @Override
    public LocalDateTime getLastPrinted() {
        OffsetDateTime odt = ensureCore().getLastPrinted();
        return odt != null ? odt.toLocalDateTime() : null;
    }

    @Override
    public void setLastPrinted(LocalDateTime value) {
        var core = ensureCore();
        core.setLastPrinted(value != null ? value.atOffset(ZoneOffset.UTC) : null);
        core.markDirty();
    }

    // ---- App properties ----

    @Override
    public String getAppVersion() {
        String v = ensureApp().getAppVersion();
        return v != null ? v : "";
    }

    @Override
    public String getNameOfApplication() {
        String v = ensureApp().getApplication();
        return v != null ? v : "";
    }

    @Override
    public void setNameOfApplication(String value) {
        var app = ensureApp();
        app.setApplication(value);
        app.markDirty();
    }

    @Override
    public String getCompany() {
        String v = ensureApp().getCompany();
        return v != null ? v : "";
    }

    @Override
    public void setCompany(String value) {
        var app = ensureApp();
        app.setCompany(value);
        app.markDirty();
    }

    @Override
    public String getManager() {
        String v = ensureApp().getManager();
        return v != null ? v : "";
    }

    @Override
    public void setManager(String value) {
        var app = ensureApp();
        app.setManager(value);
        app.markDirty();
    }

    @Override
    public String getPresentationFormat() {
        String v = ensureApp().getPresentationFormat();
        return v != null ? v : "";
    }

    @Override
    public void setPresentationFormat(String value) {
        var app = ensureApp();
        app.setPresentationFormat(value);
        app.markDirty();
    }

    @Override
    public String getApplicationTemplate() {
        String v = ensureApp().getTemplate();
        return v != null ? v : "";
    }

    @Override
    public void setApplicationTemplate(String value) {
        var app = ensureApp();
        app.setTemplate(value);
        app.markDirty();
    }

    @Override
    public String getHyperlinkBase() {
        String v = ensureApp().getHyperlinkBase();
        return v != null ? v : "";
    }

    @Override
    public void setHyperlinkBase(String value) {
        var app = ensureApp();
        app.setHyperlinkBase(value);
        app.markDirty();
    }

    @Override
    public Duration getTotalEditingTime() {
        Integer minutes = ensureApp().getTotalTime();
        if (minutes != null) {
            return Duration.ofMinutes(minutes);
        }
        return Duration.ZERO;
    }

    @Override
    public void setTotalEditingTime(Duration value) {
        Objects.requireNonNull(value, "value");
        var app = ensureApp();
        app.setTotalTime((int) (value.toSeconds() / 60));
        app.markDirty();
    }

    @Override
    public boolean getSharedDoc() {
        Boolean val = ensureApp().getSharedDoc();
        return val != null ? val : false;
    }

    @Override
    public void setSharedDoc(boolean value) {
        var app = ensureApp();
        app.setSharedDoc(value);
        app.markDirty();
    }

    @Override
    public boolean getScaleCrop() {
        Boolean val = ensureApp().getScaleCrop();
        return val != null ? val : false;
    }

    @Override
    public void setScaleCrop(boolean value) {
        var app = ensureApp();
        app.setScaleCrop(value);
        app.markDirty();
    }

    @Override
    public boolean getLinksUpToDate() {
        Boolean val = ensureApp().getLinksUpToDate();
        return val != null ? val : false;
    }

    @Override
    public void setLinksUpToDate(boolean value) {
        var app = ensureApp();
        app.setLinksUpToDate(value);
        app.markDirty();
    }

    @Override
    public boolean getHyperlinksChanged() {
        Boolean val = ensureApp().getHyperlinksChanged();
        return val != null ? val : false;
    }

    @Override
    public void setHyperlinksChanged(boolean value) {
        var app = ensureApp();
        app.setHyperlinksChanged(value);
        app.markDirty();
    }

    // ---- Read-only statistics ----

    @Override
    public int getSlides() {
        Integer v = ensureApp().getSlides();
        return v != null ? v : 0;
    }

    @Override
    public int getHiddenSlides() {
        Integer v = ensureApp().getHiddenSlides();
        return v != null ? v : 0;
    }

    @Override
    public int getNotes() {
        Integer v = ensureApp().getNotes();
        return v != null ? v : 0;
    }

    @Override
    public int getParagraphs() {
        Integer v = ensureApp().getParagraphs();
        return v != null ? v : 0;
    }

    @Override
    public int getWords() {
        Integer v = ensureApp().getWords();
        return v != null ? v : 0;
    }

    @Override
    public int getMultimediaClips() {
        Integer v = ensureApp().getMmClips();
        return v != null ? v : 0;
    }

    // ---- Heading pairs and titles of parts ----

    @Override
    public List<IHeadingPair> getHeadingPairs() {
        List<IHeadingPair> result = new ArrayList<>();
        for (HeadingPairData data : ensureApp().getHeadingPairs()) {
            HeadingPair hp = new HeadingPair();
            hp.initInternal(data.name(), data.count());
            result.add(hp);
        }
        return result;
    }

    @Override
    public List<String> getTitlesOfParts() {
        return List.copyOf(ensureApp().getTitlesOfParts());
    }

    // ---- Custom properties ----

    @Override
    public int getCountOfCustomProperties() {
        return ensureCustom().getCount();
    }

    @Override
    public void getCustomPropertyValue(String name, List<Object> out) {
        Object val = ensureCustom().getValue(name);
        out.clear();
        if (val != null) {
            out.add(val);
        }
    }

    @Override
    public void setCustomPropertyValue(String name, Object value) {
        Objects.requireNonNull(name, "name");
        Objects.requireNonNull(value, "value");
        ensureCustom().setValue(name, value);
    }

    @Override
    public String getCustomPropertyName(int index) {
        return ensureCustom().getName(index);
    }

    @Override
    public boolean removeCustomProperty(String name) {
        return ensureCustom().remove(name);
    }

    @Override
    public boolean containsCustomProperty(String name) {
        return ensureCustom().contains(name);
    }

    @Override
    public void clearCustomProperties() {
        ensureCustom().clear();
    }

    // ---- Built-in property operations ----

    @Override
    public void clearBuiltInProperties() {
        ensureCore().clear();
        ensureApp().clear();
    }

    // ---- Save ----

    /**
     * Serializes all loaded parts back to the package.
     */
    /**
     * Sets the statistics that describe the document rather than the author.
     *
     * <p>These are recomputed from the presentation before every save; they are
     * not user-settable, and copying them out of a template or a source document
     * is how a file comes to advertise a slide count that is not its own.</p>
     *
     * @param slides     the number of slides
     * @param hidden     the number of hidden slides
     * @param notes      the number of slides that have a notes page
     * @param paragraphs the number of paragraphs across slides and notes
     * @param words      the number of words across slides and notes
     */
    void updateStatistics(int slides, int hidden, int notes, int paragraphs, int words) {
        var app = ensureApp();
        app.setSlides(slides);
        app.setHiddenSlides(hidden);
        app.setNotes(notes);
        app.setParagraphs(paragraphs);
        app.setWords(words);
        app.markDirty();
    }

    public void save() {
        if (corePart != null) corePart.save();
        if (appPart != null) appPart.save();
        if (customPart != null) customPart.save();
    }
}
