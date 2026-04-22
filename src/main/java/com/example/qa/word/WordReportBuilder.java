package com.example.qa.word;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.beans.Introspector;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Modifier;
import java.lang.reflect.Method;
import java.lang.reflect.RecordComponent;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.Consumer;

/**
 * Small facade over Apache POI for creating or extending Word QA evidence reports.
 */
public final class WordReportBuilder implements AutoCloseable {

    private static final String FONT_DEFAULT = "Calibri";
    private static final String FONT_CODE = "Consolas";
    private static final DateTimeFormatter EXECUTION_DATE_FORMAT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    private final XWPFDocument document;
    private final Path outputPath;
    private boolean saved;

    private WordReportBuilder(XWPFDocument document, Path outputPath) {
        this.document = Objects.requireNonNull(document, "document is required");
        this.outputPath = Objects.requireNonNull(outputPath, "outputPath is required");
    }

    public static WordReportBuilder create(Path outputPath) {
        return new WordReportBuilder(new XWPFDocument(), outputPath);
    }

    public static WordReportBuilder open(Path existingPath) {
        Objects.requireNonNull(existingPath, "existingPath is required");

        if (!Files.exists(existingPath)) {
            throw new WordReportException("Word document does not exist: " + existingPath);
        }

        try (var inputStream = Files.newInputStream(existingPath)) {
            return new WordReportBuilder(new XWPFDocument(inputStream), existingPath);
        } catch (IOException exception) {
            throw new WordReportException("Could not open Word document: " + existingPath, exception);
        }
    }

    public static WordReportBuilder openOrCreate(Path outputPath) {
        return Files.exists(outputPath) ? open(outputPath) : create(outputPath);
    }

    public WordReportBuilder withMetadata(String title, String subject, String author) {
        POIXMLProperties.CoreProperties properties = document.getProperties().getCoreProperties();
        properties.setTitle(title);
        properties.setSubjectProperty(subject);
        properties.setCreator(author);
        return this;
    }

    public WordReportBuilder withHeaderText(String text) {
        XWPFHeader header = document.createHeader(HeaderFooterType.DEFAULT);
        XWPFParagraph paragraph = header.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        writeRun(paragraph, text, FONT_DEFAULT, 9, false);
        return this;
    }

    public WordReportBuilder withFooterText(String text) {
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        writeRun(paragraph, text, FONT_DEFAULT, 9, false);
        return this;
    }

    public WordReportBuilder addParagraph(String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfter(160);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_DEFAULT);
        run.setFontSize(11);
        writeMultilineText(run, text);

        return this;
    }

    public WordReportBuilder addHeading(String text, HeadingLevel level) {
        Objects.requireNonNull(level, "level is required");

        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setStyle(level.getWordStyleId());
        paragraph.setSpacingBefore(200);
        paragraph.setSpacingAfter(120);
        applyOutlineLevel(paragraph, level);

        if (level == HeadingLevel.TITLE || level == HeadingLevel.SUBTITLE) {
            paragraph.setAlignment(ParagraphAlignment.CENTER);
        }

        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_DEFAULT);
        run.setBold(true);
        run.setFontSize(level.getFontSize());
        run.setText(nullToEmpty(text));

        if (level == HeadingLevel.SUBTITLE) {
            run.setItalic(true);
        }

        return this;
    }

    public WordReportBuilder addTableOfContents() {
        return addTableOfContents("Tabla de contenido", 1, 3);
    }

    public WordReportBuilder addTableOfContents(String title, int fromLevel, int toLevel) {
        if (fromLevel < 1 || toLevel < fromLevel) {
            throw new WordReportException("Invalid table of contents level range.");
        }

        addHeading(title, HeadingLevel.SECTION);

        XWPFParagraph paragraph = document.createParagraph();
        var field = paragraph.getCTP().addNewFldSimple();
        field.setInstr(String.format(Locale.ROOT, "TOC \\o \"%d-%d\" \\h \\z \\u", fromLevel, toLevel));
        field.addNewR().addNewT().setStringValue("Abra el documento en Word y actualice este campo.");

        document.enforceUpdateFields();
        return this;
    }

    public WordReportBuilder addCollapsibleSection(
            String title,
            HeadingLevel level,
            Consumer<WordReportBuilder> content
    ) {
        Objects.requireNonNull(content, "content is required");

        addHeading(title, level);
        content.accept(this);
        return this;
    }

    public WordReportBuilder addCodeBlock(String code) {
        XWPFTable table = document.createTable(1, 1);
        table.setWidth("100%");
        table.setCellMargins(120, 120, 120, 120);

        XWPFTableCell cell = table.getRow(0).getCell(0);
        cell.setColor("F3F4F6");
        removeDefaultParagraph(cell);

        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setSpacingBefore(60);
        paragraph.setSpacingAfter(60);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_CODE);
        run.setFontSize(9);
        run.setColor("111827");
        writeMultilineText(run, code);

        return this;
    }

    public WordReportBuilder addBulletList(List<String> items) {
        if (items == null || items.isEmpty()) {
            return this;
        }

        for (String item : items) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setIndentationLeft(400);
            writeRun(paragraph, "- " + nullToEmpty(item), FONT_DEFAULT, 11, false);
        }

        return this;
    }

    public WordReportBuilder addNumberedList(List<String> items) {
        if (items == null || items.isEmpty()) {
            return this;
        }

        int index = 1;
        for (String item : items) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setIndentationLeft(400);
            writeRun(paragraph, index++ + ". " + nullToEmpty(item), FONT_DEFAULT, 11, false);
        }

        return this;
    }

    public WordReportBuilder addTableFromObject(Object object) {
        if (object == null) {
            return addParagraph("No hay datos disponibles.");
        }

        return addTableFromMap(extractValues(object));
    }

    public WordReportBuilder addTableFromObjects(List<?> objects) {
        if (objects == null || objects.isEmpty()) {
            return addParagraph("No hay filas disponibles.");
        }

        List<Map<String, Object>> rows = objects.stream()
                .filter(Objects::nonNull)
                .map(this::extractValues)
                .toList();

        if (rows.isEmpty()) {
            return addParagraph("No hay filas disponibles.");
        }

        Set<String> headers = new LinkedHashSet<>();
        rows.forEach(row -> headers.addAll(row.keySet()));

        return addTableFromRows(headers, rows);
    }

    public WordReportBuilder addTableFromMap(Map<String, ?> map) {
        if (map == null || map.isEmpty()) {
            return addParagraph("No hay datos disponibles.");
        }

        XWPFTable table = document.createTable(map.size() + 1, 2);
        styleTable(table);
        writeTableRow(table.getRow(0), List.of("Campo", "Valor"), true);

        int rowIndex = 1;
        for (Map.Entry<String, ?> entry : map.entrySet()) {
            writeTableRow(
                    table.getRow(rowIndex++),
                    List.of(entry.getKey(), stringify(entry.getValue())),
                    false
            );
        }

        return this;
    }

    public WordReportBuilder addTableFromMaps(List<Map<String, ?>> rows) {
        if (rows == null || rows.isEmpty()) {
            return addParagraph("No hay filas disponibles.");
        }

        Set<String> headers = new LinkedHashSet<>();
        rows.forEach(row -> headers.addAll(row.keySet()));

        List<Map<String, Object>> normalizedRows = rows.stream()
                .map(row -> (Map<String, Object>) new LinkedHashMap<String, Object>(row))
                .toList();

        return addTableFromRows(headers, normalizedRows);
    }

    public WordReportBuilder addImage(Path imagePath, int widthPx, int heightPx) {
        Objects.requireNonNull(imagePath, "imagePath is required");

        if (!Files.exists(imagePath)) {
            throw new WordReportException("Image file does not exist: " + imagePath);
        }

        try (var inputStream = Files.newInputStream(imagePath)) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun run = paragraph.createRun();
            run.addPicture(
                    inputStream,
                    detectPictureType(imagePath),
                    imagePath.getFileName().toString(),
                    Units.toEMU(widthPx),
                    Units.toEMU(heightPx)
            );

            return this;
        } catch (Exception exception) {
            throw new WordReportException("Could not insert image: " + imagePath, exception);
        }
    }

    public WordReportBuilder addPageBreak() {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.createRun().addBreak(BreakType.PAGE);
        return this;
    }

    public WordReportBuilder addSeparator() {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setBorderBottom(Borders.SINGLE);
        paragraph.setSpacingAfter(180);
        return this;
    }

    public WordReportBuilder addExecutionDate() {
        return addParagraph("Fecha de ejecucion: " + EXECUTION_DATE_FORMAT.format(LocalDateTime.now()));
    }

    public WordReportBuilder addTestEvidence(String title, String description, Path screenshot) {
        addHeading(title, HeadingLevel.SECTION);
        addParagraph(description);

        if (screenshot != null && Files.exists(screenshot)) {
            addImage(screenshot, 620, 350);
        }

        return this;
    }

    public WordReportBuilder addRequestResponseEvidence(
            String title,
            String request,
            String response,
            int statusCode,
            long durationMillis
    ) {
        return addCollapsibleSection(title, HeadingLevel.SECTION, section -> section
                .addTableFromMap(Map.of(
                        "HTTP Status", statusCode,
                        "Duracion ms", durationMillis
                ))
                .addHeading("Request", HeadingLevel.SUBSECTION)
                .addCodeBlock(request)
                .addHeading("Response", HeadingLevel.SUBSECTION)
                .addCodeBlock(response));
    }

    public WordReportBuilder addLogEvidence(
            String label,
            String content
    ) {
        return addEvidenceSection(label, content);
    }

    private WordReportBuilder addEvidenceSection(String label, String content) {
        addParagraph(label);
        addCodeBlock(content);
        return this;
    }

    public void save() {
        saveAs(outputPath);
    }

    public void saveAs(Path targetPath) {
        Objects.requireNonNull(targetPath, "targetPath is required");

        try {
            Path parent = targetPath.getParent();
            if (parent != null) {
                Files.createDirectories(parent);
            }

            try (OutputStream outputStream = Files.newOutputStream(targetPath)) {
                document.write(outputStream);
            }

            saved = true;
        } catch (IOException exception) {
            throw new WordReportException("Could not save Word document: " + targetPath, exception);
        }
    }

    @Override
    public void close() {
        try {
            document.close();
        } catch (IOException exception) {
            throw new WordReportException("Could not close Word document.", exception);
        }
    }

    public boolean isSaved() {
        return saved;
    }

    private WordReportBuilder addTableFromRows(Set<String> headers, List<Map<String, Object>> rows) {
        if (headers.isEmpty()) {
            return addParagraph("No hay columnas disponibles.");
        }

        List<String> headerList = new ArrayList<>(headers);
        XWPFTable table = document.createTable(rows.size() + 1, headerList.size());
        styleTable(table);
        writeTableRow(table.getRow(0), headerList, true);

        int rowIndex = 1;
        for (Map<String, Object> rowData : rows) {
            List<String> values = headerList.stream()
                    .map(header -> stringify(rowData.get(header)))
                    .toList();

            writeTableRow(table.getRow(rowIndex++), values, false);
        }

        return this;
    }

    private void styleTable(XWPFTable table) {
        table.setWidth("100%");
        table.setCellMargins(80, 80, 80, 80);
    }

    private void writeTableRow(XWPFTableRow row, List<String> values, boolean header) {
        for (int i = 0; i < values.size(); i++) {
            XWPFTableCell cell = row.getCell(i);
            if (cell == null) {
                cell = row.addNewTableCell();
            }

            removeDefaultParagraph(cell);

            XWPFParagraph paragraph = cell.addParagraph();
            XWPFRun run = paragraph.createRun();
            run.setFontFamily(FONT_DEFAULT);
            run.setFontSize(10);
            run.setBold(header);
            run.setText(values.get(i));

            if (header) {
                cell.setColor("D9EAF7");
            }
        }
    }

    private void writeRun(XWPFParagraph paragraph, String text, String fontFamily, int fontSize, boolean bold) {
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        run.setBold(bold);
        writeMultilineText(run, text);
    }

    private void writeMultilineText(XWPFRun run, String text) {
        String safeText = nullToEmpty(text);
        String[] lines = safeText.split("\\R", -1);

        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {
                run.addBreak();
            }
            run.setText(lines[i]);
        }
    }

    private void applyOutlineLevel(XWPFParagraph paragraph, HeadingLevel level) {
        var paragraphProperties = paragraph.getCTP().isSetPPr()
                ? paragraph.getCTP().getPPr()
                : paragraph.getCTP().addNewPPr();

        var outlineLevel = paragraphProperties.isSetOutlineLvl()
                ? paragraphProperties.getOutlineLvl()
                : paragraphProperties.addNewOutlineLvl();

        outlineLevel.setVal(BigInteger.valueOf(level.getOutlineLevel()));
    }

    private void removeDefaultParagraph(XWPFTableCell cell) {
        while (!cell.getParagraphs().isEmpty()) {
            cell.removeParagraph(0);
        }
    }

    private Map<String, Object> extractValues(Object object) {
        if (object instanceof Map<?, ?> map) {
            Map<String, Object> result = new LinkedHashMap<>();
            map.forEach((key, value) -> result.put(String.valueOf(key), value));
            return result;
        }

        if (isSimpleValue(object)) {
            return new LinkedHashMap<>(Map.of("value", object));
        }

        if (object.getClass().isRecord()) {
            return extractRecordValues(object);
        }

        Map<String, Object> beanValues = extractBeanValues(object);
        return beanValues.isEmpty() ? extractFieldValues(object) : beanValues;
    }

    private Map<String, Object> extractRecordValues(Object record) {
        Map<String, Object> values = new LinkedHashMap<>();

        try {
            for (RecordComponent component : record.getClass().getRecordComponents()) {
                Method accessor = component.getAccessor();
                accessor.setAccessible(true);
                Object value = accessor.invoke(record);
                values.put(component.getName(), value);
            }
            return values;
        } catch (ReflectiveOperationException exception) {
            throw new WordReportException("Could not extract record values.", exception);
        }
    }

    private Map<String, Object> extractBeanValues(Object bean) {
        Map<String, Object> values = new LinkedHashMap<>();

        try {
            var beanInfo = Introspector.getBeanInfo(bean.getClass(), Object.class);

            for (var descriptor : beanInfo.getPropertyDescriptors()) {
                if (descriptor.getReadMethod() != null) {
                    Object value = descriptor.getReadMethod().invoke(bean);
                    values.put(descriptor.getName(), value);
                }
            }

            return values;
        } catch (ReflectiveOperationException | java.beans.IntrospectionException exception) {
            throw new WordReportException("Could not extract bean values.", exception);
        }
    }

    private Map<String, Object> extractFieldValues(Object object) {
        Map<String, Object> values = new LinkedHashMap<>();

        try {
            Class<?> currentClass = object.getClass();
            while (currentClass != null && currentClass != Object.class) {
                for (var field : currentClass.getDeclaredFields()) {
                    if (Modifier.isStatic(field.getModifiers())) {
                        continue;
                    }

                    field.setAccessible(true);
                    values.put(field.getName(), field.get(object));
                }
                currentClass = currentClass.getSuperclass();
            }

            return values;
        } catch (IllegalAccessException exception) {
            throw new WordReportException("Could not extract field values.", exception);
        }
    }

    private boolean isSimpleValue(Object value) {
        return value instanceof CharSequence
                || value instanceof Number
                || value instanceof Boolean
                || value instanceof Enum<?>
                || value instanceof TemporalAccessor;
    }

    private String stringify(Object value) {
        if (value == null) {
            return "";
        }

        if (value instanceof Collection<?> collection) {
            return String.join(", ", collection.stream().map(this::stringify).toList());
        }

        return String.valueOf(value);
    }

    private int detectPictureType(Path imagePath) {
        String filename = imagePath.getFileName().toString().toLowerCase(Locale.ROOT);

        if (filename.endsWith(".png")) {
            return Document.PICTURE_TYPE_PNG;
        }

        if (filename.endsWith(".jpg") || filename.endsWith(".jpeg")) {
            return Document.PICTURE_TYPE_JPEG;
        }

        if (filename.endsWith(".gif")) {
            return Document.PICTURE_TYPE_GIF;
        }

        if (filename.endsWith(".bmp")) {
            return Document.PICTURE_TYPE_BMP;
        }

        throw new WordReportException("Unsupported image format: " + filename);
    }

    private String nullToEmpty(String text) {
        return text == null ? "" : text;
    }
}
