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
 * Fachada fluida sobre Apache POI para crear, abrir y extender documentos Word
 * usados como reportes o evidencias de automatizacion QA.
 *
 * <p>La clase centraliza operaciones comunes como titulos, parrafos, tablas,
 * bloques de codigo, evidencias tecnicas, tabla de contenido y guardado del
 * archivo, evitando que las pruebas dependan directamente de Apache POI.</p>
 */
public final class WordReportBuilder implements AutoCloseable {

    private static final String FONT_DEFAULT = "Calibri";
    private static final String FONT_CODE = "Consolas";
    private static final DateTimeFormatter EXECUTION_DATE_FORMAT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    private final XWPFDocument document;
    private final Path outputPath;
    private boolean saved;

    /**
     * Inicializa el builder con un documento Apache POI y una ruta de salida.
     *
     * @param document documento Word en memoria
     * @param outputPath ruta usada por {@link #save()}
     */
    private WordReportBuilder(XWPFDocument document, Path outputPath) {
        this.document = Objects.requireNonNull(document, "document is required");
        this.outputPath = Objects.requireNonNull(outputPath, "outputPath is required");
    }

    /**
     * Crea un documento Word nuevo asociado a la ruta indicada.
     *
     * @param outputPath ruta donde se guardara el documento al ejecutar {@link #save()}
     * @return builder listo para agregar contenido
     */
    public static WordReportBuilder create(Path outputPath) {
        return new WordReportBuilder(new XWPFDocument(), outputPath);
    }

    /**
     * Abre un archivo Word existente para agregar o modificar contenido.
     *
     * @param existingPath ruta del archivo {@code .docx} existente
     * @return builder construido sobre el documento existente
     * @throws WordReportException si el archivo no existe o no puede abrirse
     */
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

    /**
     * Abre un documento existente si ya existe o crea uno nuevo si la ruta aun no existe.
     *
     * <p>Este metodo es util para reportes acumulativos de ejecuciones QA.</p>
     *
     * @param outputPath ruta del documento a abrir o crear
     * @return builder listo para agregar contenido
     */
    public static WordReportBuilder openOrCreate(Path outputPath) {
        return Files.exists(outputPath) ? open(outputPath) : create(outputPath);
    }

    /**
     * Agrega metadatos basicos al documento Word.
     *
     * @param title titulo del documento
     * @param subject asunto o descripcion corta
     * @param author autor o equipo responsable del reporte
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder withMetadata(String title, String subject, String author) {
        POIXMLProperties.CoreProperties properties = document.getProperties().getCoreProperties();
        properties.setTitle(title);
        properties.setSubjectProperty(subject);
        properties.setCreator(author);
        return this;
    }

    /**
     * Configura un encabezado de pagina con texto alineado a la derecha.
     *
     * @param text texto que aparecera en el encabezado del documento
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder withHeaderText(String text) {
        XWPFHeader header = document.createHeader(HeaderFooterType.DEFAULT);
        XWPFParagraph paragraph = header.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        writeRun(paragraph, text, FONT_DEFAULT, 9, false);
        return this;
    }

    /**
     * Configura un pie de pagina con texto centrado.
     *
     * @param text texto que aparecera en el pie de pagina del documento
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder withFooterText(String text) {
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        writeRun(paragraph, text, FONT_DEFAULT, 9, false);
        return this;
    }

    /**
     * Agrega un parrafo de texto normal al documento.
     *
     * <p>Si el texto contiene saltos de linea, se respetan dentro del mismo parrafo.</p>
     *
     * @param text contenido textual a insertar
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder addParagraph(String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfter(160);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_DEFAULT);
        run.setFontSize(11);
        writeMultilineText(run, text);

        return this;
    }

    /**
     * Agrega un encabezado usando estilos nativos de Word y nivel de esquema.
     *
     * <p>Estos encabezados alimentan la tabla de contenido y permiten que Word
     * simule secciones colapsables desde el panel de navegacion o el propio titulo.</p>
     *
     * @param text texto del encabezado
     * @param level nivel semantico y visual del encabezado
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Agrega una tabla de contenido con niveles 1 a 3.
     *
     * <p>La tabla se inserta como campo de Word. Microsoft Word puede solicitar
     * actualizar el campo al abrir el documento para calcular titulos y paginas.</p>
     *
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder addTableOfContents() {
        return addTableOfContents("Tabla de contenido", 1, 3);
    }

    /**
     * Agrega una tabla de contenido parametrizada.
     *
     * @param title titulo visible antes de la tabla de contenido
     * @param fromLevel primer nivel de encabezado incluido, iniciando en 1
     * @param toLevel ultimo nivel de encabezado incluido
     * @return la misma instancia para encadenar llamadas
     * @throws WordReportException si el rango de niveles es invalido
     */
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

    /**
     * Agrega una seccion con encabezado y contenido construido mediante una funcion.
     *
     * <p>La seccion queda preparada para comportarse como colapsable en Word porque
     * el titulo usa estilos de encabezado y nivel de esquema.</p>
     *
     * @param title titulo de la seccion
     * @param level nivel de encabezado de la seccion
     * @param content bloque que agrega el contenido interno de la seccion
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Agrega un bloque de texto tecnico con apariencia de consola o codigo.
     *
     * <p>Esta pensado para logs, JSON, XML, requests, responses, SQL,
     * stack traces o cualquier evidencia tecnica monoespaciada.</p>
     *
     * @param code contenido tecnico a insertar
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Agrega una lista con vineta simple.
     *
     * @param items elementos a renderizar; si es {@code null} o vacia no agrega contenido
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Agrega una lista numerada simple.
     *
     * @param items elementos a renderizar; si es {@code null} o vacia no agrega contenido
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Genera una tabla clave/valor a partir de un objeto.
     *
     * <p>Soporta {@link Map}, records, JavaBeans y, como ultima opcion,
     * campos de instancia mediante reflexion.</p>
     *
     * @param object objeto fuente de datos
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder addTableFromObject(Object object) {
        if (object == null) {
            return addParagraph("No hay datos disponibles.");
        }

        return addTableFromMap(extractValues(object));
    }

    /**
     * Genera una tabla dinamica a partir de una lista de objetos.
     *
     * <p>Las columnas se calculan uniendo los nombres de propiedades/campos
     * encontrados en todos los elementos de la lista.</p>
     *
     * @param objects lista de objetos fuente
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Genera una tabla clave/valor a partir de un mapa.
     *
     * @param map mapa con nombres de campo y valores
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Genera una tabla dinamica a partir de una lista de mapas.
     *
     * <p>Las columnas se construyen con la union ordenada de las llaves
     * presentes en los mapas.</p>
     *
     * @param rows filas representadas como mapas de columna/valor
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Inserta una imagen centrada en el documento.
     *
     * @param imagePath ruta de la imagen a insertar
     * @param widthPx ancho deseado en pixeles
     * @param heightPx alto deseado en pixeles
     * @return la misma instancia para encadenar llamadas
     * @throws WordReportException si la imagen no existe, tiene formato no soportado o no puede insertarse
     */
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

    /**
     * Inserta un salto de pagina.
     *
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder addPageBreak() {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.createRun().addBreak(BreakType.PAGE);
        return this;
    }

    /**
     * Inserta un separador horizontal sencillo.
     *
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder addSeparator() {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setBorderBottom(Borders.SINGLE);
        paragraph.setSpacingAfter(180);
        return this;
    }

    /**
     * Inserta la fecha y hora local de ejecucion con formato {@code yyyy-MM-dd HH:mm:ss}.
     *
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder addExecutionDate() {
        return addParagraph("Fecha de ejecucion: " + EXECUTION_DATE_FORMAT.format(LocalDateTime.now()));
    }

    /**
     * Agrega una evidencia de prueba con titulo, descripcion y screenshot opcional.
     *
     * @param title titulo de la evidencia o caso de prueba
     * @param description descripcion, resultado u observacion asociada
     * @param screenshot ruta opcional de la captura; si es {@code null} o no existe se omite
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder addTestEvidence(String title, String description, Path screenshot) {
        addHeading(title, HeadingLevel.SECTION);
        addParagraph(description);

        if (screenshot != null && Files.exists(screenshot)) {
            addImage(screenshot, 620, 350);
        }

        return this;
    }

    /**
     * Agrega una evidencia tecnica de API con resumen, request y response.
     *
     * @param title titulo de la evidencia API
     * @param request request serializado o log de entrada
     * @param response response serializado o log de salida
     * @param statusCode codigo HTTP obtenido
     * @param durationMillis duracion de la operacion en milisegundos
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Agrega una evidencia de log con etiqueta y contenido monoespaciado.
     *
     * @param label etiqueta descriptiva del log
     * @param content contenido completo del log
     * @return la misma instancia para encadenar llamadas
     */
    public WordReportBuilder addLogEvidence(
            String label,
            String content
    ) {
        return addEvidenceSection(label, content);
    }

    /**
     * Agrega una seccion generica de evidencia compuesta por etiqueta y bloque tecnico.
     *
     * @param label etiqueta visible antes del bloque
     * @param content contenido tecnico de la evidencia
     * @return la misma instancia para encadenar llamadas
     */
    private WordReportBuilder addEvidenceSection(String label, String content) {
        addParagraph(label);
        addCodeBlock(content);
        return this;
    }

    /**
     * Guarda el documento en la ruta asociada al builder.
     *
     * @throws WordReportException si no se puede crear el directorio o escribir el archivo
     */
    public void save() {
        saveAs(outputPath);
    }

    /**
     * Guarda el documento en una ruta especifica.
     *
     * <p>Permite abrir una plantilla o documento existente y exportar el resultado
     * con otro nombre sin sobrescribir la fuente original.</p>
     *
     * @param targetPath ruta final donde se escribira el documento
     * @throws WordReportException si no se puede crear el directorio o escribir el archivo
     */
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

    /**
     * Cierra los recursos del documento en memoria.
     *
     * @throws WordReportException si Apache POI no puede cerrar el documento
     */
    @Override
    public void close() {
        try {
            document.close();
        } catch (IOException exception) {
            throw new WordReportException("Could not close Word document.", exception);
        }
    }

    /**
     * Indica si el documento fue guardado correctamente durante el ciclo de vida del builder.
     *
     * @return {@code true} si se ejecuto {@link #save()} o {@link #saveAs(Path)} con exito
     */
    public boolean isSaved() {
        return saved;
    }

    /**
     * Construye una tabla con encabezados dinamicos y filas normalizadas.
     *
     * @param headers columnas que tendra la tabla
     * @param rows datos normalizados por nombre de columna
     * @return la misma instancia para encadenar llamadas
     */
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

    /**
     * Aplica formato base reutilizable a una tabla.
     *
     * @param table tabla Apache POI a formatear
     */
    private void styleTable(XWPFTable table) {
        table.setWidth("100%");
        table.setCellMargins(80, 80, 80, 80);
    }

    /**
     * Escribe una fila de tabla, aplicando estilo especial si corresponde al encabezado.
     *
     * @param row fila de destino
     * @param values valores de cada celda
     * @param header indica si la fila es encabezado
     */
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

    /**
     * Crea un run de texto con formato basico dentro de un parrafo.
     *
     * @param paragraph parrafo de destino
     * @param text texto a escribir
     * @param fontFamily familia tipografica
     * @param fontSize tamano de fuente en puntos
     * @param bold indica si el texto debe ir en negrita
     */
    private void writeRun(XWPFParagraph paragraph, String text, String fontFamily, int fontSize, boolean bold) {
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        run.setBold(bold);
        writeMultilineText(run, text);
    }

    /**
     * Escribe texto conservando saltos de linea dentro del run de Word.
     *
     * @param run run de destino
     * @param text texto posiblemente multilinea
     */
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

    /**
     * Configura el nivel de esquema del parrafo para navegacion, indice y colapso en Word.
     *
     * @param paragraph parrafo de encabezado
     * @param level nivel semantico que define el outline
     */
    private void applyOutlineLevel(XWPFParagraph paragraph, HeadingLevel level) {
        var paragraphProperties = paragraph.getCTP().isSetPPr()
                ? paragraph.getCTP().getPPr()
                : paragraph.getCTP().addNewPPr();

        var outlineLevel = paragraphProperties.isSetOutlineLvl()
                ? paragraphProperties.getOutlineLvl()
                : paragraphProperties.addNewOutlineLvl();

        outlineLevel.setVal(BigInteger.valueOf(level.getOutlineLevel()));
    }

    /**
     * Elimina los parrafos creados por defecto dentro de una celda de tabla.
     *
     * @param cell celda a limpiar antes de escribir contenido
     */
    private void removeDefaultParagraph(XWPFTableCell cell) {
        while (!cell.getParagraphs().isEmpty()) {
            cell.removeParagraph(0);
        }
    }

    /**
     * Extrae valores de un objeto en forma de mapa ordenado.
     *
     * @param object objeto fuente; puede ser mapa, valor simple, record, bean o clase con campos
     * @return mapa con nombres de propiedades/campos y sus valores
     */
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

    /**
     * Extrae componentes de un record Java usando sus accessors.
     *
     * @param record instancia de record
     * @return mapa ordenado con nombres de componentes y valores
     */
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

    /**
     * Extrae propiedades legibles de un JavaBean mediante introspeccion.
     *
     * @param bean instancia tipo JavaBean
     * @return mapa ordenado con propiedades y valores
     */
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

    /**
     * Extrae campos de instancia cuando el objeto no expone propiedades JavaBean.
     *
     * @param object objeto fuente
     * @return mapa ordenado con nombres de campo y valores
     */
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

    /**
     * Determina si un valor puede representarse directamente en una celda sin introspeccion.
     *
     * @param value valor a evaluar
     * @return {@code true} si es texto, numero, booleano, enum o tipo temporal
     */
    private boolean isSimpleValue(Object value) {
        return value instanceof CharSequence
                || value instanceof Number
                || value instanceof Boolean
                || value instanceof Enum<?>
                || value instanceof TemporalAccessor;
    }

    /**
     * Convierte un valor a texto seguro para insertar en Word.
     *
     * @param value valor original
     * @return representacion textual; cadena vacia si el valor es {@code null}
     */
    private String stringify(Object value) {
        if (value == null) {
            return "";
        }

        if (value instanceof Collection<?> collection) {
            return String.join(", ", collection.stream().map(this::stringify).toList());
        }

        return String.valueOf(value);
    }

    /**
     * Detecta el tipo de imagen soportado por Apache POI segun la extension del archivo.
     *
     * @param imagePath ruta de la imagen
     * @return constante de tipo de imagen esperada por Apache POI
     * @throws WordReportException si la extension no esta soportada
     */
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

    /**
     * Normaliza texto nulo a cadena vacia.
     *
     * @param text texto original
     * @return texto original o cadena vacia si era {@code null}
     */
    private String nullToEmpty(String text) {
        return text == null ? "" : text;
    }
}
