package com.example.qa.word;

/**
 * Define los niveles de encabezado soportados por el generador de reportes Word.
 *
 * <p>Cada nivel encapsula el estilo nativo de Word, el nivel de esquema usado
 * por la tabla de contenido y el tamano de fuente sugerido para el encabezado.</p>
 */
public enum HeadingLevel {
    /**
     * Encabezado principal del documento. Se usa como nivel 1 en la tabla de contenido.
     */
    TITLE("Heading1", 0, 20),

    /**
     * Subtitulo general del documento. Se usa como nivel 2 en la tabla de contenido.
     */
    SUBTITLE("Heading2", 1, 15),

    /**
     * Encabezado para secciones principales del reporte.
     */
    SECTION("Heading2", 1, 14),

    /**
     * Encabezado para subsecciones dentro de una seccion.
     */
    SUBSECTION("Heading3", 2, 12);

    private final String wordStyleId;
    private final int outlineLevel;
    private final int fontSize;

    /**
     * Crea una definicion de nivel de encabezado.
     *
     * @param wordStyleId identificador del estilo nativo de Word, por ejemplo {@code Heading1}
     * @param outlineLevel nivel de esquema usado por Word para indice y colapso de secciones
     * @param fontSize tamano de fuente sugerido para el encabezado
     */
    HeadingLevel(String wordStyleId, int outlineLevel, int fontSize) {
        this.wordStyleId = wordStyleId;
        this.outlineLevel = outlineLevel;
        this.fontSize = fontSize;
    }

    /**
     * Obtiene el identificador del estilo nativo de Word asociado al nivel.
     *
     * @return identificador del estilo de Word
     */
    public String getWordStyleId() {
        return wordStyleId;
    }

    /**
     * Obtiene el nivel de esquema usado por Word para tabla de contenido y navegacion.
     *
     * @return nivel de esquema, iniciando en cero para el nivel mas alto
     */
    public int getOutlineLevel() {
        return outlineLevel;
    }

    /**
     * Obtiene el tamano de fuente sugerido para renderizar el encabezado.
     *
     * @return tamano de fuente en puntos
     */
    public int getFontSize() {
        return fontSize;
    }
}
