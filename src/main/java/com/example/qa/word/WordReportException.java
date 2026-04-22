package com.example.qa.word;

/**
 * Excepcion de dominio para errores ocurridos durante la generacion o escritura
 * de reportes Word.
 */
public class WordReportException extends RuntimeException {

    /**
     * Crea una excepcion con un mensaje descriptivo.
     *
     * @param message descripcion del error
     */
    public WordReportException(String message) {
        super(message);
    }

    /**
     * Crea una excepcion con un mensaje descriptivo y la causa original.
     *
     * @param message descripcion del error
     * @param cause excepcion original que produjo el fallo
     */
    public WordReportException(String message, Throwable cause) {
        super(message, cause);
    }
}
