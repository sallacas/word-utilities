package com.example.qa.word;

public class WordReportException extends RuntimeException {

    public WordReportException(String message) {
        super(message);
    }

    public WordReportException(String message, Throwable cause) {
        super(message, cause);
    }
}
