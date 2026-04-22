package com.example.qa.word;

public enum HeadingLevel {
    TITLE("Heading1", 0, 20),
    SUBTITLE("Heading2", 1, 15),
    SECTION("Heading2", 1, 14),
    SUBSECTION("Heading3", 2, 12);

    private final String wordStyleId;
    private final int outlineLevel;
    private final int fontSize;

    HeadingLevel(String wordStyleId, int outlineLevel, int fontSize) {
        this.wordStyleId = wordStyleId;
        this.outlineLevel = outlineLevel;
        this.fontSize = fontSize;
    }

    public String getWordStyleId() {
        return wordStyleId;
    }

    public int getOutlineLevel() {
        return outlineLevel;
    }

    public int getFontSize() {
        return fontSize;
    }
}
