package com.example.qa.demo;

import com.example.qa.word.HeadingLevel;
import com.example.qa.word.WordReportBuilder;

import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

public class WordReUseExample {

    public static void main(String[] args) {
        Path reportPath = Path.of("target", "reports", "qa-execution-report.docx");

        try (WordReportBuilder report = WordReportBuilder.openOrCreate(reportPath)) {
            report.addCollapsibleSection("Transaction Type (Scenario Name)", HeadingLevel.SECTION, section -> section
                            .addHeading("Test Case Information", HeadingLevel.SUBSECTION)
                            .addLogEvidence("request", """
                                    Dec 15, 2012 1:42:43 AM com.journaldev.log.LoggingExample main
                                    INFO: Msg997
                                    Dec 15, 2012 1:42:43 AM com.journaldev.log.LoggingExample main
                                    INFO: Msg998
                                    Dec 15, 2012 1:42:43 AM com.journaldev.log.LoggingExample main
                                    INFO: Msg998
                                    """)
                            .addSeparator()
                            .addLogEvidence("response", """
                                    Dec 15, 2012 1:42:43 AM com.journaldev.log.LoggingExample main
                                    INFO: Msg997
                                    Dec 15, 2012 1:42:43 AM com.journaldev.log.LoggingExample main
                                    INFO: Msg998
                                    Dec 15, 2012 1:42:43 AM com.journaldev.log.LoggingExample main
                                    INFO: Msg998
                                    """)
                    )
                    .save();
        }

        System.out.println("Reporte generado en: " + reportPath.toAbsolutePath());
    }

    record ExecutionSummary(
            String suite,
            String environment,
            String browser,
            int totalTests,
            int passed,
            int failed,
            LocalDateTime executionDate
    ) {
    }

    record TestCaseResult(
            String id,
            String name,
            String status,
            long durationMillis,
            String observation
    ) {
    }

    record TestCaseInformation(
            String campoOrchestrator,
            String campoEntrada,
            String resultadoValidacion
    ) {
    }
}
