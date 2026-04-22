package com.example.qa.demo;

import com.example.qa.word.HeadingLevel;
import com.example.qa.word.WordReportBuilder;

import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

public class WordReportExample {

    public static void main(String[] args) {
        Path reportPath = Path.of("target", "reports", "qa-execution-report.docx");

        try (WordReportBuilder report = WordReportBuilder.openOrCreate(reportPath)) {
            report.withMetadata(
                            "Reporte de evidencias QA",
                            "Evidencias de ejecucion automatizada",
                            "QA Automation Team"
                    )
                    .withHeaderText("Evidencias de test suite") // Ingresar el test suite
                    .withFooterText("Team")
                    .addHeading("Reporte de ejecucion QA", HeadingLevel.TITLE) // Titulo del reporte
                    .addExecutionDate()
                    .addTableOfContents()
                    .addPageBreak()
                    .addCollapsibleSection("Resumen de la ejecucion", HeadingLevel.SECTION, section -> section
                            .addTableFromObject(new ExecutionSummary(
                                    "Regression Suite",
                                    "QA",
                                    "Chrome",
                                    8,
                                    7,
                                    1,
                                    LocalDateTime.now()
                            ))
                            .addBulletList(List.of(
                                    "Se ejecutaron validaciones funcionales principales.",
                                    "Se adjuntan datos tecnicos de request y response.",
                                    "Las secciones se pueden contraer desde Microsoft Word usando los encabezados."
                            )))
                    .addCollapsibleSection("Caso: login exitoso", HeadingLevel.SECTION, section -> section
                            .addTableFromObject(new TestCaseResult(
                                    "TC-LOGIN-001",
                                    "Login exitoso",
                                    "PASSED",
                                    1240,
                                    "El usuario ingreso al dashboard correctamente."
                            ))
                            .addHeading("Datos de entrada", HeadingLevel.SUBSECTION)
                            .addTableFromMap(Map.of(
                                    "username", "qa.user@example.com",
                                    "environment", "QA",
                                    "browser", "Chrome"
                            )))
                    .addRequestResponseEvidence(
                            "Evidencia API: autenticacion",
                            """
                                    POST /api/login HTTP/1.1
                                    Content-Type: application/json
                                    
                                    {
                                      "username": "qa.user@example.com",
                                      "password": "***"
                                    }
                                    """,
                            """
                                    HTTP/1.1 200 OK
                                    Content-Type: application/json
                                    
                                    {
                                      "status": "OK",
                                      "tokenType": "Bearer",
                                      "expiresIn": 3600
                                    }
                                    """,
                            200,
                            318
                    )
                    .addCollapsibleSection("Resultados por caso", HeadingLevel.SECTION, section -> section
                            .addTableFromObjects(List.of(
                                    new TestCaseResult("TC-LOGIN-001", "Login exitoso", "PASSED", 1240, "OK"),
                                    new TestCaseResult("TC-LOGIN-002", "Password invalido", "PASSED", 890, "Mensaje esperado"),
                                    new TestCaseResult("TC-LOGIN-003", "Usuario bloqueado", "FAILED", 1020, "No aparecio alerta esperada"),
                                    new TestCaseInformation("Orchestrator", "QA", "PASSED") // Toma de encabezado el record
                            )))
                    .addCollapsibleSection("Example codeInfo", HeadingLevel.SECTION, section -> section
                            .addHeading("Log", HeadingLevel.SUBSECTION)
                            .addLogEvidence( "label", """
                                    Dec 15, 2012 1:42:43 AM com.journaldev.log.LoggingExample main
                                    INFO: Msg997
                                    Dec 15, 2012 1:42:43 AM com.journaldev.log.LoggingExample main
                                    INFO: Msg998
                                    Dec 15, 2012 1:42:43 AM com.journaldev.log.LoggingExample main
                                    INFO: Msg998
                                    """)
                    )
                    .addPageBreak()
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
