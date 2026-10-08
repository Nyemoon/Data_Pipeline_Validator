"""
Geração de Relatórios de Auditoria em PDF via ReportLab
======================================================
Módulo utilitário puro para geração de relatórios de validação em PDF.
Desacoplado da camada de interface gráfica (Streamlit).
"""

from __future__ import annotations

import io
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from validador_planilhas import (
    SEVERITY_CRITICAL,
    SEVERITY_INFO,
    SEVERITY_WARNING,
    ValidationReport,
)


def gerar_pdf_relatorio(reports: list[ValidationReport]) -> bytes:
    """Gera um documento PDF estilizado contendo a auditoria das partições/abas."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40,
    )

    styles = getSampleStyleSheet()

    # Paleta de Cores Alinhada ao Tema
    primary_color = colors.HexColor("#2A1B3D")
    accent_color = colors.HexColor("#FF9E64")
    critical_color = colors.HexColor("#FF7B72")
    warning_color = colors.HexColor("#E3B341")
    text_color = colors.HexColor("#333333")
    muted_color = colors.HexColor("#666666")
    bg_table = colors.HexColor("#F8F9FA")

    # Estilos de Tipografia
    style_title = ParagraphStyle(
        "TitlePDF",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=primary_color,
        spaceAfter=4,
    )
    style_subtitle = ParagraphStyle(
        "SubtitlePDF",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=13,
        textColor=muted_color,
        spaceAfter=12,
    )
    style_heading = ParagraphStyle(
        "HeadingPDF",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=16,
        textColor=primary_color,
        spaceBefore=10,
        spaceAfter=4,
    )
    style_body = ParagraphStyle(
        "BodyPDF",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=12,
        textColor=text_color,
        spaceAfter=4,
    )
    style_table_header = ParagraphStyle(
        "TableHeaderPDF",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
    )
    style_table_cell = ParagraphStyle(
        "TableCellPDF",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=11,
        textColor=text_color,
    )

    story = []

    # Cabeçalho Principal
    story.append(Paragraph("Relatório de Auditoria de Dados", style_title))
    story.append(
        Paragraph(
            "Data Pipeline Validator | Análise de Integridade para Power Query",
            style_subtitle,
        )
    )
    story.append(
        HRFlowable(width="100%", thickness=1.5, color=accent_color, spaceAfter=12)
    )

    for report in reports:
        aba_nome = report.sheet_name or "Global Stream (CSV)"
        origem_nome = getattr(
            report, "source_file", getattr(report, "source_name", "Ficheiro Carregado")
        )

        story.append(
            Paragraph(f"<b>Partição / Aba:</b> {escape(str(aba_nome))}", style_heading)
        )
        story.append(
            Paragraph(
                f"<b>Origem:</b> {escape(str(origem_nome))} | "
                f"<b>Total de Linhas:</b> {report.total_rows} | "
                f"<b>Total de Colunas:</b> {report.total_cols}",
                style_body,
            )
        )

        crit_count = report.count_by_severity(SEVERITY_CRITICAL)
        warn_count = report.count_by_severity(SEVERITY_WARNING)
        info_count = report.count_by_severity(SEVERITY_INFO)

        resumo_text = (
            f"<b>Métricas de Auditoria:</b> "
            f"<font color='{critical_color.hexval()}'>Críticos: {crit_count}</font> | "
            f"<font color='{warning_color.hexval()}'>Avisos: {warn_count}</font> | "
            f"<font color='{primary_color.hexval()}'>Informativos: {info_count}</font>"
        )
        story.append(Paragraph(resumo_text, style_body))
        story.append(Spacer(1, 6))

        if report.column_rename_map:
            story.append(
                Paragraph(
                    "Mapeamento de Normalização de Colunas (snake_case)", style_heading
                )
            )
            table_data = [
                [
                    Paragraph("Coluna Original", style_table_header),
                    Paragraph("Coluna Padronizada", style_table_header),
                ]
            ]
            for orig, pad in report.column_rename_map.items():
                table_data.append(
                    [
                        Paragraph(escape(str(orig)), style_table_cell),
                        Paragraph(escape(str(pad)), style_table_cell),
                    ]
                )

            t = Table(table_data, colWidths=[240, 240])
            t.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), primary_color),
                        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                        ("TOPPADDING", (0, 0), (-1, -1), 5),
                        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [bg_table, colors.white]),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
                    ]
                )
            )
            story.append(t)
            story.append(Spacer(1, 8))

        story.append(Paragraph("Log de Ocorrências", style_heading))
        if not report.findings:
            story.append(
                Paragraph("Nenhuma anomalia registada nesta partição.", style_body)
            )
        else:
            findings_data = [
                [
                    Paragraph("Severidade", style_table_header),
                    Paragraph("Categoria", style_table_header),
                    Paragraph("Mensagem", style_table_header),
                ]
            ]
            for f in report.findings:
                sev_label = (
                    "CRÍTICO"
                    if f.severity == SEVERITY_CRITICAL
                    else ("AVISO" if f.severity == SEVERITY_WARNING else "INFO")
                )
                findings_data.append(
                    [
                        Paragraph(f"<b>{sev_label}</b>", style_table_cell),
                        Paragraph(escape(str(f.category)), style_table_cell),
                        Paragraph(escape(str(f.message)), style_table_cell),
                    ]
                )

            tf = Table(findings_data, colWidths=[70, 110, 300])
            tf.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), primary_color),
                        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                        ("TOPPADDING", (0, 0), (-1, -1), 5),
                        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [bg_table, colors.white]),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
                    ]
                )
            )
            story.append(tf)

        story.append(Spacer(1, 15))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()
