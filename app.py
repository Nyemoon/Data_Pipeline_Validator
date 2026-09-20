from __future__ import annotations
import base64
import io
import json
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException

# Importações para geração do PDF via ReportLab
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

from validador_planilhas import (
    EncodingDetectionError,
    SEVERITY_CRITICAL,
    SEVERITY_INFO,
    SEVERITY_WARNING,
    _load_csv,
    render_markdown_report,
    resolve_duplicate_sheet_names,
    run_single_validation,
)
from tratador_planilhas import tratar_dataframe_completo

# ---------------------------------------------------------------------------
# Configuração da página
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Data Pipeline Validator | Power BI",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Helpers para carregar imagens locais da pasta /imagens em Base64
# ---------------------------------------------------------------------------

def carregar_imagem_base64(caminho_relativo: str) -> str:
    caminho_completo = Path(__file__).parent / caminho_relativo
    if caminho_completo.exists():
        with open(caminho_completo, "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
            mime = "image/png" if caminho_completo.suffix.lower() == ".png" else "image/jpeg"
            return f"data:{mime};base64,{encoded}"
    return ""

# Caminhos configurados para a pasta /imagens
bg_image_path = "imagens/background.png"          
sidebar_bg_path = "imagens/runtime.png"          
logo_path = "imagens/logo.png"                   

bg_base64 = carregar_imagem_base64(bg_image_path)
sidebar_bg_base64 = carregar_imagem_base64(sidebar_bg_path)

# ---------------------------------------------------------------------------
# Função auxiliar otimizada para gerar PDF estruturado
# ---------------------------------------------------------------------------

def gerar_pdf_relatorio(reports: list) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter, 
        rightMargin=40, 
        leftMargin=40, 
        topMargin=40, 
        bottomMargin=40
    )
    
    styles = getSampleStyleSheet()
    
    # Paleta de Cores Alinhada ao Tema
    primary_color = colors.HexColor('#2A1B3D')  
    accent_color = colors.HexColor('#FF9E64')   
    critical_color = colors.HexColor('#FF7B72') 
    warning_color = colors.HexColor('#E3B341')  
    text_color = colors.HexColor('#333333')     
    muted_color = colors.HexColor('#666666')    
    bg_table = colors.HexColor('#F8F9FA')       
    
    # Estilos de Tipografia
    style_title = ParagraphStyle(
        'TitlePDF', parent=styles['Heading1'],
        fontName='Helvetica-Bold', fontSize=18, leading=22,
        textColor=primary_color, spaceAfter=4
    )
    style_subtitle = ParagraphStyle(
        'SubtitlePDF', parent=styles['Normal'],
        fontName='Helvetica', fontSize=9, leading=13,
        textColor=muted_color, spaceAfter=12
    )
    style_heading = ParagraphStyle(
        'HeadingPDF', parent=styles['Heading2'],
        fontName='Helvetica-Bold', fontSize=12, leading=16,
        textColor=primary_color, spaceBefore=10, spaceAfter=4
    )
    style_body = ParagraphStyle(
        'BodyPDF', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8.5, leading=12,
        textColor=text_color, spaceAfter=4
    )
    style_table_header = ParagraphStyle(
        'TableHeaderPDF', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=8.5, leading=11,
        textColor=colors.white
    )
    style_table_cell = ParagraphStyle(
        'TableCellPDF', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8, leading=11,
        textColor=text_color
    )

    story = []
    
    # Cabeçalho Principal
    story.append(Paragraph("Relatório de Auditoria de Dados", style_title))
    story.append(Paragraph("Data Pipeline Validator | Análise de Integridade para Power Query", style_subtitle))
    story.append(HRFlowable(width="100%", thickness=1.5, color=accent_color, spaceAfter=12))
    
    for report in reports:
        aba_nome = report.sheet_name or 'Global Stream (CSV)'
        origem_nome = getattr(report, 'source_name', 'Ficheiro Carregado')
        
        story.append(Paragraph(f"<b>Partição / Aba:</b> {aba_nome}", style_heading))
        story.append(Paragraph(f"<b>Origem:</b> {origem_nome} | <b>Total de Linhas:</b> {report.total_rows} | <b>Total de Colunas:</b> {report.total_cols}", style_body))
        
        crit_count = report.count_by_severity(SEVERITY_CRITICAL)
        warn_count = report.count_by_severity(SEVERITY_WARNING)
        info_count = report.count_by_severity(SEVERITY_INFO)
        
        resumo_text = f"<b>Métricas de Auditoria:</b> " \
                      f"<font color='{critical_color.hexval()}'>Críticos: {crit_count}</font> | " \
                      f"<font color='{warning_color.hexval()}'>Avisos: {warn_count}</font> | " \
                      f"<font color='{primary_color.hexval()}'>Informativos: {info_count}</font>"
        story.append(Paragraph(resumo_text, style_body))
        story.append(Spacer(1, 6))
        
        if report.column_rename_map:
            story.append(Paragraph("Mapeamento de Normalização de Colunas (snake_case)", style_heading))
            table_data = [[Paragraph("Coluna Original", style_table_header), Paragraph("Coluna Padronizada", style_table_header)]]
            for orig, pad in report.column_rename_map.items():
                table_data.append([
                    Paragraph(str(orig), style_table_cell),
                    Paragraph(str(pad), style_table_cell)
                ])
            
            t = Table(table_data, colWidths=[240, 240])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), primary_color),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [bg_table, colors.white]),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DDDDDD'))
            ]))
            story.append(t)
            story.append(Spacer(1, 8))
            
        story.append(Paragraph("Log de Ocorrências", style_heading))
        if not report.findings:
            story.append(Paragraph("Nenhuma anomalia registada nesta partição.", style_body))
        else:
            findings_data = [[Paragraph("Severidade", style_table_header), Paragraph("Categoria", style_table_header), Paragraph("Mensagem", style_table_header)]]
            for f in report.findings:
                sev_label = "CRÍTICO" if f.severity == SEVERITY_CRITICAL else ("AVISO" if f.severity == SEVERITY_WARNING else "INFO")
                findings_data.append([
                    Paragraph(f"<b>{sev_label}</b>", style_table_cell),
                    Paragraph(str(f.category), style_table_cell),
                    Paragraph(str(f.message), style_table_cell)
                ])
            
            tf = Table(findings_data, colWidths=[70, 110, 300])
            tf.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), primary_color),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [bg_table, colors.white]),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DDDDDD'))
            ]))
            story.append(tf)
            
        story.append(Spacer(1, 15))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ---------------------------------------------------------------------------
# Sistema visual com Glassmorphism Avançado
# ---------------------------------------------------------------------------

st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

        :root {{
            --surface-primary: rgba(20, 15, 25, 0.88);
            --surface-secondary: rgba(30, 22, 35, 0.92);
            --border-color: rgba(255, 180, 120, 0.25);
            --text-main: #FFF8F0;
            --text-muted: #D0B8A8;
            --accent-primary: #FF9E64;
            --accent-primary-dim: rgba(255, 158, 100, 0.18);
            --critical: #FF7B72;
            --critical-bg: rgba(255, 123, 114, 0.15);
            --warning: #E3B341;
            --warning-bg: rgba(227, 179, 65, 0.15);
            --success: #3FB950;
            --success-bg: rgba(63, 185, 80, 0.15);
        }}

        .stApp {{
            background: linear-gradient(135deg, rgba(15, 10, 20, 0.92) 0%, rgba(25, 15, 20, 0.77) 100%),
                        url('{bg_base64}');
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, rgba(30, 20, 30, 0.88) 0%, rgba(25, 15, 20, 0.60) 100%),
                        url('{sidebar_bg_base64}') !important;
            background-size: cover !important;
            background-position: center !important;
            border-right: 1px solid var(--border-color);
        }}

        .app-title {{
            font-family: 'Space Grotesk', sans-serif;
            font-size: 1.85rem;
            font-weight: 700;
            color: #ef9228;
            letter-spacing: -0.02em;
            margin-bottom: 0.2rem;
            text-shadow: 0 2px 6px rgba(0,0,0,0.7);
        }}
        .app-subtitle {{
            font-family: 'Space Grotesk', sans-serif;
            font-size: 0.95rem;
            color: #E8D5C8;
            margin-bottom: 2rem;
            max-width: 70ch;
            text-shadow: 0 1px 3px rgba(0,0,0,0.6);
        }}
        .section-header {{
            font-family: 'Space Grotesk', sans-serif;
            font-size: 1.1rem;
            font-weight: 600;
            color: #FFF8F0;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 0.4rem;
            margin-top: 2rem;
            margin-bottom: 1rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}

        .step-container, [data-testid="stMetric"] {{
            background-color: var(--surface-primary) !important;
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid var(--border-color) !important;
            border-radius: 10px;
            padding: 1.25rem;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
            transition: transform 0.3s ease, box-shadow 0.3s ease, border-color 0.3s ease;
            animation: fadeInUp 0.6s ease-out;
        }}

        .step-container:hover {{
            transform: translateY(-6px);
            box-shadow: 0 12px 40px 0 rgba(255, 158, 100, 0.15);
            border-color: var(--accent-primary) !important;
        }}

        @keyframes fadeInUp {{
            from {{ opacity: 0; transform: translateY(20px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        .step-index {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.75rem;
            font-weight: 600;
            color: var(--accent-primary);
            text-transform: uppercase;
            letter-spacing: 0.1em;
            margin-bottom: 0.5rem;
        }}
        .step-title {{
            font-family: 'Space Grotesk', sans-serif;
            font-size: 0.95rem;
            font-weight: 600;
            color: #FFF8F0;
            margin-bottom: 0.35rem;
        }}
        .step-desc {{
            font-size: 0.85rem;
            color: #E8D5C8;
            line-height: 1.5;
        }}

        .pipeline-banner {{
            border-radius: 8px;
            padding: 1rem 1.25rem;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.9rem;
            font-weight: 500;
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin: 1.25rem 0;
            border: 1px solid;
            backdrop-filter: blur(10px);
        }}
        .pipeline-ok {{
            background-color: var(--success-bg);
            color: #55FF70;
            border-color: rgba(63, 185, 80, 0.4);
        }}
        .pipeline-fail {{
            background-color: var(--critical-bg);
            color: #FF948A;
            border-color: rgba(255, 123, 114, 0.4);
        }}

        [data-testid="stMetricValue"] {{
            font-family: 'JetBrains Mono', monospace;
            font-weight: 600;
            color: #FFF8F0;
        }}
        [data-testid="stMetricLabel"] {{
            font-family: 'Space Grotesk', sans-serif;
            color: var(--text-muted);
            font-size: 0.8rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}

        .meta-badge {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.78rem;
            color: var(--accent-primary);
            background-color: var(--accent-primary-dim);
            border: 1px solid rgba(255, 158, 100, 0.35);
            border-radius: 4px;
            padding: 0.25rem 0.5rem;
            display: inline-block;
        }}

        [data-testid="stFileUploader"] button {{
            background-color: var(--success-bg) !important;
            color: #55FF70 !important;
            border: 1px solid rgba(63, 185, 80, 0.4) !important;
            border-radius: 6px !important;
            transition: all 0.3s ease !important;
        }}

        [data-testid="stFileUploader"] button:hover {{
            background-color: rgba(63, 185, 80, 0.3) !important;
            border-color: var(--success) !important;
            transform: translateY(-2px);
        }}

        div[data-testid="column"]:nth-of-type(1) [data-testid="stDownloadButton"] button {{
            background-color: var(--accent-primary-dim) !important;
            color: var(--accent-primary) !important;
            border: 1px solid rgba(255, 158, 100, 0.4) !important;
            border-radius: 6px !important;
            transition: all 0.3s ease !important;
        }}
        div[data-testid="column"]:nth-of-type(1) [data-testid="stDownloadButton"] button:hover {{
            background-color: rgba(255, 158, 100, 0.3) !important;
            border-color: var(--accent-primary) !important;
            transform: translateY(-2px);
        }}

        div[data-testid="column"]:nth-of-type(2) [data-testid="stDownloadButton"] button {{
            background-color: var(--success-bg) !important;
            color: #55FF70 !important;
            border: 1px solid rgba(63, 185, 80, 0.4) !important;
            border-radius: 6px !important;
            transition: all 0.3s ease !important;
        }}
        div[data-testid="column"]:nth-of-type(2) [data-testid="stDownloadButton"] button:hover {{
            background-color: rgba(63, 185, 80, 0.3) !important;
            border-color: var(--success) !important;
            transform: translateY(-2px);
        }}

        .sidebar-logo-container {{
            display: flex;
            justify-content: center;
            margin-bottom: 1.5rem;
        }}
        .sidebar-logo-container img {{
            width: 85px;
            height: 85px;
            object-fit: cover;
            border-radius: 50%;
            border: 2px solid var(--accent-primary);
            box-shadow: 0 4px 14px rgba(0,0,0,0.5);
        }}

        [data-testid="stMarkdownContainer"] {{ color: var(--text-main); }}
        [data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li {{ color: var(--text-main); }}
        [data-testid="stCaptionContainer"] {{ color: var(--text-muted) !important; }}
        [data-testid="stWidgetLabel"] p {{ color: var(--text-main); }}
        [data-testid="stFileUploaderDropzone"] {{
            background-color: var(--surface-secondary) !important;
            border-color: var(--border-color) !important;
        }}
        [data-testid="stFileUploaderDropzone"] span, [data-testid="stFileUploaderDropzone"] small {{
            color: var(--text-muted) !important;
        }}
        [data-testid="stTabs"] [data-baseweb="tab"] p {{ color: var(--text-muted); }}
        [data-testid="stTabs"] [aria-selected="true"] p {{ color: var(--accent-primary) !important; }}
        [data-testid="stExpander"] {{
            background-color: var(--surface-secondary) !important;
            border: 1px solid var(--border-color) !important;
            border-radius: 8px;
        }}
        [data-testid="stExpander"] summary p {{ color: var(--text-main) !important; }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"], [data-testid="stSidebar"] label p {{
            color: var(--text-main);
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Cabeçalho Principal
# ---------------------------------------------------------------------------

st.markdown('<div class="app-title">Data Ingestion Validator</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-subtitle">Inspeção estática, normalização de esquemas e validação '
    'de restrições estruturais para fontes tabulares antes do carregamento no Power Query.</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Sidebar de Parâmetros
# ---------------------------------------------------------------------------

with st.sidebar:
    if (Path(__file__).parent / logo_path).exists():
        logo_base64 = carregar_imagem_base64(logo_path)
        st.markdown(
            f'<div class="sidebar-logo-container"><img src="{logo_base64}" alt="Logo"></div>',
            unsafe_allow_html=True,
        )

    st.markdown("### Runtime Config")
    st.caption("Injeção de contratos de dados opcionais.")

    uploaded_config = st.file_uploader(
        "Schema Definition (config.json)",
        type=["json"],
        help="Contrato JSON contendo colunas obrigatórias, tipagens e chaves primárias.",
    )

    with st.expander("Esquema de Referência"):
        st.code(
            """{
  "colunas_obrigatorias": ["id_pedido", "data_venda"],
  "tipos_esperados": {
    "id_pedido": "numero",
    "data_venda": "data"
  },
  "chave_duplicata": "id_pedido"
}""",
            language="json",
        )

config_data = {}
if uploaded_config:
    try:
        config_data = json.load(uploaded_config)
        chaves_validas = {"colunas_obrigatorias", "tipos_esperados", "chave_duplicata"}
        chaves_extras = set(config_data.keys()) - chaves_validas
        if chaves_extras:
            st.sidebar.warning(f"Atributos desconhecidos ignorados: {', '.join(sorted(chaves_extras))}")
        st.sidebar.success("Contrato de esquema carregado.")
    except Exception as err:
        st.sidebar.error(f"Erro no parse do JSON: {err}")

required_columns = config_data.get("colunas_obrigatorias", [])
expected_types = config_data.get("tipos_esperados", {})
duplicate_key = config_data.get("chave_duplicata", None)

# ---------------------------------------------------------------------------
# Layout de Upload & Execução
# ---------------------------------------------------------------------------

col_source, col_ctrl = st.columns([2, 1], gap="medium")

with col_source:
    uploaded_file = st.file_uploader(
        "Selecionar origem de dados",
        type=["xlsx", "xls", "csv"],
        help="Suporte a ficheiros tabulares estritos (.xlsx, .xls, .csv)",
    )

if not uploaded_file:
    st.markdown('<div class="section-header">Fluxo Operacional</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3, gap="medium")
    with c1:
        st.markdown(
            '<div class="step-container"><div class="step-index">Phase 01</div>'
            '<div class="step-title">Inbound Source</div>'
            '<div class="step-desc">Carregamento de ficheiros tabulares brutos com deteção automática de codificação e delimitadores.</div>'
            '</div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            '<div class="step-container"><div class="step-index">Phase 02</div>'
            '<div class="step-title">Schema Validation</div>'
            '<div class="step-desc">Varredura de nulos, desvios de tipagem, unicidade de chaves e padronização para snake_case.</div>'
            '</div>',
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            '<div class="step-container"><div class="step-index">Phase 03</div>'
            '<div class="step-title">Sanitized Export</div>'
            '<div class="step-desc">Geração de relatórios de auditoria em Markdown/PDF e pacotes de dados prontos para consumo analítico.</div>'
            '</div>',
            unsafe_allow_html=True,
        )

if uploaded_file:
    file_extension = Path(uploaded_file.name).suffix.lower()
    current_signature = (uploaded_file.name, uploaded_file.size)

    with col_ctrl:
        st.markdown("### Execução")
        st.markdown(
            f'<div class="meta-badge">{uploaded_file.name} | {uploaded_file.size / 1024:.1f} KB</div>',
            unsafe_allow_html=True,
        )
        st.write("")
        run_validation = st.button("Executar Auditoria", type="primary", use_container_width=True)

    if run_validation:
        with st.status("Processando pipeline de dados...", expanded=False) as status:
            try:
                status.write("A ler fluxo de dados...")
                if file_extension == ".csv":
                    df = _load_csv(uploaded_file)
                    dfs_to_validate = {None: df}
                    sheet_save_keys: dict[str | None, str] = {None: "default"}
                else:
                    xl = pd.ExcelFile(uploaded_file)
                    sheet_save_keys = resolve_duplicate_sheet_names(xl.sheet_names)
                    dfs_to_validate = {
                        sname: pd.read_excel(xl, sheet_name=sname)
                        for sname in xl.sheet_names
                    }

                status.write("A aplicar regras de integridade e contratos...")
                reports = []
                dfs_cleaned = {}

                for sname, df in dfs_to_validate.items():
                    df_clean, report = run_single_validation(
                        df,
                        source_name=uploaded_file.name,
                        sheet_name=sname,
                        required_columns=required_columns,
                        expected_types=expected_types,
                        duplicate_key=duplicate_key,
                    )
                    # Injetando o nome de origem caso não venha nativo no report
                    report.source_name = uploaded_file.name
                    
                    save_key = sheet_save_keys[sname]
                    if save_key != sname and sname is not None:
                        report.add(
                            SEVERITY_WARNING,
                            "Namespace Collision",
                            f"Aba '{sname}' normalizada para '{save_key}' devido a conflito de nomenclatura case-insensitive.",
                        )
                    reports.append(report)

                    df_tratado = tratar_dataframe_completo(
                        df_clean,
                        chave_duplicata=duplicate_key,
                        tipos_esperados=expected_types,
                    )
                    dfs_cleaned[save_key] = df_tratado

                status.update(label="Auditoria concluída com sucesso", state="complete")

                st.session_state["validation_result"] = {
                    "reports": reports,
                    "dfs_cleaned": dfs_cleaned,
                    "file_extension": file_extension,
                    "signature": current_signature,
                }

            except EncodingDetectionError:
                status.update(label="Falha de Codificação", state="error")
                st.error("Falha ao resolver codificação do ficheiro CSV. Reencarregue o ficheiro em UTF-8.")
                st.session_state.pop("validation_result", None)
            except (zipfile.BadZipFile, InvalidFileException, ValueError):
                status.update(label="Formato Inválido", state="error")
                st.error("Estrutura do ficheiro corrompida ou formato incompatível com os parsers suportados.")
                st.session_state.pop("validation_result", None)
            except Exception as e:
                status.update(label="Erro Crítico de Execução", state="error")
                st.error(f"Exceção não tratada durante o processamento: {e}")
                st.session_state.pop("validation_result", None)

    result = st.session_state.get("validation_result")

    if result and result["signature"] != current_signature:
        st.info("Source alterado. Acione 'Executar Auditoria' para recalcular o pipeline.")
    elif result:
        reports = result["reports"]
        dfs_cleaned = result["dfs_cleaned"]
        result_file_extension = result["file_extension"]

        total_criticos = sum(r.count_by_severity(SEVERITY_CRITICAL) for r in reports)
        total_avisos = sum(r.count_by_severity(SEVERITY_WARNING) for r in reports)

        st.markdown('<div class="section-header">Relatório de Auditoria</div>', unsafe_allow_html=True)

        if total_criticos == 0:
            st.markdown(
                '<div class="pipeline-banner pipeline-ok">'
                '<span>STATUS: PASS (Integridade estrutural validada)</span>'
                '<span>0 CRITICAL ERRORS</span>'
                '</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f'<div class="pipeline-banner pipeline-fail">'
                f'<span>STATUS: FAILED (Inconsistências críticas detetadas)</span>'
                f'<span>{total_criticos} CRITICAL ISSUE(S)</span>'
                '</div>',
                unsafe_allow_html=True,
            )

        m1, m2, m3 = st.columns(3)
        m1.metric("Partições / Abas", len(reports))
        m2.metric("Erros Críticos", total_criticos)
        m3.metric("Avisos", total_avisos)

        st.write("")

        tab_labels = [
            (save_key if r.sheet_name is not None else "Default Stream")
            for r, save_key in zip(reports, dfs_cleaned.keys())
        ]
        tabs = st.tabs(tab_labels)

        for i, (report, df_final) in enumerate(zip(reports, dfs_cleaned.values())):
            with tabs[i]:
                st.markdown(f"**Partição:** `{report.sheet_name or 'Global Stream'}`")

                rows_in = report.total_rows
                rows_out = len(df_final)
                diff_rows = rows_in - rows_out
                vol_str = f"Linhas: {rows_in} original | {rows_out} sanitizada"
                if diff_rows > 0:
                    vol_str += f" ({diff_rows} eliminadas)"
                vol_str += f" | Colunas: {report.total_cols}"

                st.markdown(f'<div class="meta-badge">{vol_str}</div>', unsafe_allow_html=True)
                st.write("")

                if report.column_rename_map:
                    st.markdown("##### Mapeamento de Normalização de Colunas")
                    rename_df = pd.DataFrame(
                        list(report.column_rename_map.items()),
                        columns=["Original", "Padronizado (snake_case)"],
                    )
                    st.dataframe(rename_df, use_container_width=True, hide_index=True)

                st.markdown("##### Log de Ocorrências")
                if not report.findings:
                    st.info("Nenhuma anomalia registada nesta partição.")
                else:
                    groups = [
                        (SEVERITY_CRITICAL, "Crítico", st.error, True),
                        (SEVERITY_WARNING, "Aviso", st.warning, True),
                        (SEVERITY_INFO, "Informativo", st.info, False),
                    ]
                    for sev, label_sev, comp, default_exp in groups:
                        subset = [f for f in report.findings if f.severity == sev]
                        if not subset:
                            continue
                        with st.expander(f"[{label_sev.upper()}] - {len(subset)} ocorrência(s)", expanded=default_exp):
                            for f in subset:
                                comp(f"[{f.category}] {f.message}")

        st.markdown('<div class="section-header">Artifacts & Export</div>', unsafe_allow_html=True)

        d1, d2 = st.columns(2)

        markdown_report = render_markdown_report(reports)

        with d1:
            report_format = st.radio("Formato do Relatório", ["Markdown (.md)", "PDF (.pdf)"], horizontal=True, label_visibility="collapsed")
            if report_format == "Markdown (.md)":
                st.download_button(
                    label="Baixar Relatório de Auditoria (.md)",
                    data=markdown_report,
                    file_name="audit_report.md",
                    mime="text/markdown",
                    use_container_width=True,
                    key="btn_down_md",
                )
            else:
                pdf_bytes = gerar_pdf_relatorio(reports)
                st.download_button(
                    label="Baixar Relatório de Auditoria (.pdf)",
                    data=pdf_bytes,
                    file_name="audit_report.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="btn_down_pdf",
                )

        with d2:
            st.write("") 
            if len(dfs_cleaned) == 1 and result_file_extension == ".csv":
                single_df = next(iter(dfs_cleaned.values()))
                csv_bytes = single_df.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    label="Baixar Dataset Tratado (.csv)",
                    data=csv_bytes,
                    file_name="dataset_sanitized.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="btn_down_csv",
                )
            else:
                buffer = io.BytesIO()
                if len(dfs_cleaned) == 1:
                    single_df = next(iter(dfs_cleaned.values()))
                    single_df.to_excel(buffer, index=False)
                    out_name = "dataset_sanitized.xlsx"
                else:
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        for sname, df_s in dfs_cleaned.items():
                            df_s.to_excel(writer, sheet_name=sname, index=False)
                    out_name = "dataset_multitab_sanitized.xlsx"

                st.download_button(
                    label="Baixar Dataset Tratado (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_down_xlsx",
                )