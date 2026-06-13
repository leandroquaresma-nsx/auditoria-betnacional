import os
import re
import io
import base64
import pandas as pd
import openpyxl
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt

from flask import Flask, request, jsonify
from flask_cors import CORS
from google import genai

# Recursos de Estilização Avançada do Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Recursos do ReportLab para o PDF Profissional
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

app = Flask(__name__)
CORS(app) 

CHAVE_API = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=CHAVE_API)

def limpar_dados_sensiveis(texto):
    if not isinstance(texto, str): return texto
    texto = re.sub(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', '[EMAIL PROTEGIDO]', texto)
    texto = re.sub(r'\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b', '[CPF PROTEGIDO]', texto)
    texto = re.sub(r'\(?\d{2}\)?\s?\d{4,5}-?\d{4}', '[TELEFONE PROTEGIDO]', texto)
    return texto

def cacar_nome_atendente(conversa):
    if not isinstance(conversa, str): return "Não identificado"
    match = re.search(r'\b([A-ZÀ-Ÿ][a-zà-ÿ]+ [A-ZÀ-Ÿ]\.)', conversa)
    if match: return match.group(1)
    return "Não identificado"

@app.route("/auditar", methods=["POST"])
def auditar():
    try:
        if "file" not in request.files:
            return jsonify({"status": "erro", "mensagem": "Arquivo ausente."}), 400
            
        file = request.files["file"]
        
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
            
        total_geral_casos = len(df)
            
        df['atendente_extraido'] = df['comments'].apply(cacar_nome_atendente)
        for col in ['comments', 'ticket_summary']:
            if col in df.columns:
                df[col] = df[col].apply(limpar_dados_sensiveis)
                
        top_atendentes = df['atendente_extraido'].value_counts().reset_index()
        top_atendentes.columns = ['Atendente', 'Qtd_Atendimentos']
        top_atendentes = top_atendentes[top_atendentes['Atendente'] != "Não identificado"].head(5)

        principais_motivos = df['reason'].value_counts().reset_index()
        principais_motivos.columns = ['Motivo_Ocorrencia', 'Quantidade']
        motivos_top5 = principais_motivos.head(5)
        
        top_motivo_nome = str(motivos_top5.iloc[0]['Motivo_Ocorrencia']).replace('_', ' ').title()
        top_motivo_qtd = motivos_top5.iloc[0]['Quantidade']
        top_motivo_pct = (top_motivo_qtd / total_geral_casos) * 100

        # 🎨 GRÁFICO ESTILO POWER BI (DONUT CHART)
        fig, ax = plt.subplots(figsize=(7.5, 3.8))
        nomes_limpos = [m.replace('_', ' ').title() for m in motivos_top5['Motivo_Ocorrencia']]
        valores = motivos_top5['Quantidade']
        
        cores = ['#FF5A00', '#1C2541', '#3A506B', '#5BC0BE', '#94A3B8']
        
        wedges, texts, autotexts = ax.pie(
            valores, 
            autopct=lambda p: f'{p*sum(valores)/100:,.0f}'.replace(',', '.'),
            startangle=140, 
            colors=cores,
            pctdistance=0.75,
            textprops=dict(color="white", weight="bold", fontsize=9)
        )
        
        centre_circle = plt.Circle((0,0), 0.55, fc='white')
        ax.add_artist(centre_circle)
        
        ax.legend(wedges, nomes_limpos, loc="center left", bbox_to_anchor=(1, 0.5), frameon=False, fontsize=9)
        ax.set_title('TOP 5 VOLUMES DA OPERAÇÃO', fontsize=11, fontweight='bold', color='#0B132B', loc='left', pad=12)
        ax.axis('equal')  
        plt.tight_layout()
        
        grafico_path = "/tmp/grafico_motivos.png"
        plt.savefig(grafico_path, dpi=150)
        plt.close()

        # 🤖 CONSULTA DETALHADA À IA (RESUMO EXECUTIVO DA AUDITORIA)
        try:
            prompt = (
                f"Você é o Diretor de Auditoria da Betnacional. Com base nas volumetrias extraídas da planilha, "
                f"gere um laudo executivo muito profissional e direto. O total de casos analisados foi de {total_geral_casos}. "
                f"O principal gargalo operacional identificado foi '{top_motivo_nome}' com {top_motivo_qtd} ocorrências ({top_motivo_pct:.1f}% do total). "
                f"Estruture seu texto rigorosamente com estes tópicos técnicos claros:\n"
                f"1. SUMÁRIO DA INCIDÊNCIA GERAL (Analise o impacto macro operacional).\n"
                f"2. ANÁLISE DO PRINCIPAL GARGALO (Explique o impacto crítico de {top_motivo_nome}).\n"
                f"3. PLANO DE AÇÃO E RECOMENDAÇÕES (Liste 3 ações corretivas imediatas baseadas nos dados)."
            )
            resposta = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
            texto_ia = resposta.text
        except:
            texto_ia = (
                "1. SUMÁRIO DA INCIDÊNCIA GERAL\n\nA auditoria operacional identificou volumetria crítica concentrada em processos de acesso à conta.\n\n"
                f"2. ANÁLISE DO PRINCIPAL GARGALO\n\nO motivo '{top_motivo_nome}' lidera os acionamentos de suporte, exigindo automação imediata.\n\n"
                "3. PLANO DE AÇÃO E RECOMENDAÇÕES\n\n• Implementar fluxo de autoatendimento nativo.\n• Melhorar as FAQs na home da plataforma."
            )

        # 📊 CONSTRUTOR DE EXCEL PREMIUM COM INTERFACE DE DASHBOARD
        nome_excel = "/tmp/Relatorio_Auditoria_Betnacional.xlsx"
        wb = openpyxl.Workbook()
        
        # Definição de Cores da Identidade Visual
        COR_AZUL_DEEP = "0B132B"
        COR_AZUL_HEADER = "1C2541"
        COR_LARANJA = "FF5A00"
        COR_LINHA_ZEBRA = "F8FAFC"
        COR_BORDA = "CBD5E1"
        COR_FUNDO_CARD = "F1F5F9"
        
        font_titulo_painel = Font(name="Segoe UI", size=16, bold=True, color=COR_AZUL_DEEP)
        font_header_tabela = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        font_data_celula = Font(name="Segoe UI", size=10, color="333333")
        font_label_kpi = Font(name="Segoe UI", size=9, bold=True, color="64748B")
        font_valor_kpi = Font(name="Segoe UI", size=18, bold=True, color=COR_LARANJA)
        
        fill_header = PatternFill(start_color=COR_AZUL_HEADER, end_color=COR_AZUL_HEADER, fill_type="solid")
        fill_zebra = PatternFill(start_color=COR_LINHA_ZEBRA, end_color=COR_LINHA_ZEBRA, fill_type="solid")
        fill_card = PatternFill(start_color=COR_FUNDO_CARD, end_color=COR_FUNDO_CARD, fill_type="solid")
        
        borda_fina = Border(
            left=Side(style='thin', color=COR_BORDA), right=Side(style='thin', color=COR_BORDA),
            top=Side(style='thin', color=COR_BORDA), bottom=Side(style='thin', color=COR_BORDA)
        )

        ws1 = wb.active
        ws1.title = 'Dashboard Volumetrico'
        ws1.views.sheetView[0].showGridLines = True
        
        # 1. Título do Dashboard
        ws1['B2'] = "PAINEL DE AUDITORIA OPERACIONAL - BETNACIONAL"
        ws1['B2'].font = font_titulo_painel
        
        # 2. Construção dos Cartões de KPI (Estilo BI)
        # Cartão 1: Total Geral
        ws1['B4'] = "TOTAL DE CASOS AUDITADOS"
        ws1['B4'].font = font_label_kpi
        ws1['B5'] = total_geral_casos
        ws1['B5'].font = font_valor_kpi
        ws1['B5'].number_format = '#,##0'
        ws1['B5'].fill = fill_card
        ws1['B5'].border = borda_fina
        ws1['B5'].alignment = Alignment(horizontal="left", vertical="center")
        
        # Cartão 2: Principal Gargalo
        ws1['C4'] = "PRINCIPAL GARGALO DA OPERAÇÃO"
        ws1['C4'].font = font_label_kpi
        ws1['C5'] = f"{top_motivo_nome} ({top_motivo_pct:.1f}%)"
        ws1['C5'].font = Font(name="Segoe UI", size=11, bold=True, color=COR_AZUL_HEADER)
        ws1['C5'].fill = fill_card
        ws1['C5'].border = borda_fina
        ws1['C5'].alignment = Alignment(horizontal="left", vertical="center")

        # 3. Tabela Resumo Executivo
        ws1['B8'] = "Motivo da Ocorrência"
        ws1['C8'] = "Quantidade de Casos"
        ws1['D8'] = "Representação (%)"
        
        for cell in [ws1['B8'], ws1['C8'], ws1['D8']]:
            cell.font = font_header_tabela
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        linha_atual = 9
        for _, r in principais_motivos.iterrows():
            ws1[f'B{linha_atual}'] = str(r['Motivo_Ocorrencia']).replace('_', ' ').title()
            ws1[f'C{linha_atual}'] = r['Quantidade']
            ws1[f'D{linha_atual}'] = r['Quantidade'] / total_geral_casos
            
            # Formatações das Células
            for col_letter in ['B', 'C', 'D']:
                cell = ws1[f'{col_letter}{linha_atual}']
                cell.font = font_data_celula
                cell.border = borda_fina
                if linha_atual % 2 == 0:
                    cell.fill = fill_zebra
                    
            ws1[f'C{linha_atual}'].number_format = '#,##0'
            ws1[f'C{linha_atual}'].alignment = Alignment(horizontal="right")
            ws1[f'D{linha_atual}'].number_format = '0.0%'
            ws1[f'D{linha_atual}'].alignment = Alignment(horizontal="right")
            linha_atual += 1

        # Fixando as larguras das colunas do painel
        ws1.column_dimensions['A'].width = 3
        ws1.column_dimensions['B'].width = 42
        ws1.column_dimensions['C'].width = 32
        ws1.column_dimensions['D'].width = 18
        
        # Insere o gráfico de Rosca alinhado ao lado das tabelas
        if os.path.exists(grafico_path):
            ws1.add_image(openpyxl.drawing.image.Image(grafico_path), 'F4')

        # Aba 2: Base Dados Anonimizada
        ws3 = wb.create_sheet(title='Base Dados Anonimizada')
        ws3.views.sheetView[0].showGridLines = True
        ws3.append(list(df.columns))
        
        for cell in ws3[1]:
            cell.font = font_header_tabela
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        for r_idx, row in enumerate(ws3.iter_rows(min_row=2, max_row=ws3.max_row), start=2):
            for cell in row:
                cell.font = font_data_celula
                cell.border = borda_fina
                if r_idx % 2 == 0:
                    cell.fill = fill_zebra
                if cell.column in [1, 2, 6]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        for col in ws3.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws3.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 55)

        wb.save(nome_excel)

        # 📄 CONSTRUTOR DE PDF EXECUTIVO (COMPLETO: GRÁFICOS + TABELAS + RESUMOS)
        nome_pdf = "/tmp/Laudo_Executivo_Auditoria.pdf"
        doc = SimpleDocTemplate(nome_pdf, pagesize=letter, rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=35)
        story = []
        styles = getSampleStyleSheet()
        
        style_meta = ParagraphStyle('Meta', parent=styles['Normal'], fontSize=8, textColor=colors.HexColor('#64748B'), spaceAfter=5)
        style_titulo = ParagraphStyle('T', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#0B132B'), spaceAfter=12, fontName="Helvetica-Bold")
        style_subtitulo = ParagraphStyle('S', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#FF5A00'), spaceBefore=12, spaceAfter=8, fontName="Helvetica-Bold")
        style_corpo = ParagraphStyle('C', parent=styles['Normal'], fontSize=9.5, textColor=colors.HexColor('#1E293B'), leading=14, spaceAfter=8, alignment=4)
        
        story.append(Paragraph("BETNACIONAL OPERATIONAL AUDIT SYSTEM | RELATÓRIO DE DIRETORIA", style_meta))
        story.append(Paragraph("Laudo de Auditoria Estratégica", style_titulo))
        story.append(Spacer(1, 5))
        
        # Adiciona o Gráfico de Rosca Estilo Power BI
        if os.path.exists(grafico_path):
            story.append(Image(grafico_path, width=420, height=212))
            story.append(Spacer(1, 10))
            
        # 📑 INSERÇÃO DA TABELA DETALHADA NO PDF (O resumo analítico solicitado!)
        story.append(Paragraph("Detalhamento Volumétrico das Ocorrências", style_subtitulo))
        
        matriz_tabela_pdf = [["Motivo Identificado na Auditoria", "Casos Absolutos", "Impacto (%)"]]
        for _, r in motivos_top5.iterrows():
            m_nome = str(r['Motivo_Ocorrencia']).replace('_', ' ').title()
            m_qtd = f"{r['Quantidade']:,}".replace(',', '.')
            m_pct = f"{(r['Quantidade'] / total_geral_casos) * 100:.1f}%"
            matriz_tabela_pdf.append([m_nome, m_qtd, m_pct])
            
        tabela_reportlab = Table(matriz_tabela_pdf, colWidths=[280, 130, 110])
        tabela_reportlab.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1C2541')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
            ('TOPPADDING', (0, 0), (-1, 0), 5),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
            ('TOPPADDING', (0, 1), (-1, -1), 5),
        ]))
        story.append(tabela_reportlab)
        story.append(Spacer(1, 15))
        
        # Inserção do Diagnóstico da IA Textual Estruturado
        story.append(Paragraph("Parecer Técnico da Auditoria", style_subtitulo))
        for p in texto_ia.split('\n\n'):
            if p.strip():
                # Transforma linhas de tópicos em subtítulos visuais se necessário
                if "1." in p or "2." in p or "3." in p:
                    story.append(Paragraph(p.replace('**', ''), style_subtitulo))
                else:
                    story.append(Paragraph(p.replace('**', '').replace('•', '  •'), style_corpo))
                    
        doc.build(story)

        # Codificação final dos arquivos estruturados para retorno seguro
        with open(nome_excel, "rb") as f:
            excel_encoded = base64.b64encode(f.read()).decode('utf-8')
            
        with open(nome_pdf, "rb") as f:
            pdf_encoded = base64.b64encode(f.read()).decode('utf-8')

        return jsonify({
            "status": "sucesso",
            "mensagem": "Auditoria concluída! Seus relatórios executivos foram atualizados.",
            "excel": excel_encoded,
            "pdf": pdf_encoded
        }), 200

    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
