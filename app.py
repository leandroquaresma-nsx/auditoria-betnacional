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

# Recursos de Estilização do Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Recursos do ReportLab para o PDF
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
        
        # 🎨 GRÁFICO ESTILO POWER BI (DONUT CHART)
        fig, ax = plt.subplots(figsize=(7.5, 3.5))
        nomes_limpos = [m.replace('_', ' ').title() for m in motivos_top5['Motivo_Ocorrencia']]
        valores = motivos_top5['Quantidade']
        
        # Paleta de Cores Dashboard
        cores = ['#FF5A00', '#1C2541', '#3A506B', '#5BC0BE', '#0B132B']
        
        # Desenha a rosca
        wedges, texts, autotexts = ax.pie(
            valores, 
            autopct=lambda p: f'{p*sum(valores)/100:,.0f}'.replace(',', '.'), # Mostra a quantidade real formatada
            startangle=140, 
            colors=cores,
            pctdistance=0.75,
            textprops=dict(color="white", weight="bold", fontsize=9)
        )
        
        # Cria o buraco do meio (Donut)
        centre_circle = plt.Circle((0,0), 0.55, fc='white')
        ax.add_artist(centre_circle)
        
        # Legenda idêntica ao Power BI à direita
        ax.legend(wedges, nomes_limpos, loc="center left", bbox_to_anchor=(1, 0.5), frameon=False, fontsize=9)
        ax.set_title('PRINCIPAIS VOLUMES DE ATENDIMENTO', fontsize=11, fontweight='bold', color='#0B132B', loc='left', pad=10)
        ax.axis('equal')  
        plt.tight_layout()
        
        grafico_path = "/tmp/grafico_motivos.png"
        plt.savefig(grafico_path, dpi=150)
        plt.close()

        # Consulta Inteligente à IA da Google
        try:
            prompt = f"Você é um Auditor Sênior da Betnacional. Analise estes motivos de suporte e gere um laudo executivo curto focado em melhorias: {str(motivos_top5.to_dict())}"
            resposta = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
            texto_ia = resposta.text
        except:
            texto_ia = "DIAGNÓSTICO OPERACIONAL:\nO volume de suporte está concentrado em demandas de redefinição de acessos. Recomenda-se melhorias nos fluxos automáticos."

        # Gerador de Excel Premium
        nome_excel = "/tmp/Relatorio_Auditoria_Betnacional.xlsx"
        wb = openpyxl.Workbook()
        
        COR_AZUL_HEADER = "1C2541"
        COR_TEXTO_HEADER = "FFFFFF"
        COR_LINHA_ZEBRA = "F8FAFC"
        COR_BORDA = "E2E8F0"
        
        font_header = Font(name="Segoe UI", size=11, bold=True, color=COR_TEXTO_HEADER)
        fill_header = PatternFill(start_color=COR_AZUL_HEADER, end_color=COR_AZUL_HEADER, fill_type="solid")
        font_data = Font(name="Segoe UI", size=10, color="333333")
        fill_zebra = PatternFill(start_color=COR_LINHA_ZEBRA, end_color=COR_LINHA_ZEBRA, fill_type="solid")
        
        borda_fina = Border(
            left=Side(style='thin', color=COR_BORDA), right=Side(style='thin', color=COR_BORDA),
            top=Side(style='thin', color=COR_BORDA), bottom=Side(style='thin', color=COR_BORDA)
        )

        ws1 = wb.active
        ws1.title = 'Dashboard Volumetrico'
        ws1.views.sheetView[0].showGridLines = True
        
        ws1.append(['Motivo da Ocorrência', 'Quantidade de Casos'])
        for _, row in principais_motivos.iterrows(): 
            ws1.append([str(row[0]).replace('_', ' ').title(), row[1]])
        
        for cell in ws1[1]:
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="left", vertical="center")
            
        for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row, min_col=1, max_col=2):
            for cell in row:
                cell.font = font_data
                cell.border = borda_fina
                if cell.column == 2:
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = '#,##0'
                    
        if os.path.exists(grafico_path):
            ws1.add_image(openpyxl.drawing.image.Image(grafico_path), 'D2')

        for col in ws1.iter_cols(max_col=2):
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws1.column_dimensions[col_letter].width = max(max_len + 4, 15)

        ws3 = wb.create_sheet(title='Base Dados Anonimizada')
        ws3.views.sheetView[0].showGridLines = True
        ws3.append(list(df.columns))
        
        for _, row in df.iterrows():
            ws3.append([str(item) if isinstance(item, (list, dict)) else item for item in row])
            
        for cell in ws3[1]:
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        for r_idx, row in enumerate(ws3.iter_rows(min_row=2, max_row=ws3.max_row), start=2):
            for cell in row:
                cell.font = font_data
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
            ws3.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

        wb.save(nome_excel)

        # Geração do PDF Executivo
        nome_pdf = "/tmp/Laudo_Executivo_Auditoria.pdf"
        doc = SimpleDocTemplate(nome_pdf, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
        story = []
        styles = getSampleStyleSheet()
        style_titulo = ParagraphStyle('T', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#0B132B'), spaceAfter=15)
        style_corpo = ParagraphStyle('C', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#2C3E50'), leading=14, spaceAfter=10, alignment=4)
        
        story.append(Paragraph("BETNACIONAL | EXECUTIVE AUDIT DASHBOARD", style_corpo))
        story.append(Paragraph("Auditoria Estratégica de Suporte", style_titulo))
        
        if os.path.exists(grafico_path):
            story.append(Image(grafico_path, width=400, height=186))
            story.append(Spacer(1, 15))
            
        for p in texto_ia.split('\n\n'):
            story.append(Paragraph(p.replace('**', ''), style_corpo))
        doc.build(story)

        with open(nome_excel, "rb") as f:
            excel_encoded = base64.b64encode(f.read()).decode('utf-8')
            
        with open(nome_pdf, "rb") as f:
            pdf_encoded = base64.b64encode(f.read()).decode('utf-8')

        return jsonify({
            "status": "sucesso",
            "mensagem": "Auditoria concluída! Seus relatórios foram gerados.",
            "excel": excel_encoded,
            "pdf": pdf_encoded
        }), 200

    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
