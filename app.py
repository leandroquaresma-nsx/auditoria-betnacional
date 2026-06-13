import os
import re
import base64
import pandas as pd
import openpyxl
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
from datetime import datetime

from flask import Flask, request, jsonify
from flask_cors import CORS
from google import genai

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

app = Flask(__name__)
CORS(app) 

CHAVE_API = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=CHAVE_API)

# Caminho onde o servidor vai procurar a sua logo (se existir)
LOGO_PATH = "logo.png"

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

def classificar_origem(motivo):
    motivo = str(motivo).lower()
    if any(k in motivo for k in ['senha', 'acesso', 'dados', 'mfa', 'usuário', 'esqueci', 'tutorial']):
        return 'Erro/Dúvida do Cliente'
    return 'Falha da Plataforma/Sistema'

# NOVO MOTOR: Calculadora de Tempo Médio de Atendimento (TMA)
def calcular_tma(conversa):
    if not isinstance(conversa, str): return 0
    tempos = re.findall(r'\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]', conversa)
    if len(tempos) >= 2:
        try:
            t_inicio = datetime.strptime(tempos[0], '%Y-%m-%d %H:%M:%S')
            t_fim = datetime.strptime(tempos[-1], '%Y-%m-%d %H:%M:%S')
            minutos = abs((t_fim - t_inicio).total_seconds() / 60.0)
            return minutos if minutos < 300 else 0 # Ignora outliers irreais (ex: tickets abertos dias)
        except: return 0
    return 0

@app.route("/auditar", methods=["POST"])
def auditar():
    try:
        if "file" not in request.files:
            return jsonify({"status": "erro", "mensagem": "Arquivo ausente."}), 400
            
        file = request.files["file"]
        df = pd.read_csv(file) if file.filename.endswith('.csv') else pd.read_excel(file)
        
        total_geral_casos = len(df)
        
        # Enriquecimento de Dados
        df['atendente_extraido'] = df['comments'].apply(cacar_nome_atendente)
        df['origem_erro'] = df['reason'].apply(classificar_origem)
        df['tags'] = df['tags'].fillna('')
        df['is_inativo'] = df['tags'].str.contains('inativo', case=False)
        df['tma_minutos'] = df['comments'].apply(calcular_tma) # Cálculo do tempo real
        
        for col in ['comments', 'ticket_summary']:
            if col in df.columns:
                df[col] = df[col].apply(limpar_dados_sensiveis)

        # KPIs Estratégicos
        qtd_inativos = int(df['is_inativo'].sum())
        pct_inativos = (qtd_inativos / total_geral_casos) * 100
        
        origem_counts = df['origem_erro'].value_counts()
        qtd_erro_cliente = int(origem_counts.get('Erro/Dúvida do Cliente', 0))
        pct_erro_cliente = (qtd_erro_cliente / total_geral_casos) * 100
        
        tma_medio = df[df['tma_minutos'] > 0]['tma_minutos'].mean()
        tma_medio = tma_medio if pd.notnull(tma_medio) else 0

        top_atendentes = df[df['atendente_extraido'] != "Não identificado"]['atendente_extraido'].value_counts().head(5)
        
        principais_motivos = df['reason'].value_counts().reset_index()
        principais_motivos.columns = ['Motivo_Ocorrencia', 'Quantidade']
        motivos_top5 = principais_motivos.head(5)
        top_motivo_nome = str(motivos_top5.iloc[0]['Motivo_Ocorrencia']).replace('_', ' ').title()
        top_motivo_qtd = motivos_top5.iloc[0]['Quantidade']

        # --- GERADOR DE GRÁFICOS ---
        cores_bet = ['#FF5A00', '#1C2541', '#3A506B', '#5BC0BE', '#94A3B8']
        
        # Gráfico 1: Motivos
        fig1, ax1 = plt.subplots(figsize=(6, 3))
        ax1.pie(motivos_top5['Quantidade'], colors=cores_bet, startangle=140, pctdistance=0.75, textprops=dict(color="white", weight="bold", fontsize=8))
        ax1.add_artist(plt.Circle((0,0), 0.55, fc='white'))
        ax1.legend([m.replace('_', ' ').title() for m in motivos_top5['Motivo_Ocorrencia']], loc="center left", bbox_to_anchor=(0.9, 0.5), frameon=False, fontsize=8)
        ax1.set_title('Top 5 Ocorrências (Volume)', fontsize=10, fontweight='bold', color='#0B132B', loc='left')
        plt.tight_layout()
        path_motivos = "/tmp/g_motivos.png"
        fig1.savefig(path_motivos, dpi=150)
        plt.close(fig1)

        # Gráfico 2: Origem
        fig2, ax2 = plt.subplots(figsize=(4, 3))
        ax2.pie(origem_counts, labels=origem_counts.index, autopct='%1.1f%%', colors=['#FF5A00', '#1C2541'], startangle=90, textprops=dict(color="white", weight="bold", fontsize=9))
        ax2.set_title('Responsabilidade do Erro', fontsize=10, fontweight='bold', color='#0B132B')
        plt.tight_layout()
        path_origem = "/tmp/g_origem.png"
        fig2.savefig(path_origem, dpi=150)
        plt.close(fig2)

        # --- CONSULTA À IA (AGORA COM TEMPO DE ATENDIMENTO) ---
        prompt = (f"Atue como Diretor de Operações da Betnacional. Analise estes dados auditados:\n"
                  f"- Total de Casos: {total_geral_casos}\n"
                  f"- Ocorrência Crítica: {top_motivo_nome} ({top_motivo_qtd} casos)\n"
                  f"- Abandono/Inatividade: {pct_inativos:.1f}%\n"
                  f"- Origem: {pct_erro_cliente:.1f}% falha do cliente vs {100-pct_erro_cliente:.1f}% da plataforma.\n"
                  f"- Tempo Médio de Atendimento (TMA): {tma_medio:.1f} minutos por cliente.\n\n"
                  f"Crie um laudo muito executivo, usando EXATAMENTE estes 4 subtítulos:\n"
                  f"1. AVALIAÇÃO DO GARGALO E TMA (Analise o impacto do tempo médio de {tma_medio:.1f} min e a volumetria)\n"
                  f"2. ATRITO E CULPA OPERACIONAL (Analise se a culpa é do cliente ou do sistema)\n"
                  f"3. ANÁLISE DE INATIVIDADE (Por que {pct_inativos:.1f}% abandonam o chat?)\n"
                  f"4. TRÊS AÇÕES DIRETIVAS (Recomendações curtas e grossas)")
        try:
            texto_ia = client.models.generate_content(model='gemini-2.5-flash', contents=prompt).text
        except:
            texto_ia = "Falha ao processar análise da IA."

        # --- GERADOR DE EXCEL TURBINADO COM LOGO E TMA ---
        nome_excel = "/tmp/Relatorio_Auditoria_Betnacional.xlsx"
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = 'Painel Executivo BI'
        ws1.views.sheetView[0].showGridLines = False
        
        # Inserir Logo se existir
        if os.path.exists(LOGO_PATH):
            img_logo = openpyxl.drawing.image.Image(LOGO_PATH)
            img_logo.width = 180
            img_logo.height = 45
            ws1.add_image(img_logo, 'B2')
            linha_inicio_kpi = 5
        else:
            ws1['B2'] = "PAINEL EXECUTIVO DE AUDITORIA - BETNACIONAL"
            ws1['B2'].font = Font(size=18, bold=True, color="0B132B")
            linha_inicio_kpi = 4
        
        borda = Border(left=Side(style='thin', color="CBD5E1"), right=Side(style='thin', color="CBD5E1"), top=Side(style='thin', color="CBD5E1"), bottom=Side(style='thin', color="CBD5E1"))
        
        # 5 KPIs agora!
        kpis = [
            ("TOTAL DE CASOS", total_geral_casos, f'B{linha_inicio_kpi}', f'B{linha_inicio_kpi+1}'),
            ("GARGALO ATUAL", f"{top_motivo_nome}", f'C{linha_inicio_kpi}', f'C{linha_inicio_kpi+1}'),
            ("ERRO CLIENTE (%)", f"{pct_erro_cliente:.1f}%", f'D{linha_inicio_kpi}', f'D{linha_inicio_kpi+1}'),
            ("INATIVIDADE (%)", f"{pct_inativos:.1f}%", f'E{linha_inicio_kpi}', f'E{linha_inicio_kpi+1}'),
            ("TMA (TEMPO MÉDIO)", f"{tma_medio:.1f} Minutos", f'F{linha_inicio_kpi}', f'F{linha_inicio_kpi+1}')
        ]
        
        for titulo, valor, c_tit, c_val in kpis:
            ws1[c_tit] = titulo
            ws1[c_tit].font = Font(size=9, bold=True, color="64748B")
            ws1[c_val] = valor
            ws1[c_val].font = Font(size=14, bold=True, color="FF5A00" if "TMA" in titulo else "0B132B")
            ws1[c_val].fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            ws1[c_val].border = borda

        ws1.column_dimensions['B'].width = 18
        ws1.column_dimensions['C'].width = 30
        ws1.column_dimensions['D'].width = 18
        ws1.column_dimensions['E'].width = 18
        ws1.column_dimensions['F'].width = 22

        ws1.add_image(openpyxl.drawing.image.Image(path_motivos), f'B{linha_inicio_kpi+3}')
        ws1.add_image(openpyxl.drawing.image.Image(path_origem), f'B{linha_inicio_kpi+18}')
        
        ws2 = wb.create_sheet('Base Enriquecida')
        ws2.append(list(df.columns))
        for _, row in df.iterrows(): ws2.append([str(item) for item in row])
        wb.save(nome_excel)

        # --- GERADOR DE PDF EXECUTIVO COM LOGO E TMA ---
        nome_pdf = "/tmp/Laudo_Auditoria.pdf"
        doc = SimpleDocTemplate(nome_pdf, pagesize=letter, rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=35)
        story = []
        styles = getSampleStyleSheet()
        
        st_tit = ParagraphStyle('Tit', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#0B132B'), spaceAfter=15)
        st_sub = ParagraphStyle('Sub', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#FF5A00'), spaceBefore=12, spaceAfter=8)
        st_txt = ParagraphStyle('Txt', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#1E293B'), leading=14, spaceAfter=8)
        
        # Cabeçalho com Logo
        if os.path.exists(LOGO_PATH):
            story.append(RLImage(LOGO_PATH, width=160, height=40, hAlign='LEFT'))
            story.append(Spacer(1, 15))
            
        story.append(Paragraph("LAUDO DE AUDITORIA OPERACIONAL", st_tit))
        story.append(Paragraph("Documento gerado automaticamente pelo Sistema de Inteligência Operacional.", st_txt))
        story.append(Spacer(1, 10))
        story.append(RLImage(path_motivos, width=380, height=190))
        story.append(Spacer(1, 10))
        
        for paragrafo in texto_ia.split('\n'):
            texto = paragrafo.replace('**', '').strip()
            if not texto: continue
            if texto[0].isdigit() and '.' in texto[:3]: 
                story.append(Paragraph(texto, st_sub))
            else:
                story.append(Paragraph(texto, st_txt))
                
        story.append(PageBreak()) 
        story.append(Paragraph("MÉTRICAS APROFUNDADAS DE ATRITO E PERFORMANCE", st_tit))
        
        tabela_graficos = Table([[RLImage(path_origem, width=220, height=165)]])
        story.append(tabela_graficos)
        
        doc.build(story)

        with open(nome_excel, "rb") as f: exc_64 = base64.b64encode(f.read()).decode('utf-8')
        with open(nome_pdf, "rb") as f: pdf_64 = base64.b64encode(f.read()).decode('utf-8')

        return jsonify({"status": "sucesso", "mensagem": "Auditoria de Alta Performance Finalizada!", "excel": exc_64, "pdf": pdf_64}), 200

    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
