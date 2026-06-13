import os
import re
import base64
import pandas as pd
import openpyxl
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt

from flask import Flask, request, jsonify
from flask_cors import CORS
from google import genai

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
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

# Algoritmo para definir a origem do erro (Cliente vs Plataforma)
def classificar_origem(motivo):
    motivo = str(motivo).lower()
    # Palavras que indicam falha/dúvida do usuário
    if any(k in motivo for k in ['senha', 'acesso', 'dados', 'mfa', 'usuário', 'esqueci', 'tutorial']):
        return 'Erro/Dúvida do Cliente'
    # O resto é classificado como falha estrutural, sistema, bug ou offline
    return 'Falha da Plataforma/Sistema'

@app.route("/auditar", methods=["POST"])
def auditar():
    try:
        if "file" not in request.files:
            return jsonify({"status": "erro", "mensagem": "Arquivo ausente."}), 400
            
        file = request.files["file"]
        df = pd.read_csv(file) if file.filename.endswith('.csv') else pd.read_excel(file)
        
        total_geral_casos = len(df)
        
        # Processamento e Limpeza
        df['atendente_extraido'] = df['comments'].apply(cacar_nome_atendente)
        df['origem_erro'] = df['reason'].apply(classificar_origem)
        df['tags'] = df['tags'].fillna('')
        df['is_inativo'] = df['tags'].str.contains('inativo', case=False)
        
        for col in ['comments', 'ticket_summary']:
            if col in df.columns:
                df[col] = df[col].apply(limpar_dados_sensiveis)

        # 1. KPI Inatividade
        qtd_inativos = int(df['is_inativo'].sum())
        pct_inativos = (qtd_inativos / total_geral_casos) * 100

        # 2. KPI Origem do Erro
        origem_counts = df['origem_erro'].value_counts()
        qtd_erro_cliente = int(origem_counts.get('Erro/Dúvida do Cliente', 0))
        pct_erro_cliente = (qtd_erro_cliente / total_geral_casos) * 100

        # 3. KPI Atendentes
        top_atendentes = df[df['atendente_extraido'] != "Não identificado"]['atendente_extraido'].value_counts().head(5)
        
        # 4. KPI Motivos
        principais_motivos = df['reason'].value_counts().reset_index()
        principais_motivos.columns = ['Motivo_Ocorrencia', 'Quantidade']
        motivos_top5 = principais_motivos.head(5)
        top_motivo_nome = str(motivos_top5.iloc[0]['Motivo_Ocorrencia']).replace('_', ' ').title()
        top_motivo_qtd = motivos_top5.iloc[0]['Quantidade']

        # --- GERADOR DE GRÁFICOS (MATPLOTLIB) ---
        cores_bet = ['#FF5A00', '#1C2541', '#3A506B', '#5BC0BE', '#94A3B8']
        
        # Gráfico 1: Top Motivos (Donut)
        fig1, ax1 = plt.subplots(figsize=(6, 3))
        ax1.pie(motivos_top5['Quantidade'], colors=cores_bet, startangle=140, pctdistance=0.75,
                textprops=dict(color="white", weight="bold", fontsize=8))
        ax1.add_artist(plt.Circle((0,0), 0.55, fc='white'))
        ax1.legend([m.replace('_', ' ').title() for m in motivos_top5['Motivo_Ocorrencia']], 
                   loc="center left", bbox_to_anchor=(0.9, 0.5), frameon=False, fontsize=8)
        ax1.set_title('Top 5 Motivos de Acionamento', fontsize=10, fontweight='bold', color='#0B132B', loc='left')
        plt.tight_layout()
        path_motivos = "/tmp/g_motivos.png"
        fig1.savefig(path_motivos, dpi=150)
        plt.close(fig1)

        # Gráfico 2: Erro Cliente vs Plataforma (Pie)
        fig2, ax2 = plt.subplots(figsize=(4, 3))
        ax2.pie(origem_counts, labels=origem_counts.index, autopct='%1.1f%%', colors=['#FF5A00', '#1C2541'],
                startangle=90, textprops=dict(color="white", weight="bold", fontsize=9))
        ax2.set_title('Origem do Atrito', fontsize=10, fontweight='bold', color='#0B132B')
        plt.tight_layout()
        path_origem = "/tmp/g_origem.png"
        fig2.savefig(path_origem, dpi=150)
        plt.close(fig2)

        # Gráfico 3: Taxa de Inatividade (Pie)
        fig3, ax3 = plt.subplots(figsize=(4, 3))
        ax3.pie([total_geral_casos - qtd_inativos, qtd_inativos], labels=['Ativos/Resolvidos', 'Abandono (Inativos)'], 
                autopct='%1.1f%%', colors=['#5BC0BE', '#94A3B8'], startangle=90, textprops=dict(weight="bold", fontsize=9))
        ax3.set_title('Taxa de Inatividade', fontsize=10, fontweight='bold', color='#0B132B')
        plt.tight_layout()
        path_inativos = "/tmp/g_inativos.png"
        fig3.savefig(path_inativos, dpi=150)
        plt.close(fig3)

        # Gráfico 4: Top Atendentes (Bar Horizontal)
        fig4, ax4 = plt.subplots(figsize=(5, 3))
        ax4.barh(top_atendentes.index[::-1], top_atendentes.values[::-1], color='#3A506B')
        ax4.set_title('Performance: Top 5 Atendentes', fontsize=10, fontweight='bold', color='#0B132B', loc='left')
        ax4.spines['top'].set_visible(False)
        ax4.spines['right'].set_visible(False)
        plt.tight_layout()
        path_atendentes = "/tmp/g_atendentes.png"
        fig4.savefig(path_atendentes, dpi=150)
        plt.close(fig4)

        # --- CONSULTA À IA DA GOOGLE (RELATÓRIO PROFUNDO) ---
        prompt = (f"Atue como um Consultor de Operações Sênior da Betnacional. Analise os seguintes dados reais da operação:\n"
                  f"- Total de Casos: {total_geral_casos}\n"
                  f"- Ocorrência Crítica: {top_motivo_nome} ({top_motivo_qtd} casos)\n"
                  f"- Taxa de Abandono/Inatividade pelo usuário: {pct_inativos:.1f}%\n"
                  f"- Origem dos Problemas: {pct_erro_cliente:.1f}% são dúvidas/erros do cliente (ex: senha) vs "
                  f"{100-pct_erro_cliente:.1f}% falhas da plataforma.\n"
                  f"- Top Atendentes: {dict(top_atendentes)}\n\n"
                  f"Escreva um laudo de auditoria impecável, formal e analítico, usando EXATAMENTE estes 4 subtítulos:\n"
                  f"1. DIAGNÓSTICO DO GARGALO (Fale sobre a volumetria e ofensores)\n"
                  f"2. ATRITO E INATIVIDADE (Analise a culpa do cliente x plataforma e o abandono de chat)\n"
                  f"3. PERFORMANCE DO TIME (Destaque o trabalho dos atendentes citados)\n"
                  f"4. PLANO DE AÇÃO ESTRATÉGICO (Liste 3 recomendações claras e aplicáveis)")
        
        try:
            texto_ia = client.models.generate_content(model='gemini-2.5-flash', contents=prompt).text
        except:
            texto_ia = "1. DIAGNÓSTICO DO GARGALO\nFalha ao consultar IA.\n\n2. ATRITO E INATIVIDADE\nFalha ao consultar IA.\n\n3. PERFORMANCE DO TIME\nFalha ao consultar IA.\n\n4. PLANO DE AÇÃO ESTRATÉGICO\nFalha ao consultar IA."

        # --- GERADOR DE EXCEL PREMIUM ---
        nome_excel = "/tmp/Relatorio_Auditoria_Betnacional.xlsx"
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = 'Dashboard Operacional'
        ws1.views.sheetView[0].showGridLines = False
        
        borda = Border(left=Side(style='thin', color="CBD5E1"), right=Side(style='thin', color="CBD5E1"),
                       top=Side(style='thin', color="CBD5E1"), bottom=Side(style='thin', color="CBD5E1"))

        ws1['B2'] = "PAINEL EXECUTIVO DE AUDITORIA - BETNACIONAL"
        ws1['B2'].font = Font(size=16, bold=True, color="0B132B")
        
        # Cartões de KPI
        kpis = [
            ("TOTAL DE CASOS", total_geral_casos, 'B4', 'B5'),
            ("GARGALO PRINCIPAL", f"{top_motivo_nome}", 'C4', 'C5'),
            ("ERRO DO CLIENTE (%)", f"{pct_erro_cliente:.1f}%", 'D4', 'D5'),
            ("INATIVIDADE (%)", f"{pct_inativos:.1f}%", 'E4', 'E5')
        ]
        
        for titulo, valor, c_tit, c_val in kpis:
            ws1[c_tit] = titulo
            ws1[c_tit].font = Font(size=9, bold=True, color="64748B")
            ws1[c_val] = valor
            ws1[c_val].font = Font(size=14, bold=True, color="FF5A00")
            ws1[c_val].fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            ws1[c_val].border = borda

        ws1.column_dimensions['B'].width = 20
        ws1.column_dimensions['C'].width = 35
        ws1.column_dimensions['D'].width = 20
        ws1.column_dimensions['E'].width = 20

        # Inserir Gráficos no Excel
        ws1.add_image(openpyxl.drawing.image.Image(path_motivos), 'B7')
        ws1.add_image(openpyxl.drawing.image.Image(path_origem), 'B23')
        ws1.add_image(openpyxl.drawing.image.Image(path_inativos), 'D23')

        # Aba de Dados
        ws2 = wb.create_sheet('Base Enriquecida')
        ws2.append(list(df.columns))
        for _, row in df.iterrows():
            ws2.append([str(item) for item in row])
        wb.save(nome_excel)

        # --- GERADOR DE PDF EXECUTIVO (MÚLTIPLAS PÁGINAS) ---
        nome_pdf = "/tmp/Laudo_Auditoria.pdf"
        doc = SimpleDocTemplate(nome_pdf, pagesize=letter, rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=35)
        story = []
        styles = getSampleStyleSheet()
        
        st_tit = ParagraphStyle('Tit', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#0B132B'), spaceAfter=15)
        st_sub = ParagraphStyle('Sub', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor('#FF5A00'), spaceBefore=10, spaceAfter=8)
        st_txt = ParagraphStyle('Txt', parent=styles['Normal'], fontSize=9.5, textColor=colors.HexColor('#1E293B'), leading=14, spaceAfter=8)
        
        # PÁGINA 1: Cabeçalho, Resumo da IA e Gráfico Principal
        story.append(Paragraph("LAUDO DE AUDITORIA OPERACIONAL - BETNACIONAL", st_tit))
        story.append(Paragraph("Este documento apresenta a análise profunda da volumetria, comportamento do cliente e performance do time.", st_txt))
        story.append(Image(path_motivos, width=380, height=190))
        story.append(Spacer(1, 10))
        
        # Textos da IA formatados
        for paragrafo in texto_ia.split('\n'):
            texto = paragrafo.replace('**', '').strip()
            if not texto: continue
            if texto[0].isdigit() and '.' in texto[:3]: # Identifica os Títulos (1., 2., etc)
                story.append(Paragraph(texto, st_sub))
            else:
                story.append(Paragraph(texto, st_txt))
                
        story.append(PageBreak()) # Força a criação da Página 2 para não cortar nada
        
        # PÁGINA 2: Métricas Aprofundadas (Gráficos Lado a Lado)
        story.append(Paragraph("MÉTRICAS APROFUNDADAS DE ATRITO E PERFORMANCE", st_tit))
        
        tabela_graficos = Table([
            [Image(path_origem, width=220, height=165), Image(path_inativos, width=220, height=165)]
        ])
        story.append(tabela_graficos)
        story.append(Spacer(1, 10))
        story.append(Image(path_atendentes, width=350, height=210))
        
        doc.build(story)

        # Preparar retorno
        with open(nome_excel, "rb") as f: exc_64 = base64.b64encode(f.read()).decode('utf-8')
        with open(nome_pdf, "rb") as f: pdf_64 = base64.b64encode(f.read()).decode('utf-8')

        return jsonify({"status": "sucesso", "mensagem": "Auditoria Premium Finalizada!", "excel": exc_64, "pdf": pdf_64}), 200

    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
