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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

app = Flask(__name__)
CORS(app) 

CHAVE_API = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=CHAVE_API)

LOGO_PATH = "logo.png"
SLA_META_MINUTOS = 15 # Meta corporativa de tempo de resolução

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

def calcular_tma(conversa):
    if not isinstance(conversa, str): return 0
    tempos = re.findall(r'\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]', conversa)
    if len(tempos) >= 2:
        try:
            t_inicio = datetime.strptime(tempos[0], '%Y-%m-%d %H:%M:%S')
            t_fim = datetime.strptime(tempos[-1], '%Y-%m-%d %H:%M:%S')
            minutos = abs((t_fim - t_inicio).total_seconds() / 60.0)
            return minutos if minutos < 300 else 0
        except: return 0
    return 0

def adicionar_rodape(canvas, doc):
    canvas.saveState()
    canvas.setFont('Helvetica', 8)
    canvas.setFillColor(colors.HexColor('#64748B'))
    canvas.drawString(35, 20, "Betnacional | Relatório Executivo de Inteligência Operacional - STRICTLY CONFIDENTIAL")
    canvas.drawRightString(letter[0]-35, 20, f"Página {doc.page}")
    canvas.restoreState()

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
        df['tma_minutos'] = df['comments'].apply(calcular_tma)
        df['cumpriu_sla'] = df['tma_minutos'] <= SLA_META_MINUTOS
        
        for col in ['comments', 'ticket_summary']:
            if col in df.columns: df[col] = df[col].apply(limpar_dados_sensiveis)

        # KPIs Corporativos
        qtd_inativos = int(df['is_inativo'].sum())
        pct_inativos = (qtd_inativos / total_geral_casos) * 100
        
        origem_counts = df['origem_erro'].value_counts()
        pct_erro_cliente = (int(origem_counts.get('Erro/Dúvida do Cliente', 0)) / total_geral_casos) * 100
        pct_erro_sistema = 100 - pct_erro_cliente
        
        casos_validos_tma = df[df['tma_minutos'] > 0]
        tma_medio_global = casos_validos_tma['tma_minutos'].mean() if not casos_validos_tma.empty else 0
        pct_dentro_sla = (len(casos_validos_tma[casos_validos_tma['cumpriu_sla']]) / len(casos_validos_tma) * 100) if not casos_validos_tma.empty else 0

        top_atendentes = df[df['atendente_extraido'] != "Não identificado"]['atendente_extraido'].value_counts().head(5)
        
        motivos_agrupados = df.groupby('reason').agg(
            Quantidade=('reason', 'count'),
            TMA_Medio=('tma_minutos', lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 0)
        ).reset_index().sort_values(by='Quantidade', ascending=False)
        motivos_top5 = motivos_agrupados.head(5)
        top_motivo_nome = str(motivos_top5.iloc[0]['reason']).replace('_', ' ').title()
        top_motivo_qtd = motivos_top5.iloc[0]['Quantidade']

        # --- GERADOR DE GRÁFICOS PREMIUM ---
        cores_bet = ['#FF5A00', '#1C2541', '#3A506B', '#5BC0BE', '#CBD5E1']
        font_titulo_grafico = {'fontsize': 10, 'fontweight': 'bold', 'color': '#0B132B'}
        
        fig1, ax1 = plt.subplots(figsize=(6.5, 3))
        ax1.pie(motivos_top5['Quantidade'], colors=cores_bet, startangle=140, pctdistance=0.75, textprops=dict(color="white", weight="bold", fontsize=8))
        ax1.add_artist(plt.Circle((0,0), 0.55, fc='white'))
        ax1.legend([m.replace('_', ' ').title() for m in motivos_top5['reason']], loc="center left", bbox_to_anchor=(0.9, 0.5), frameon=False, fontsize=8)
        ax1.set_title('Concentração de Volume (Curva ABC)', fontdict=font_titulo_grafico, loc='left')
        plt.tight_layout()
        path_motivos = "/tmp/g_motivos.png"
        fig1.savefig(path_motivos, dpi=150)
        plt.close(fig1)

        fig2, ax2 = plt.subplots(figsize=(4, 3))
        ax2.pie([pct_dentro_sla, 100-pct_dentro_sla], labels=[f'No Prazo (<{SLA_META_MINUTOS}m)', 'Estourado'], autopct='%1.1f%%', colors=['#1C2541', '#FF5A00'], startangle=90, textprops=dict(color="white", weight="bold", fontsize=9))
        ax2.set_title('Aderência ao SLA', fontdict=font_titulo_grafico)
        plt.tight_layout()
        path_sla = "/tmp/g_sla.png"
        fig2.savefig(path_sla, dpi=150)
        plt.close(fig2)

        fig4, ax4 = plt.subplots(figsize=(6, 3))
        ax4.barh(top_atendentes.index[::-1], top_atendentes.values[::-1], color='#3A506B')
        ax4.set_title('Performance Individual (Top 5 Operadores)', fontdict=font_titulo_grafico, loc='left')
        ax4.spines['top'].set_visible(False)
        ax4.spines['right'].set_visible(False)
        plt.tight_layout()
        path_atendentes = "/tmp/g_atendentes.png"
        fig4.savefig(path_atendentes, dpi=150)
        plt.close(fig4)

        # --- IA CORPORATIVA (FRAMEWORK SCR - McKinsey Style) ---
        prompt = (f"Atue como um Sócio Diretor de Consultoria Estratégica analisando a operação da Betnacional.\n"
                  f"DADOS MACRO: {total_geral_casos} casos. TMA Global: {tma_medio_global:.1f} min. SLA de Sucesso: {pct_dentro_sla:.1f}%.\n"
                  f"GARGALO: '{top_motivo_nome}' causou {top_motivo_qtd} contatos.\n"
                  f"COMPORTAMENTO: {pct_inativos:.1f}% abandono. {pct_erro_cliente:.1f}% falha do cliente vs {pct_erro_sistema:.1f}% sistema.\n\n"
                  f"Escreva um Parecer Executivo rigoroso, sem jargões desnecessários, usando EXATAMENTE estes títulos:\n"
                  f"1. EXECUTIVE SUMMARY (Resumo de 3 linhas do impacto financeiro/operacional)\n"
                  f"2. SITUAÇÃO (O contexto atual de volumetria e SLA)\n"
                  f"3. COMPLICAÇÃO (Onde o gargalo '{top_motivo_nome}' e a inatividade ferem a empresa)\n"
                  f"4. RESOLUÇÃO ESTRATÉGICA (3 táticas imediatas de alto impacto em bullet points)")
        try:
            texto_ia = client.models.generate_content(model='gemini-2.5-flash', contents=prompt).text
        except:
            texto_ia = "Falha ao processar análise avançada corporativa da IA."

        # --- EXCEL NÍVEL ENTERPRISE ---
        nome_excel = "/tmp/Relatorio_Auditoria_Betnacional.xlsx"
        wb = openpyxl.Workbook()
        
        # ABA 1: DASHBOARD (Interface limpa como SaaS)
        ws1 = wb.active
        ws1.title = 'Dashboard Executivo'
        ws1.views.sheetView[0].showGridLines = False
        
        # Fundo cinza claro corporativo para todo o dashboard
        fill_fundo = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        for row in ws1.iter_rows(min_row=1, max_row=40, min_col=1, max_col=15):
            for cell in row: cell.fill = fill_fundo

        borda_kpi = Border(left=Side(style='thin', color="CBD5E1"), right=Side(style='thin', color="CBD5E1"), top=Side(style='thin', color="CBD5E1"), bottom=Side(style='thin', color="CBD5E1"))
        
        linha_kpi = 2
        if os.path.exists(LOGO_PATH):
            img = openpyxl.drawing.image.Image(LOGO_PATH)
            img.width, img.height = 160, 40
            ws1.add_image(img, 'B2')
            linha_kpi = 5
        else:
            ws1['B2'] = "DASHBOARD DE INTELIGÊNCIA OPERACIONAL"
            ws1['B2'].font = Font(size=16, bold=True, color="0B132B")
            linha_kpi = 4
        
        kpis = [
            ("VOLUMETRIA TOTAL", total_geral_casos, f'B{linha_kpi}', f'B{linha_kpi+1}'),
            ("OFENSOR PRINCIPAL", f"{top_motivo_nome}", f'C{linha_kpi}', f'C{linha_kpi+1}'),
            ("SUCESSO SLA (<15m)", f"{pct_dentro_sla:.1f}%", f'D{linha_kpi}', f'D{linha_kpi+1}'),
            ("TMA GLOBAL", f"{tma_medio_global:.1f} min", f'E{linha_kpi}', f'E{linha_kpi+1}'),
            ("ABANDONO (CHURN)", f"{pct_inativos:.1f}%", f'F{linha_kpi}', f'F{linha_kpi+1}')
        ]
        
        for titulo, valor, c_tit, c_val in kpis:
            ws1[c_tit] = titulo
            ws1[c_tit].font = Font(size=8, bold=True, color="64748B")
            ws1[c_val] = valor
            ws1[c_val].font = Font(size=14, bold=True, color="1C2541")
            ws1[c_val].fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") # Fundo branco no cartão
            ws1[c_val].border = borda_kpi

        for col, width in zip(['A','B','C','D','E','F','G'], [2, 18, 30, 18, 18, 18, 2]): ws1.column_dimensions[col].width = width

        ws1.add_image(openpyxl.drawing.image.Image(path_motivos), f'B{linha_kpi+3}')
        ws1.add_image(openpyxl.drawing.image.Image(path_sla), f'E{linha_kpi+3}')
        ws1.add_image(openpyxl.drawing.image.Image(path_atendentes), f'B{linha_kpi+19}')

        # ABA 2: BASE COM CONGELAMENTO DE PAINÉIS
        ws3 = wb.create_sheet('Database Consolidada')
        ws3.append(list(df.columns))
        for cell in ws3[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = PatternFill(start_color="1C2541", end_color="1C2541", fill_type="solid")
        
        for _, row in df.iterrows(): ws3.append([str(item) for item in row])
        
        ws3.auto_filter.ref = ws3.dimensions
        ws3.freeze_panes = "A2" # Congela o cabeçalho!
        for col in ws3.columns: ws3.column_dimensions[get_column_letter(col[0].column)].width = 18

        wb.save(nome_excel)

        # --- GERADOR DE PDF NÍVEL DIRETORIA ---
        nome_pdf = "/tmp/Laudo_Auditoria.pdf"
        doc = SimpleDocTemplate(nome_pdf, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
        story = []
        styles = getSampleStyleSheet()
        
        st_capa_tit = ParagraphStyle('CapaTit', parent=styles['Heading1'], fontSize=28, textColor=colors.HexColor('#0B132B'), alignment=1, spaceAfter=20, fontName="Helvetica-Bold")
        st_capa_sub = ParagraphStyle('CapaSub', parent=styles['Heading2'], fontSize=14, textColor=colors.HexColor('#FF5A00'), alignment=1)
        st_tit = ParagraphStyle('Tit', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#0B132B'), spaceAfter=15, fontName="Helvetica-Bold")
        st_sub = ParagraphStyle('Sub', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#FF5A00'), spaceBefore=15, spaceAfter=8, fontName="Helvetica-Bold")
        st_txt = ParagraphStyle('Txt', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#334155'), leading=15, spaceAfter=10)
        
        # Estilo para Caixa de Destaque
        st_destaque = ParagraphStyle('Dest', parent=styles['Normal'], fontSize=11, textColor=colors.HexColor('#0B132B'), leading=16, fontName="Helvetica-Oblique")
        
        # PÁGINA 1: CAPA
        story.append(Spacer(1, 100))
        if os.path.exists(LOGO_PATH): story.append(RLImage(LOGO_PATH, width=220, height=55))
        story.append(Spacer(1, 50))
        story.append(Paragraph("RELATÓRIO DE INTELIGÊNCIA OPERACIONAL", st_capa_tit))
        story.append(Paragraph("Auditoria Estratégica e Avaliação de SLA", st_capa_sub))
        story.append(Spacer(1, 150))
        story.append(Paragraph(f"Emitido em: {datetime.now().strftime('%d/%m/%Y')}", ParagraphStyle('C', alignment=1, textColor=colors.HexColor('#64748B'))))
        story.append(PageBreak()) 
        
        # PÁGINA 2: PARECER EXECUTIVO E CAIXA DE DESTAQUE
        story.append(Paragraph("1. PARECER EXECUTIVO ESTRATÉGICO", st_tit))
        
        # Construção da Caixa de Leitura Rápida (Highlight Box)
        caixa_texto = [
            [Paragraph("<b>LEITURA RÁPIDA (FAST TRACK):</b>", st_txt)],
            [Paragraph(f"A operação processou <b>{total_geral_casos} casos</b> com um Tempo Médio (TMA) de <b>{tma_medio_global:.1f} minutos</b>. "
                       f"Atualmente, <b>{pct_dentro_sla:.1f}%</b> dos atendimentos cumprem a meta de SLA (<15 min). O ofensor principal é '<b>{top_motivo_nome}</b>', "
                       f"e a taxa de abandono do cliente fixa-se em <b>{pct_inativos:.1f}%</b>.", st_destaque)]
        ]
        tabela_destaque = Table(caixa_texto, colWidths=[460])
        tabela_destaque.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
            ('LEFTMARGIN', (0,0), (-1,-1), 15),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(tabela_destaque)
        story.append(Spacer(1, 20))
        
        # Texto da IA Estruturado
        for paragrafo in texto_ia.split('\n'):
            texto = paragrafo.replace('**', '').strip()
            if not texto: continue
            if texto[0].isdigit() and '.' in texto[:3]: 
                story.append(Paragraph(texto, st_sub))
            else:
                story.append(Paragraph(texto, st_txt))
                
        story.append(PageBreak()) 
        
        # PÁGINA 3: METRIFICAÇÃO VISUAL
        story.append(Paragraph("2. PAINEL DE INTELIGÊNCIA VISUAL", st_tit))
        story.append(RLImage(path_motivos, width=440, height=200))
        story.append(Spacer(1, 15))
        
        tabela_graficos = Table([[RLImage(path_sla, width=230, height=170), RLImage(path_atendentes, width=230, height=170)]])
        story.append(tabela_graficos)
        
        doc.build(story, onFirstPage=adicionar_rodape, onLaterPages=adicionar_rodape)

        with open(nome_excel, "rb") as f: exc_64 = base64.b64encode(f.read()).decode('utf-8')
        with open(nome_pdf, "rb") as f: pdf_64 = base64.b64encode(f.read()).decode('utf-8')

        return jsonify({"status": "sucesso", "mensagem": "Auditoria Corporativa C-Level gerada com sucesso!", "excel": exc_64, "pdf": pdf_64}), 200

    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
