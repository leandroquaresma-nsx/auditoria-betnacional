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
        
        # Gráfico Motivos
        fig, ax = plt.subplots(figsize=(6, 3))
        nomes_limpos = [m.replace('_', ' ').title() for m in motivos_top5['Motivo_Ocorrencia']]
        ax.barh(nomes_limpos, motivos_top5['Quantidade'], color='#FF5A00')
        ax.invert_yaxis()
        ax.set_title('TOP OCORRÊNCIAS NO SUPORTE', fontsize=11, fontweight='bold', color='#0B132B')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
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

        # Geração do Excel estruturado
        nome_excel = "/tmp/Relatorio_Auditoria_Betnacional.xlsx"
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = 'Dashboard Volumetrico'
        ws1.append(['Motivo da Ocorrência', 'Quantidade de Casos'])
        for _, row in principais_motivos.iterrows(): ws1.append(list(row))
        if os.path.exists(grafico_path):
            ws1.add_image(openpyxl.drawing.image.Image(grafico_path), 'D2')
        
        ws3 = wb.create_sheet(title='Base Dados Anonimizada')
        ws3.append(list(df.columns))
        for _, row in df.iterrows():
            ws3.append([str(item) if isinstance(item, (list, dict)) else item for item in row])
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
            story.append(Image(grafico_path, width=330, height=165))
            story.append(Spacer(1, 10))
            
        for p in texto_ia.split('\n\n'):
            story.append(Paragraph(p.replace('**', ''), style_corpo))
        doc.build(story)

        # Transformar arquivos em base64 para devolver ao navegador
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
