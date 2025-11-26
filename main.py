import os
import datetime
import locale
import smtplib
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas

# --- IMPORTANTE: Certifique-se que o arquivo pulsar.py está dentro da pasta 'modules' 
# ou ajuste o import abaixo caso esteja na mesma pasta raiz ---
from modules import graficos, reme, pulsar, sites, providers

load_dotenv()

# --- CONFIGURAÇÕES DO RELATÓRIO ---
SUPERVISOR_CARGO = "Supervisor Técnico"
HEADER_TEXT = "RELATÓRIO DIÁRIO - SUPERVISOR TÉCNICO/4° CTA"

# --- CONFIGURAÇÕES DE EMAIL ---
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")
SENHA_APP_GMAIL = os.getenv("SENHA_APP_GMAIL")
EMAIL_DESTINATARIO = os.getenv("EMAIL_DESTINATARIO")

# --- DICIONÁRIO DE MAPEAMENTO (KIT ID -> NOME DA OM) ---
MAPEAMENTO_OM = {
    "KIT304062259": "Cmdo 1ª Bda Inf Sl",
    "KIT304059560": "Cmdo 2ª Bda Inf Sl",
    "KIT304135659": "2ª Bda Inf Sl",
    "KITP00237489": "Cmdo 16ª Bda Inf Sl",
    "KIT304132110": "Cmdo 17ª Bda Inf Sl",
    "KIT304059859": "4º BIS - DEF - Epitaciolândia",
    "KIT304039763": "4º BIS - 2º PEF - Assis Brasil",
    "KIT304131574": "4º BIS - 3º PEF - Plácido de Castro",
    "KIT304039768": "4º BIS - 4º PEF - Santa Rosa do Purus",
    "KIT304132336": "5º BIS – 2º PEF - Querari",
    "KIT304039241": "5º BIS – 3º PEF - São Joaquim",
    "KIT303910747": "5º BIS – 4º PEF - Cucuí",
    "KIT304039236": "5º BIS – 5º PEF - Maturacá",
    "KIT304039230": "5º BIS – 6º PEF - Pari-Cachoeira",
    "KIT304039765": "5º BIS – 7º PEF - Tunuí",
    "KIT304135657": "7º BIS - 1º PEF - Bonfim",
    "KIT304059878": "7º BIS - 3º PEF - Pacaraima",
    "KIT304039242": "7º BIS - 4º PEF - Surucucu",
    "KIT304039235": "7º BIS - 5º PEF - Auaris",
    "KIT304044880": "7º BIS - 6º PEF - Uiramutã",
    "KIT304059852": "7º BIS - Base Pakilapi",
    "KIT304059547": "7º BIS - Base Kaianaú",
    "KIT303901850": "7º BIS - DEF Waikas",
    "KIT304059879": "8º BIS - 2º PEF - Ipiranga",
    "KIT304104044": "8º BIS - 4º PEF - Estirão do Equador",
    "KIT304039752": "61º BIS - DEF- Marechal Thaumaturgo",
    "KIT304132549": "34º BIS - Oiapoque",
    "KIT303903287": "34º BIS - Vila Brasil",
    "KIT304131555": "34º BIS - Tiriós",
    "KIT304039747": "3º BIS",
    "KIT304132264": "6º BIS - 1º PEF - Príncipe da Beira",
    "KIT303844328": "17º BIS",
    "KIT304132552": "17º BIS – 3º PEF-Vila Bittencourt",
    "KIT304039751": "HGuT",
    "KIT304059544": "2º B Log Sl",
    "KIT304039748": "21ª Cia E Cnst",
    "KIT304132127": "7º BEC (Destacamento)",
    "KIT304145670": "BI-02(CIGS)",
    "KIT303729090": "CMDO 8º BIS - Tabatinga",
    "KIT304132551": "4º CTA 02 - Manaus",
    "KIT304145658": "Cmdo 6º BIS 02",
    "KIT304132540": "2º PEF - Normandia",
    "KIT304145662": "4º CTA 01 - Manaus",
    "KIT304059853": "1º PEF Yauaretê"
}

# --- CLASSE PERSONALIZADA PARA SUMÁRIO AUTOMÁTICO ---
class RelatorioDocTemplate(SimpleDocTemplate):
    def afterFlowable(self, flowable):
        """Monitora o fluxo e avisa o Sumário quando encontra um título"""
        if flowable.__class__.__name__ == 'Paragraph':
            text = flowable.getPlainText()
            style_name = flowable.style.name
            
            # Captura Títulos Principais (Estilo H2_Custom)
            if style_name == 'H2_Custom':
                self.notify('TOCEntry', (0, text, self.page))
            
            # Captura Subtítulos (Estilo SubSection)
            elif style_name == 'SubSection':
                self.notify('TOCEntry', (1, text, self.page))

def get_data_por_extenso():
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        data_str = datetime.datetime.now().strftime('Manaus, %d de %B de %Y')
    except:
        dt = datetime.datetime.now()
        meses = {
            1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril',
            5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto',
            9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
        }
        data_str = f"Manaus, {dt.day} de {meses[dt.month]} de {dt.year}"
    return data_str

def enviar_email_com_anexo(arquivo_pdf):
    print(f">> Preparando envio de email para: {EMAIL_DESTINATARIO}...")
    
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = EMAIL_DESTINATARIO
        msg['Subject'] = f"Relatório Técnico Diário - {datetime.datetime.now().strftime('%d/%m/%Y')}"

        corpo = f"""
        Prezado Supervisor Técnico,

        Segue em anexo o Relatório Técnico Integrado gerado automaticamente.
        
        Data: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}
        Origem: Servidor de Monitoramento (VM)
        """
        msg.attach(MIMEText(corpo, 'plain'))

        # Anexar PDF
        with open(arquivo_pdf, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {os.path.basename(arquivo_pdf)}",
        )
        msg.attach(part)

        # Conectar ao Gmail (SMTP)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_APP_GMAIL)
        text = msg.as_string()
        server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIO, text)
        server.quit()
        
        print("✅ Email enviado com sucesso!")
        return True

    except Exception as e:
        print(f"❌ Falha ao enviar email: {e}")
        return False

def header_footer_template(canvas, doc):
    canvas.saveState()
    w, h = A4
    
    # Cabeçalho
    canvas.setStrokeColor(colors.black)
    canvas.setFillColor(colors.lightgrey)
    canvas.roundRect(2*cm, h - 2.5*cm, w - 4*cm, 1.2*cm, 5, fill=1, stroke=1)
    
    canvas.setFillColor(colors.black)
    canvas.setFont("Helvetica-Bold", 12)
    canvas.drawCentredString(w/2, h - 1.9*cm, HEADER_TEXT)
    
    # Rodapé
    page_num = canvas.getPageNumber()
    canvas.setFont("Helvetica", 10)
    canvas.drawRightString(w - 2*cm, 1.5*cm, f"pág. {page_num}")
    
    canvas.restoreState()

def generate_unified_report():
    folder_out = "output"
    if not os.path.exists(folder_out): os.makedirs(folder_out)
    
    data_filename = datetime.datetime.now().strftime('%Y-%m-%d')
    filename = f"{folder_out}/Relatorio_Integrado_{data_filename}.pdf"
    
    print(f"\n=== GERAÇÃO DE RELATÓRIO: {data_filename} ===\n")

    # --- 1. PRÉ-COLETA DE DADOS ---
    print(">> Coletando dados (Providers, Gráficos, REME, Pulsar, Sites)...")
    
    res_prov, res_tun = providers.collect_providers_data()
    dados_graficos = graficos.collect_graph_images()
    dados_reme = reme.collect_reme_data()
    
    # COLETA DO PULSAR (STARLINK)
    try:
        # Chama a função que você definiu no pulsar3.py
        # headless=True para rodar em servidor
        dados_pulsar = pulsar.extrair_dados_starlink(headless=True)
    except Exception as e:
        print(f"❌ Erro ao extrair dados do Pulsar: {e}")
        dados_pulsar = []

    dados_sites, ocorrencias_sites = sites.collect_sites_data()

    # --- 2. CONSTRUÇÃO DO DOCUMENTO ---
    doc = RelatorioDocTemplate(filename, pagesize=A4, 
                            topMargin=3.5*cm, bottomMargin=2.5*cm, 
                            rightMargin=2*cm, leftMargin=2*cm)
    story = []
    styles = getSampleStyleSheet()

    # Estilos
    style_center = ParagraphStyle('Center', parent=styles['Normal'], alignment=TA_CENTER, fontSize=12)
    style_bold_center = ParagraphStyle('BoldCenter', parent=styles['Normal'], alignment=TA_CENTER, fontSize=12, fontName='Helvetica-Bold')
    
    # Estilos de Título para o Sumário
    style_h2 = ParagraphStyle('H2_Custom', parent=styles['Heading2'], fontSize=12, spaceBefore=15, spaceAfter=10, fontName='Helvetica-Bold', textTransform='uppercase')
    style_sub = ParagraphStyle('SubSection', parent=styles['Normal'], fontSize=11, spaceBefore=10, spaceAfter=5, fontName='Helvetica-Bold')
    
    style_cell_center = ParagraphStyle('CellC', fontSize=9, alignment=TA_CENTER)
    style_cell_small = ParagraphStyle('CellS', fontSize=8, alignment=TA_CENTER)

    # CAPA
    story.append(Spacer(1, 1*cm))
    story.append(Paragraph(get_data_por_extenso(), style_center))
    story.append(Spacer(1, 1.5*cm))
    story.append(Paragraph(SUPERVISOR_CARGO, style_center))
    story.append(Spacer(1, 2*cm))
    
    # --- SUMÁRIO AUTOMÁTICO ---
    story.append(Paragraph("<u>SUMÁRIO</u>", ParagraphStyle('SumarioTitle', parent=style_bold_center, fontSize=14)))
    story.append(Spacer(1, 1*cm))
    
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle(fontName='Helvetica-Bold', fontSize=10, name='TOCHeading1', leftIndent=20, firstLineIndent=-20, spaceBefore=5, leading=12),
        ParagraphStyle(fontName='Helvetica', fontSize=10, name='TOCHeading2', leftIndent=40, firstLineIndent=-20, spaceBefore=0, leading=12),
    ]
    toc.dots = '.' 
    story.append(toc)
    story.append(PageBreak())

    sec_num = 1

    # ==========================================================
    # SEÇÃO 1: PROVEDORES
    # ==========================================================
    story.append(Paragraph(f"{sec_num}. STATUS DE PROVEDORES DE INTERNET E BBI", style_h2))
    
    if res_prov:
        data_p = [['N', 'LINK', 'TESTE / OBS', 'LINK WAN', 'STATUS']]
        colors_p = []
        for i, p in enumerate(res_prov, 1):
            texto_teste = p['teste_str']
            if p['observacao']: texto_teste += f"\n{p['observacao']}"
            bg = colors.lime if p['status'] == "GREEN" else colors.red
            colors_p.append(bg)
            
            p_link = Paragraph(p['link'], style_cell_center)
            p_teste = Paragraph(texto_teste, style_cell_small)
            data_p.append([str(i), p_link, p_teste, p['link_wan'], ""])

        t_prov = Table(data_p, colWidths=[1*cm, 4.5*cm, 6.5*cm, 3*cm, 2*cm], repeatRows=1)
        sty_prov = [
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]
        for idx, c in enumerate(colors_p):
            sty_prov.append(('BACKGROUND', (4, idx+1), (4, idx+1), c))
            
        t_prov.setStyle(TableStyle(sty_prov))
        story.append(t_prov)
        story.append(Spacer(1, 0.2*cm))
        story.append(Paragraph("*ping / traceroute da internet para (atestar a rota ao AS via operadora).<br/>**traceroute da EBNET.", ParagraphStyle('Note', fontSize=8)))

    if res_tun:
        story.append(Spacer(1, 0.5*cm))
        story.append(Paragraph("Status dos túneis configurados para contingência (REME MAO -> 4º CTA)", style_sub))
        
        data_t = [['TÚNEL', 'TESTE / OBS', 'STATUS']]
        colors_t = []
        for t in res_tun:
            texto_teste = t['teste_str']
            if t['observacao']: texto_teste += f"\n{t['observacao']}"
            bg = colors.lime if t['status'] == "GREEN" else colors.red
            colors_t.append(bg)
            
            p_tun = Paragraph(t['tunel'], style_cell_center)
            p_teste = Paragraph(texto_teste, style_cell_small)
            data_t.append([p_tun, p_teste, ""])

        t_tun = Table(data_t, colWidths=[6*cm, 9*cm, 2*cm], repeatRows=1)
        sty_tun = [
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]
        for idx, c in enumerate(colors_t):
            sty_tun.append(('BACKGROUND', (2, idx+1), (2, idx+1), c))
        t_tun.setStyle(TableStyle(sty_tun))
        story.append(t_tun)

    story.append(PageBreak())
    sec_num += 1

    # ==========================================================
    # SEÇÃO 2: GRÁFICOS
    # ==========================================================
    story.append(Paragraph(f"{sec_num}. GRÁFICOS DE BANDA E LATÊNCIA PELOS LINKS DE ENTRADA", style_h2))
    
    if dados_graficos:
        for i, (titulo, img_bytes) in enumerate(dados_graficos, 1):
            story.append(Paragraph(f"{sec_num}.{i} {titulo}", style_sub))
            story.append(Spacer(1, 0.2*cm))
            img = Image(img_bytes, width=16*cm, height=4.2*cm)
            story.append(img)
            story.append(Spacer(1, 0.5*cm))
    else:
        story.append(Paragraph("Gráficos indisponíveis.", styles['Normal']))

    story.append(PageBreak())
    sec_num += 1

    # ==========================================================
    # SEÇÃO 3: REME
    # ==========================================================
    if dados_reme:
        for secao in dados_reme:
            story.append(Paragraph(f"{sec_num}. {secao['titulo']}", style_h2))
            
            if not secao['dados']:
                story.append(Paragraph("Sem dados para esta localidade.", styles['Normal']))
            else:
                t_data = [['N', 'OM', 'PoP (IP)', 'Status']]
                r_colors = []
                for i, dado in enumerate(secao['dados'], start=1):
                    cor_nome = dado['cor']
                    c = colors.lightgrey
                    if cor_nome == "RED": c = colors.red
                    elif cor_nome == "YELLOW": c = colors.yellow
                    elif cor_nome == "GREEN": c = colors.lime
                    
                    p_om = Paragraph(dado['om'], ParagraphStyle('Cell', fontSize=8, alignment=TA_CENTER))
                    t_data.append([str(i), p_om, dado['ip'], dado['status']])
                    r_colors.append(c)
                    
                t = Table(t_data, colWidths=[1*cm, 7*cm, 6*cm, 3*cm], repeatRows=1)
                sty = [
                    ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                    ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,-1), 8)
                ]
                for idx, color_bg in enumerate(r_colors):
                    sty.append(('BACKGROUND', (3, idx+1), (3, idx+1), color_bg))
                t.setStyle(TableStyle(sty))
                story.append(t)
            
            story.append(Spacer(1, 0.8*cm))
            sec_num += 1
    else:
        story.append(Paragraph(f"{sec_num}. STATUS DOS POP REME", style_h2))
        story.append(Paragraph("Não foi possível coletar dados da REME.", styles['Normal']))
        sec_num += 1

    story.append(PageBreak())

    # ==========================================================
    # SEÇÃO 4: PULSAR (STARLINK)
    # ==========================================================
    story.append(Paragraph(f"{sec_num}. STATUS DOS PONTOS SATELITAIS - STARLINK", style_h2))
    
    if dados_pulsar and len(dados_pulsar) > 0:
        # Cabeçalho da Tabela
        header = [['N', 'OM', 'PoP', 'STATUS']]
        t_data = header
        p_colors = []
        
        # Filtra registros que tenham pelo menos OM ou PoP preenchidos
        valid_pulsar = [p for p in dados_pulsar if p.get('om') or p.get('pop')]
        
        # Estilo específico para a célula da OM (Alinhada à esquerda)
        style_cell_om = ParagraphStyle('CellOM', parent=styles['Normal'], fontSize=8, alignment=TA_LEFT, leading=9)
        
        for i, p in enumerate(valid_pulsar, 1):
            # 1. Recupera o Kit ID (PoP) e o nome original vindo do extrator
            kit_id = p.get('pop', 'N/A').strip() 
            om_original = p.get('om', 'Desconhecido').strip()
            
            # 2. LOGICA DE MAPEAMENTO:
            # Tenta encontrar o Kit ID no dicionário MAPEAMENTO_OM.
            # Se encontrar, usa o nome do dicionário. Se não, usa o om_original vindo do site.
            om_final = MAPEAMENTO_OM.get(kit_id, om_original)
            
            # 3. Lógica de Status/Cor
            status_raw = p.get('status', 'DESCONHECIDO')
            c = colors.lightgrey # Cor padrão se for desconhecido
            
            if status_raw == "VERDE":
                c = colors.lime # Verde para online
            elif status_raw == "VERMELHO":
                c = colors.red  # Vermelho para offline
            elif status_raw == "AMARELO":
                c = colors.yellow
            
            # Monta os dados da linha
            p_n = str(i)
            p_om = Paragraph(om_final, style_cell_om) # Nome formatado
            p_pop = kit_id if kit_id != "N/A" else "-" # Mostra o Kit ID
            
            # Adiciona a linha na matriz da tabela
            # A última coluna "" é vazia de texto pois será pintada
            t_data.append([p_n, p_om, p_pop, ""])
            p_colors.append(c)
            
        # Verifica se há dados para gerar a tabela
        if len(t_data) > 1:
            # Definição das larguras (Total ~17cm)
            col_widths = [1*cm, 10*cm, 4*cm, 2*cm]
            
            t = Table(t_data, colWidths=col_widths, repeatRows=1)
            
            sty = [
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), # Fundo Cinza no Cabeçalho
                ('ALIGN', (0,0), (-1,-1), 'CENTER'), # Alinhamento Geral Centro
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), # Negrito no Cabeçalho
                ('FONTSIZE', (0,0), (-1,-1), 8),
                # Alinhamento específico da Coluna OM (índice 1) para Esquerda
                ('ALIGN', (1,1), (1,-1), 'LEFT'), 
            ]
            
            # Aplica as cores na coluna STATUS (índice 3) linha por linha
            for idx, color_bg in enumerate(p_colors):
                sty.append(('BACKGROUND', (3, idx+1), (3, idx+1), color_bg))
            
            t.setStyle(TableStyle(sty))
            story.append(t)
            
            # Resumo Estatístico no rodapé da seção
            online_count = len([x for x in valid_pulsar if x.get('status') == "VERDE"])
            offline_count = len([x for x in valid_pulsar if x.get('status') == "VERMELHO"])
            total_count = len(valid_pulsar)
            
            story.append(Spacer(1, 0.3*cm))
            texto_resumo = f"<b>Resumo Starlink:</b> Total: {total_count} | <font color='green'>Online: {online_count}</font> | <font color='red'>Offline: {offline_count}</font>"
            story.append(Paragraph(texto_resumo, styles['Normal']))
            
    else:
        story.append(Paragraph("Sem dados do Pulsar disponíveis no momento.", styles['Normal']))

    story.append(PageBreak())
    sec_num += 1

    # ==========================================================
    # SEÇÃO 5: SITES
    # ==========================================================
    story.append(Paragraph(f"{sec_num}. FUNCIONAMENTO DOS SITES HOSPEDADOS", style_h2))
    
    if dados_sites:
        header = [['Ord', 'OM', 'Endereço', 'Status', 'Ocorrência']]
        t_data = header
        s_colors = []
        style_url = ParagraphStyle('URL', fontSize=7, alignment=TA_CENTER, wordWrap='CJK')
        style_cell = ParagraphStyle('CellS', fontSize=8, alignment=TA_CENTER)

        for row in dados_sites:
            bg = colors.lime if row[5] == "GREEN" else colors.red
            p_om = Paragraph(row[1], style_cell)
            p_url = Paragraph(row[2], style_url)
            t_data.append([row[0], p_om, p_url, row[3], row[4]])
            s_colors.append(bg)
            
        t = Table(t_data, colWidths=[1*cm, 3*cm, 9*cm, 2*cm, 2*cm], repeatRows=1)
        sty = [
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 8)
        ]
        for idx, c in enumerate(s_colors):
            sty.append(('BACKGROUND', (3, idx+1), (3, idx+1), c))
        t.setStyle(TableStyle(sty))
        story.append(t)
    else:
        story.append(Paragraph("Sem dados de sites.", styles['Normal']))
        
    if ocorrencias_sites:
        story.append(Spacer(1, 0.5*cm))
        story.append(Paragraph(f"{sec_num}.1 DETALHAMENTO DE OCORRÊNCIAS", style_h2))
        data_oc = [['ID', 'AÇÃO DE MITIGAÇÃO / OBSERVAÇÃO', 'GDH RESOLUÇÃO']] + ocorrencias_sites
        t_oc = Table(data_oc, colWidths=[1.5*cm, 12.5*cm, 3*cm])
        t_oc.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('ALIGN', (1,1), (1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 9)
        ]))
        story.append(t_oc)

    # GERAÇÃO FINAL
    try:
        doc.multiBuild(story, onFirstPage=header_footer_template, onLaterPages=header_footer_template)
        print(f"\n✅ Relatório gerado com sucesso: {filename}")
        
        # --- ENVIO DE EMAIL ---
        enviar_email_com_anexo(filename)
        
    except Exception as e:
        print(f"\n❌ Erro crítico: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    generate_unified_report()