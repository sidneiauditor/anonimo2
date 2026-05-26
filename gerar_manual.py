import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, Image
)
from reportlab.platypus.flowables import KeepTogether

OUTPUT = "Manual_Anonimo_v1.0.pdf"
LOGO = "logo_anonimo.png"
W, H = A4

# ── Cores ────────────────────────────────────────────────────────────────────
AZUL_ESCURO  = colors.HexColor("#0F172A")
AZUL         = colors.HexColor("#2563EB")
AZUL_CLARO   = colors.HexColor("#EFF6FF")
VERDE        = colors.HexColor("#10B981")
CINZA        = colors.HexColor("#64748B")
CINZA_CLARO  = colors.HexColor("#F8FAFC")
VERMELHO     = colors.HexColor("#EF4444")
BRANCO       = colors.white

# ── Estilos ──────────────────────────────────────────────────────────────────
styles = getSampleStyleSheet()

def estilo(nome, parent='Normal', **kw):
    s = ParagraphStyle(nome, parent=styles[parent], **kw)
    return s

titulo_capa   = estilo('TituloCapa',   fontSize=32, textColor=BRANCO,
                        alignment=TA_CENTER, fontName='Helvetica-Bold', leading=40)
sub_capa      = estilo('SubCapa',      fontSize=14, textColor=colors.HexColor("#94A3B8"),
                        alignment=TA_CENTER, fontName='Helvetica')
versao_capa   = estilo('VersaoCapa',   fontSize=11, textColor=colors.HexColor("#60A5FA"),
                        alignment=TA_CENTER, fontName='Helvetica')
titulo_sec    = estilo('TituloSec',    fontSize=16, textColor=AZUL_ESCURO,
                        fontName='Helvetica-Bold', spaceBefore=8, spaceAfter=3)
titulo_sub    = estilo('TituloSub',    fontSize=12, textColor=AZUL,
                        fontName='Helvetica-Bold', spaceBefore=6, spaceAfter=2)
corpo         = estilo('Corpo',        fontSize=10, textColor=colors.HexColor("#1E293B"),
                        leading=14, alignment=TA_JUSTIFY, spaceAfter=4)
nota          = estilo('Nota',         fontSize=9,  textColor=CINZA,
                        leading=14, fontName='Helvetica-Oblique', spaceAfter=4)
passo_num     = estilo('PassoNum',     fontSize=10, textColor=BRANCO,
                        fontName='Helvetica-Bold', alignment=TA_CENTER)
passo_texto   = estilo('PassoTexto',   fontSize=10, textColor=AZUL_ESCURO,
                        leading=15, fontName='Helvetica-Bold')
passo_desc    = estilo('PassoDesc',    fontSize=9,  textColor=CINZA,
                        leading=14)
rodape_txt    = estilo('Rodape',       fontSize=8,  textColor=CINZA,
                        alignment=TA_CENTER)
aviso_txt     = estilo('Aviso',        fontSize=9,  textColor=colors.HexColor("#92400E"),
                        leading=14)

# ── Helpers ──────────────────────────────────────────────────────────────────
def hr(cor=colors.HexColor("#E2E8F0"), esp=3):
    return [Spacer(1, esp), HRFlowable(width="100%", thickness=1, color=cor), Spacer(1, esp)]

def caixa_aviso(texto, cor_borda=colors.HexColor("#FCD34D"), cor_fundo=colors.HexColor("#FFFBEB")):
    t = Table([[Paragraph(texto, aviso_txt)]], colWidths=[14*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), cor_fundo),
        ('BOX',        (0,0), (-1,-1), 1, cor_borda),
        ('ROUNDEDCORNERS', [6]),
        ('LEFTPADDING',  (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
        ('TOPPADDING',   (0,0), (-1,-1), 8),
        ('BOTTOMPADDING',(0,0), (-1,-1), 8),
    ]))
    return t

def caixa_info(texto):
    return caixa_aviso(texto, cor_borda=AZUL, cor_fundo=AZUL_CLARO)

def tabela_passos(passos):
    rows = []
    for num, titulo, desc in passos:
        num_cell  = Table([[Paragraph(str(num), passo_num)]],
                          colWidths=[1*cm], rowHeights=[1*cm])
        num_cell.setStyle(TableStyle([
            ('BACKGROUND',   (0,0), (-1,-1), AZUL),
            ('ROUNDEDCORNERS', [4]),
            ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ]))
        texto_cell = [Paragraph(titulo, passo_texto), Paragraph(desc, passo_desc)]
        rows.append([num_cell, texto_cell])

    t = Table(rows, colWidths=[1.4*cm, 13.6*cm], rowHeights=None)
    t.setStyle(TableStyle([
        ('VALIGN',        (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING',   (1,0), (1,-1), 10),
        ('TOPPADDING',    (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [CINZA_CLARO, BRANCO]),
        ('LINEBELOW',     (0,0), (-1,-2), 0.5, colors.HexColor("#E2E8F0")),
        ('BOX',           (0,0), (-1,-1), 1,   colors.HexColor("#E2E8F0")),
    ]))
    return t

# ── Numeração de páginas ──────────────────────────────────────────────────────
def rodape(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(CINZA)
    canvas.setFont('Helvetica', 8)
    canvas.drawCentredString(W/2, 1.2*cm,
        f"Anônimo v1.0  ·  Manual do Utilizador  ·  Página {doc.page}")
    canvas.setStrokeColor(colors.HexColor("#E2E8F0"))
    canvas.setLineWidth(0.5)
    canvas.line(2*cm, 1.5*cm, W - 2*cm, 1.5*cm)
    canvas.restoreState()

def sem_rodape(canvas, doc):
    pass

# ── Conteúdo ──────────────────────────────────────────────────────────────────
story = []

# ── CAPA ─────────────────────────────────────────────────────────────────────
fundo = Table([['']],  colWidths=[W - 4*cm], rowHeights=[H - 4*cm])
fundo.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), AZUL_ESCURO)]))

capa_content = []
capa_content.append(Spacer(1, 3*cm))

if os.path.exists(LOGO):
    img = Image(LOGO, width=4*cm, height=4*cm, kind='proportional')
    img.hAlign = 'CENTER'
    capa_content.append(img)
    capa_content.append(Spacer(1, 0.8*cm))

capa_content.append(Paragraph("Anônimo", titulo_capa))
capa_content.append(Spacer(1, 0.1*cm))
capa_content.append(Paragraph("Anonimização de Dados Sensíveis", sub_capa))
capa_content.append(Spacer(1, 0.25*cm))
capa_content.append(Paragraph("v1.0  ·  Manual do Utilizador", versao_capa))
capa_content.append(Spacer(1, 2*cm))

linha_capa = Table([['']],
    colWidths=[8*cm], rowHeights=[0.05*cm])
linha_capa.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#3B82F6")),
]))
linha_capa.hAlign = 'CENTER'
capa_content.append(linha_capa)
capa_content.append(Spacer(1, 1.5*cm))

info_rows = [
    ["Coordenadoria de Inteligência Fiscal"],
    ["Processamento 100% local · Sem envio de dados"],
]
for row in info_rows:
    capa_content.append(Paragraph(row[0], estilo('InfoCapa', fontSize=10,
        textColor=colors.HexColor("#94A3B8"), alignment=TA_CENTER)))
    capa_content.append(Spacer(1, 0.2*cm))

capa_table = Table([[ capa_content ]],
    colWidths=[W - 4*cm], rowHeights=[H - 6*cm])
capa_table.setStyle(TableStyle([
    ('BACKGROUND',   (0,0), (-1,-1), AZUL_ESCURO),
    ('VALIGN',       (0,0), (-1,-1), 'TOP'),
    ('LEFTPADDING',  (0,0), (-1,-1), 40),
    ('RIGHTPADDING', (0,0), (-1,-1), 40),
    ('TOPPADDING',   (0,0), (-1,-1), 0),
    ('BOTTOMPADDING',(0,0), (-1,-1), 0),
]))
story.append(capa_table)
story.append(PageBreak())

# ── ÍNDICE ────────────────────────────────────────────────────────────────────
story.append(Paragraph("Índice", titulo_sec))
*hr(),
story += hr()

indice = [
    ("1.", "Introdução", "3"),
    ("2.", "Requisitos do Sistema", "3"),
    ("3.", "Como Iniciar o Sistema", "4"),
    ("4.", "Guia de Uso — Passo a Passo", "4"),
    ("    4.1", "Carregar Arquivo", "5"),
    ("    4.2", "Analisar Estrutura", "5"),
    ("    4.3", "Selecionar Campos", "6"),
    ("    4.4", "Anonimizar Dados", "6"),
    ("    4.5", "Reverter Tokens", "7"),
    ("    4.6", "Processar PDF / OCR", "7"),
    ("5.", "Segurança e Privacidade", "6"),
]

idx_rows = [[Paragraph(n, corpo), Paragraph(t, corpo), Paragraph(p, corpo)]
            for n, t, p in indice]
idx_table = Table(idx_rows, colWidths=[1.5*cm, 12*cm, 1.5*cm])
idx_table.setStyle(TableStyle([
    ('ALIGN',        (2,0), (2,-1), 'RIGHT'),
    ('FONTNAME',     (2,0), (2,-1), 'Helvetica-Bold'),
    ('TEXTCOLOR',    (2,0), (2,-1), AZUL),
    ('LINEBELOW',    (0,0), (-1,-2), 0.3, colors.HexColor("#E2E8F0")),
    ('TOPPADDING',   (0,0), (-1,-1), 4),
    ('BOTTOMPADDING',(0,0), (-1,-1), 4),
]))
story.append(idx_table)
story.append(PageBreak())

# ── 1. INTRODUÇÃO ─────────────────────────────────────────────────────────────
story.append(Paragraph("1. Introdução", titulo_sec))
story += hr()
story.append(Paragraph(
    "O <b>Anônimo v1.0</b> é uma ferramenta de anonimização de dados sensíveis "
    "desenvolvida para uso interno da <b>Coordenadoria de Inteligência Fiscal</b>. "
    "Permite substituir informações confidenciais — como CNPJ, CPF, razão social e "
    "endereços — por tokens criptográficos determinísticos, garantindo a proteção "
    "dos dados sem perda de rastreabilidade.", corpo))
story.append(Paragraph(
    "Todo o processamento ocorre <b>exclusivamente no computador local</b>. "
    "Nenhum dado é transmitido pela internet, tornando a ferramenta adequada "
    "para ambientes com restrições de segurança.", corpo))

story.append(Spacer(1, 0.1*cm))
story.append(caixa_info(
    "<b>Formatos suportados:</b> .xlsx (Excel), .csv, .docx (Word), .txt e .pdf (com OCR)"))
story.append(Spacer(1, 0.2*cm))

# ── 2. REQUISITOS ─────────────────────────────────────────────────────────────
story.append(Paragraph("2. Requisitos do Sistema", titulo_sec))
story += hr()

req = [
    ["Componente", "Requisito"],
    ["Sistema Operacional", "Windows 10 / 11"],
    ["Python", "Versão 3.8 ou superior (para servidor local)"],
    ["Navegador", "Google Chrome, Edge ou Firefox (versão atual)"],
    ["Memória RAM", "Mínimo 4 GB (recomendado 8 GB para arquivos grandes)"],
    ["Espaço em disco", "Pasta do sistema + arquivos a processar"],
    ["Rede", "Não necessária — funcionamento 100% offline"],
]
t = Table(req, colWidths=[5*cm, 10*cm])
t.setStyle(TableStyle([
    ('BACKGROUND',   (0,0), (-1,0),  AZUL_ESCURO),
    ('TEXTCOLOR',    (0,0), (-1,0),  BRANCO),
    ('FONTNAME',     (0,0), (-1,0),  'Helvetica-Bold'),
    ('FONTSIZE',     (0,0), (-1,-1), 9),
    ('ROWBACKGROUNDS',(0,1),(-1,-1), [CINZA_CLARO, BRANCO]),
    ('GRID',         (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
    ('TOPPADDING',   (0,0), (-1,-1), 6),
    ('BOTTOMPADDING',(0,0), (-1,-1), 6),
    ('LEFTPADDING',  (0,0), (-1,-1), 8),
]))
story.append(t)

# ── 3. COMO INICIAR ───────────────────────────────────────────────────────────
story.append(Paragraph("3. Como Iniciar o Sistema", titulo_sec))
story += hr()
story.append(Paragraph(
    "O Anônimo requer um servidor local para funcionar corretamente. "
    "Os módulos de leitura de PDF (PDF.js), OCR (Tesseract) e processamento "
    "de planilhas não operam via protocolo <i>file://</i>.", corpo))

story.append(Spacer(1, 0.1*cm))
story.append(Paragraph("Procedimento de inicialização:", titulo_sub))

inicio_passos = [
    ("1", "Localize o arquivo Iniciar-Anonimo.bat",
     "Abra o Windows Explorer e navegue até a pasta do sistema."),
    ("2", "Execute com duplo-clique",
     "Clique duas vezes sobre o arquivo Iniciar-Anonimo.bat. Uma janela de terminal será aberta."),
    ("3", "Aguarde o servidor iniciar",
     "O terminal exibirá a mensagem confirmando que o servidor está rodando na porta 8000."),
    ("4", "O navegador abrirá automaticamente",
     "A aplicação será aberta em http://localhost:8000/Anonimo.html no seu navegador padrão."),
    ("5", "Para encerrar",
     "Clique na janela do terminal e pressione Ctrl+C para parar o servidor."),
]
story.append(tabela_passos(inicio_passos))
story.append(Spacer(1, 0.15*cm))
story.append(caixa_aviso(
    "<b>Atenção:</b> Não feche a janela do terminal enquanto estiver usando o sistema. "
    "Ao fechar o terminal, o servidor é encerrado e a aplicação para de funcionar."))

# ── 4. GUIA DE USO ────────────────────────────────────────────────────────────
story.append(Paragraph("4. Guia de Uso — Passo a Passo", titulo_sec))
story += hr()
story.append(Paragraph(
    "O sistema é organizado em <b>6 etapas</b> visíveis no painel lateral esquerdo. "
    "As etapas 1 a 4 formam o fluxo principal de anonimização. "
    "A etapa 5 é para reversão e a etapa 6 para processamento de PDFs.", corpo))
story.append(Spacer(1, 0.15*cm))

# Visão geral das etapas
etapas_visao = [
    ["Etapa", "Nome", "Função"],
    ["1", "Carregar Arquivo",   "Selecionar o arquivo de dados a processar"],
    ["2", "Analisar Estrutura", "Visualizar colunas e campos detectados"],
    ["3", "Selecionar Campos",  "Escolher quais dados serão anonimizados"],
    ["4", "Anonimizar Dados",   "Gerar arquivo anonimizado e chave JSON"],
    ["5", "Reverter Tokens",    "Restaurar dados originais com a chave JSON"],
    ["6", "Processar PDF / OCR","Anonimizar documentos PDF com OCR"],
]
tv = Table(etapas_visao, colWidths=[1.5*cm, 5*cm, 8.5*cm])
tv.setStyle(TableStyle([
    ('BACKGROUND',    (0,0), (-1,0),  AZUL),
    ('TEXTCOLOR',     (0,0), (-1,0),  BRANCO),
    ('FONTNAME',      (0,0), (-1,0),  'Helvetica-Bold'),
    ('FONTSIZE',      (0,0), (-1,-1), 9),
    ('ALIGN',         (0,0), (0,-1),  'CENTER'),
    ('ROWBACKGROUNDS',(0,1), (-1,-1), [CINZA_CLARO, BRANCO]),
    ('GRID',          (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
]))
story.append(tv)
story.append(Spacer(1, 0.25*cm))

# 4.1 Carregar Arquivo
story.append(Paragraph("4.1  Carregar Arquivo", titulo_sub))
story += hr(esp=3)
story.append(Paragraph(
    "Na tela inicial, você verá uma área de envio com borda tracejada. "
    "Para carregar um arquivo:", corpo))
story.append(Paragraph("• <b>Arraste</b> o arquivo diretamente para a área indicada; ou", corpo))
story.append(Paragraph("• <b>Clique</b> na área e selecione o arquivo pelo explorador de arquivos.", corpo))
story.append(Paragraph(
    "Após o carregamento, o sistema avança automaticamente para a etapa 2.", corpo))
story.append(caixa_info(
    "<b>Formatos aceitos:</b> .xlsx, .csv, .docx e .txt. "
    "Para PDF, utilize a etapa 6."))
story.append(Spacer(1, 0.15*cm))

# 4.2 Analisar Estrutura
story.append(Paragraph("4.2  Analisar Estrutura", titulo_sub))
story += hr(esp=3)
story.append(Paragraph(
    "O sistema exibe a estrutura interna do arquivo carregado: "
    "abas (Excel), colunas, tipos inferidos e amostras de valores. "
    "Esta etapa é apenas de leitura — revise os dados antes de continuar.", corpo))
story.append(Paragraph(
    "Clique em <b>Avançar</b> para ir para a seleção de campos.", corpo))
story.append(Spacer(1, 0.15*cm))

# 4.3 Selecionar Campos
story.append(Paragraph("4.3  Selecionar Campos", titulo_sub))
story += hr(esp=3)
story.append(Paragraph(
    "A tabela de seleção lista todos os campos detectados no arquivo, "
    "com classificação automática de risco:", corpo))

risco_data = [
    ["Classificação", "Cor", "Exemplos de Campos"],
    ["CRÍTICO",  "Vermelho", "CNPJ, CPF"],
    ["ALTO",     "Laranja",  "Razão Social, Nome, Empresa"],
    ["MÉDIO",    "Amarelo",  "E-mail, Endereço, Município"],
    ["BAIXO",    "Cinza",    "Campos genéricos não identificados"],
]
tr = Table(risco_data, colWidths=[3*cm, 2.5*cm, 9.5*cm])
tr.setStyle(TableStyle([
    ('BACKGROUND',    (0,0), (-1,0), AZUL_ESCURO),
    ('TEXTCOLOR',     (0,0), (-1,0), BRANCO),
    ('FONTNAME',      (0,0), (-1,0), 'Helvetica-Bold'),
    ('FONTSIZE',      (0,0), (-1,-1), 9),
    ('ROWBACKGROUNDS',(0,1), (-1,-1), [CINZA_CLARO, BRANCO]),
    ('GRID',          (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
    ('TOPPADDING',    (0,0), (-1,-1), 5),
    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
]))
story.append(tr)
story.append(Spacer(1, 0.1*cm))
story.append(Paragraph(
    "Use os <b>botões de alternância</b> (toggle) na última coluna para ativar "
    "ou desativar cada campo. Campos CRÍTICO e ALTO ficam ativados por padrão. "
    "Clique em <b>Anonimizar</b> para prosseguir.", corpo))

# 4.4 Anonimizar Dados
story.append(PageBreak())
story.append(Paragraph("4.4  Anonimizar Dados", titulo_sub))
story += hr(esp=3)
story.append(Paragraph(
    "Após a tokenização, dois arquivos ficam disponíveis para download:", corpo))

cel = estilo('CelTabela', fontSize=9, textColor=colors.HexColor("#1E293B"), leading=13)
cel_hdr = estilo('CelHeader', fontSize=9, textColor=BRANCO, fontName='Helvetica-Bold')
dl_data = [
    [Paragraph("Arquivo", cel_hdr), Paragraph("Descrição", cel_hdr)],
    [Paragraph("Salvar Arquivo", cel),
     Paragraph("Documento anonimizado no mesmo formato original, com dados sensíveis substituídos por tokens no formato [TKN-XXXXXXXX].", cel)],
    [Paragraph("Salvar Chave JSON", cel),
     Paragraph("Dicionário de reversão. Guarde com segurança — sem ele não é possível recuperar os dados originais.", cel)],
]
td = Table(dl_data, colWidths=[3.8*cm, 12.2*cm])
td.setStyle(TableStyle([
    ('BACKGROUND',    (0,0), (-1,0), AZUL_ESCURO),
    ('ROWBACKGROUNDS',(0,1), (-1,-1), [CINZA_CLARO, BRANCO]),
    ('GRID',          (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ('RIGHTPADDING',  (0,0), (-1,-1), 8),
    ('VALIGN',        (0,0), (-1,-1), 'TOP'),
]))
story.append(td)
story.append(Spacer(1, 0.1*cm))
story.append(caixa_aviso(
    "<b>Importante:</b> Salve sempre os dois arquivos. A chave JSON é indispensável "
    "para reverter a anonimização futuramente."))
story.append(Spacer(1, 0.2*cm))

# 4.5 Reverter Tokens
story.append(Paragraph("4.5  Reverter Tokens", titulo_sub))
story += hr(esp=3)
story.append(Paragraph(
    "Para restaurar os dados originais de um arquivo anonimizado:", corpo))
story.append(Paragraph(
    "1. Acesse a etapa <b>5. Reverter Tokens</b> no menu lateral.", corpo))
story.append(Paragraph(
    "2. Carregue o <b>arquivo anonimizado</b> no campo 'Arquivo Anonimizado'.", corpo))
story.append(Paragraph(
    "3. Carregue o <b>arquivo .json</b> de chaves no campo 'Chave de Reversão'.", corpo))
story.append(Paragraph(
    "4. Clique em <b>Reverter Dados</b> para baixar o arquivo restaurado.", corpo))
story.append(Spacer(1, 0.1*cm))
story.append(caixa_info(
    "A reversão é determinística: os mesmos tokens sempre produzem os mesmos "
    "valores originais, desde que a chave JSON correta seja fornecida."))
story.append(Spacer(1, 0.2*cm))

# 4.6 PDF / OCR
story.append(Paragraph("4.6  Processar PDF / OCR", titulo_sub))
story += hr(esp=3)
story.append(Paragraph(
    "Esta etapa permite anonimizar documentos PDF, incluindo PDFs digitalizados "
    "(imagens) via OCR com Tesseract.", corpo))
story.append(Paragraph(
    "• <b>PDF digital:</b> processado diretamente pelo PDF.js.", corpo))
story.append(Paragraph(
    "• <b>PDF escaneado:</b> o OCR extrai o texto das imagens antes da anonimização.", corpo))
story.append(Spacer(1, 0.1*cm))
story.append(caixa_aviso(
    "<b>Requisito:</b> Esta funcionalidade exige que o sistema esteja sendo acessado "
    "via servidor local (http://localhost:8000). Não funciona ao abrir o arquivo "
    "diretamente pelo Windows Explorer (file://)."))

# ── 5. SEGURANÇA ──────────────────────────────────────────────────────────────
story.append(Paragraph("5. Segurança e Privacidade", titulo_sec))
story += hr()

cel_v = estilo('CelVerde', fontSize=9, textColor=colors.HexColor("#1E293B"), leading=13)
cel_v_hdr = estilo('CelVerdeHdr', fontSize=9, textColor=BRANCO, fontName='Helvetica-Bold')
seg = [
    [Paragraph("Característica", cel_v_hdr), Paragraph("Detalhes", cel_v_hdr)],
    [Paragraph("Processamento local", cel_v),
     Paragraph("Todo o processamento ocorre no computador local. Nenhum dado é enviado a servidores externos.", cel_v)],
    [Paragraph("Sem dependência de rede", cel_v),
     Paragraph("O sistema funciona completamente offline após a instalação.", cel_v)],
    [Paragraph("Tokenização determinística", cel_v),
     Paragraph("O mesmo valor sempre gera o mesmo token, permitindo consistência entre arquivos.", cel_v)],
    [Paragraph("Chave reversível", cel_v),
     Paragraph("A reversão só é possível com a chave JSON gerada no momento da anonimização.", cel_v)],
    [Paragraph("Sem armazenamento persistente", cel_v),
     Paragraph("Dados processados existem apenas na memória RAM. Ao fechar o navegador, são descartados.", cel_v)],
]
ts = Table(seg, colWidths=[5*cm, 11*cm])
ts.setStyle(TableStyle([
    ('BACKGROUND',    (0,0), (-1,0), colors.HexColor("#064E3B")),
    ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.HexColor("#F0FDF4"), BRANCO]),
    ('GRID',          (0,0), (-1,-1), 0.5, colors.HexColor("#D1FAE5")),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ('RIGHTPADDING',  (0,0), (-1,-1), 8),
    ('VALIGN',        (0,0), (-1,-1), 'TOP'),
]))
story.append(ts)
story.append(Spacer(1, 0.25*cm))


# ── GERAR ─────────────────────────────────────────────────────────────────────
doc = SimpleDocTemplate(
    OUTPUT,
    pagesize=A4,
    leftMargin=2*cm, rightMargin=2*cm,
    topMargin=2*cm, bottomMargin=2.5*cm,
    title="Manual Anônimo v1.0",
    author="Coordenadoria de Inteligência Fiscal",
)

def first_page(canvas, doc):
    sem_rodape(canvas, doc)

def later_pages(canvas, doc):
    rodape(canvas, doc)

doc.build(story, onFirstPage=first_page, onLaterPages=later_pages)
print(f"Manual gerado: {OUTPUT}")
