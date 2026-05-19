"""Build PIT Global Fellows presentation using brand colors."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Brand palette
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
PINK        = RGBColor(0xF0, 0x5B, 0x89)
ORANGE      = RGBColor(0xF3, 0x70, 0x21)
BLUE        = RGBColor(0x50, 0x88, 0xC7)
GRAY        = RGBColor(0x5C, 0x57, 0x55)
LILAC       = RGBColor(0xB8, 0x92, 0xD7)
MAGENTA     = RGBColor(0xBD, 0x1A, 0x8D)
HOT_PINK    = RGBColor(0xED, 0x0C, 0x6E)
GREEN       = RGBColor(0x8C, 0xB1, 0x29)
CYAN_BLUE   = RGBColor(0x00, 0x9E, 0xDB)
TEAL        = RGBColor(0x4E, 0xBD, 0xC7)
BLACK       = RGBColor(0x00, 0x00, 0x00)
LIGHT_BG    = RGBColor(0xF7, 0xF5, 0xF9)

FONT = "Calibri"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def add_rect(slide, x, y, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, size=18, bold=False, color=GRAY,
             font=FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_bullets(slide, x, y, w, h, items, size=18, color=GRAY, bullet_color=MAGENTA):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        r1 = p.add_run()
        r1.text = "•  "
        r1.font.name = FONT
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = bullet_color
        r2 = p.add_run()
        r2.text = item
        r2.font.name = FONT
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
    return tb


def base_slide(accent=MAGENTA, top_bar=True, bg=WHITE):
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, SW, SH, bg)
    if top_bar:
        add_rect(s, 0, 0, SW, Inches(0.35), accent)
        add_rect(s, 0, Inches(0.35), SW, Inches(0.05), HOT_PINK)
    # footer
    add_text(s, Inches(0.5), Inches(7.05), Inches(8), Inches(0.35),
             "PIT Global Fellows  ·  PIT SJC 2025", size=10, color=GRAY)
    return s


def add_title(slide, title, subtitle=None, accent=MAGENTA):
    add_rect(slide, Inches(0.5), Inches(0.85), Inches(0.12), Inches(0.55), accent)
    add_text(slide, Inches(0.75), Inches(0.75), Inches(12), Inches(0.7),
             title, size=32, bold=True, color=GRAY)
    if subtitle:
        add_text(slide, Inches(0.75), Inches(1.35), Inches(12), Inches(0.45),
                 subtitle, size=15, color=PINK, bold=False)


# -------------------- SLIDE 1: COVER --------------------
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, MAGENTA)
# diagonal accent bands
add_rect(s, 0, Inches(6.7), SW, Inches(0.18), HOT_PINK)
add_rect(s, 0, Inches(6.95), SW, Inches(0.08), ORANGE)
add_rect(s, 0, Inches(7.10), SW, Inches(0.05), LILAC)
# decorative right side blocks
add_rect(s, Inches(11.2), Inches(0.5), Inches(0.18), Inches(1.4), HOT_PINK)
add_rect(s, Inches(11.5), Inches(0.5), Inches(0.18), Inches(1.4), ORANGE)
add_rect(s, Inches(11.8), Inches(0.5), Inches(0.18), Inches(1.4), LILAC)
add_rect(s, Inches(12.1), Inches(0.5), Inches(0.18), Inches(1.4), CYAN_BLUE)

add_text(s, Inches(0.8), Inches(0.6), Inches(8), Inches(0.5),
         "PIT SJC", size=14, bold=True, color=WHITE)

add_text(s, Inches(0.8), Inches(2.6), Inches(11.5), Inches(1.4),
         "PIT Global Fellows", size=64, bold=True, color=WHITE)
add_rect(s, Inches(0.8), Inches(3.95), Inches(1.5), Inches(0.08), HOT_PINK)

add_text(s, Inches(0.8), Inches(4.15), Inches(11.5), Inches(0.7),
         "Programa de Embaixadores em Inovação e Internacionalização",
         size=22, color=WHITE)

# Tag pill
pill = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                          Inches(0.8), Inches(5.4), Inches(2.5), Inches(0.5))
pill.fill.solid()
pill.fill.fore_color.rgb = HOT_PINK
pill.line.fill.background()
pill.shadow.inherit = False
tf = pill.text_frame
tf.margin_left = tf.margin_right = Emu(0)
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "PIT SJC — 2025"
r.font.name = FONT
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = WHITE


# -------------------- SLIDE 2 --------------------
s = base_slide(MAGENTA)
add_title(s, "O que é o PIT Global Fellows?",
          "Visão Geral do Projeto", MAGENTA)
add_bullets(s, Inches(0.75), Inches(2.1), Inches(12), Inches(4.5), [
    "Programa estruturado de formação, atuação prática e networking",
    "Conecta universitários do Vale do Paraíba ao ecossistema do PIT SJC",
    "Não é voluntariado operacional — é desenvolvimento real com impacto mensurável",
    "Exposição direta a delegações, startups e parceiros internacionais",
], size=20)


# -------------------- SLIDE 3 --------------------
s = base_slide(HOT_PINK)
add_title(s, "Dois problemas. Uma solução.",
          "O problema que resolvemos", HOT_PINK)

# Left column
col_y = Inches(2.15)
col_h = Inches(4.6)
add_rect(s, Inches(0.5), col_y, Inches(6.0), col_h, LIGHT_BG)
add_rect(s, Inches(0.5), col_y, Inches(0.12), col_h, MAGENTA)
add_text(s, Inches(0.85), Inches(2.3), Inches(5.5), Inches(0.5),
         "Para o PIT", size=22, bold=True, color=MAGENTA)
add_bullets(s, Inches(0.85), Inches(2.95), Inches(5.5), Inches(3.7), [
    "Equipe de 3 pessoas com demanda crescente",
    "Visitas e eventos internacionais sem cobertura adequada",
    "Trabalho estratégico comprometido por demandas operacionais",
], size=15, bullet_color=MAGENTA)

# Right column
add_rect(s, Inches(6.83), col_y, Inches(6.0), col_h, LIGHT_BG)
add_rect(s, Inches(6.83), col_y, Inches(0.12), col_h, ORANGE)
add_text(s, Inches(7.18), Inches(2.3), Inches(5.5), Inches(0.5),
         "Para os Estudantes", size=22, bold=True, color=ORANGE)
add_bullets(s, Inches(7.18), Inches(2.95), Inches(5.5), Inches(3.7), [
    "Ausência de espaços reais de prática em internacionalização",
    "Programas existentes são simulação, não contato direto",
    "Dificuldade de construir portfólio relevante ainda na graduação",
], size=15, bullet_color=ORANGE)


# -------------------- SLIDE 4 --------------------
s = base_slide(BLUE)
add_title(s, "Quem são os Fellows?", "Público-alvo", BLUE)
add_bullets(s, Inches(0.75), Inches(2.1), Inches(12), Inches(4.8), [
    "Estudantes universitários a partir do 2º semestre",
    "Cursos: RI, Administração, Engenharias, Comunicação, TI, Inovação",
    "Universidades: UNIFESP SJC, ITA, UNESP, UFABC, INATEL, Fatec SJC",
    "Perfil: curioso, proativo, interesse em inovação e mundo internacional",
    "Inglês intermediário ou avançado (desejável)",
], size=19, bullet_color=BLUE)


# -------------------- SLIDE 5 --------------------
s = base_slide(GREEN)
add_title(s, "O que o Fellow ganha", "Proposta de valor", GREEN)
add_bullets(s, Inches(0.75), Inches(2.0), Inches(12), Inches(5.0), [
    "Contato real com delegações e profissionais internacionais",
    "Certificado institucional do PIT SJC",
    "Mentoria com profissionais do ecossistema",
    "Acesso a eventos e missões normalmente fechados",
    "Portfólio real: relatórios, eventos, conteúdos publicados",
    "Networking com empresas dos clusters do parque",
    "Prática de inglês e espanhol em contexto profissional real",
], size=17, bullet_color=GREEN)


# -------------------- SLIDE 6: 4 NÚCLEOS --------------------
s = base_slide(LILAC)
add_title(s, "Como o programa se organiza",
          "Estrutura: os 4 núcleos", LILAC)

boxes = [
    ("Eventos & Protocolo",
     "Apoio a visitas, delegações e eventos internacionais", MAGENTA),
    ("Conteúdo & Comunicação",
     "Newsletter, artigos, redes sociais, relatórios", ORANGE),
    ("Relações Universitárias",
     "Articulação com IES, recrutamento de turmas", BLUE),
    ("Inteligência & Inovação",
     "Tendências globais, benchmarkings, análises", GREEN),
]
box_w = Inches(2.95)
box_h = Inches(4.2)
gap   = Inches(0.18)
total = box_w * 4 + gap * 3
start_x = Emu((SW - total) / 2)
top = Inches(2.4)
for i, (title, desc, col) in enumerate(boxes):
    x = start_x + Emu(i * (box_w + gap))
    add_rect(s, x, top, box_w, box_h, LIGHT_BG)
    add_rect(s, x, top, box_w, Inches(0.6), col)
    # number circle
    num = s.shapes.add_shape(MSO_SHAPE.OVAL,
                             x + Inches(0.25), top + Inches(0.9),
                             Inches(0.7), Inches(0.7))
    num.fill.solid(); num.fill.fore_color.rgb = col
    num.line.fill.background(); num.shadow.inherit = False
    ntf = num.text_frame
    ntf.margin_left = ntf.margin_right = Emu(0)
    ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
    np_ = ntf.paragraphs[0]; np_.alignment = PP_ALIGN.CENTER
    nr = np_.add_run(); nr.text = f"0{i+1}"
    nr.font.name = FONT; nr.font.size = Pt(16)
    nr.font.bold = True; nr.font.color.rgb = WHITE

    add_text(s, x + Inches(0.25), top + Inches(1.8),
             box_w - Inches(0.5), Inches(0.9),
             title, size=16, bold=True, color=col)
    add_text(s, x + Inches(0.25), top + Inches(2.7),
             box_w - Inches(0.5), Inches(1.4),
             desc, size=13, color=GRAY)


# -------------------- SLIDE 7: JORNADA --------------------
s = base_slide(CYAN_BLUE)
add_title(s, "Do processo seletivo ao alumni",
          "Jornada do Fellow", CYAN_BLUE)

steps = [
    ("1", "Inscrição\n& Seleção", MAGENTA),
    ("2", "Onboarding\n& Imersão", HOT_PINK),
    ("3", "Atuação\npor Núcleo", ORANGE),
    ("4", "Avaliação\nde Ciclo", GREEN),
    ("5", "Certificação", CYAN_BLUE),
    ("6", "Alumni\nNetwork", LILAC),
]
n = len(steps)
step_w = Inches(1.85)
arrow_w = Inches(0.20)
total_w = step_w * n + arrow_w * (n - 1)
start_x = Emu((SW - total_w) / 2)
top = Inches(3.0)
step_h = Inches(2.2)

for i, (num, label, col) in enumerate(steps):
    x = start_x + Emu(i * (step_w + arrow_w))
    add_rect(s, x, top, step_w, step_h, LIGHT_BG)
    add_rect(s, x, top, step_w, Inches(0.5), col)

    add_text(s, x, top + Inches(0.05), step_w, Inches(0.4),
             num, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)

    add_text(s, x + Inches(0.1), top + Inches(0.7),
             step_w - Inches(0.2), Inches(1.4),
             label, size=14, bold=True, color=GRAY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    if i < n - 1:
        ax = x + step_w
        arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                   ax, top + Inches(0.8),
                                   arrow_w, Inches(0.6))
        arrow.fill.solid(); arrow.fill.fore_color.rgb = CYAN_BLUE
        arrow.line.fill.background(); arrow.shadow.inherit = False


# -------------------- SLIDE 8 --------------------
s = base_slide(ORANGE)
add_title(s, "Estrutura leve e escalável", "Governança", ORANGE)
add_bullets(s, Inches(0.75), Inches(2.0), Inches(12), Inches(5.0), [
    "1 Coordenador interno no PIT (~3-4h/semana)",
    "Fellows Líderes por núcleo (a partir do 2º ciclo)",
    "Reunião mensal de 1h com todos os fellows",
    "Ferramentas: Notion, WhatsApp/Slack, Google Forms, Canva",
    "Carga dos fellows: 6 a 10 horas mensais",
], size=19, bullet_color=ORANGE)


# -------------------- SLIDE 9: ROADMAP --------------------
s = base_slide(TEAL)
add_title(s, "Plano de implementação — Piloto",
          "Roadmap do piloto", TEAL)

phases = [
    ("Fase 1", "Semanas 1–4", "Preparação",
     ["Aprovação interna",
      "Identidade visual",
      "Contato com universidades",
      "Validação jurídica"], MAGENTA),
    ("Fase 2", "Semanas 5–8", "Seleção & Onboarding",
     ["Inscrições abertas",
      "Dinâmica em grupo",
      "Seleção de 8–12 fellows",
      "Onboarding"], ORANGE),
    ("Fase 3", "Semanas 9–24", "Operação & Validação",
     ["Execução nos núcleos",
      "Primeiro evento real",
      "Avaliações",
      "Ajustes"], GREEN),
]
ph_w = Inches(4.0)
ph_h = Inches(4.4)
gap = Inches(0.22)
total_w = ph_w * 3 + gap * 2
start_x = Emu((SW - total_w) / 2)
top = Inches(2.3)
for i, (label, weeks, name, items, col) in enumerate(phases):
    x = start_x + Emu(i * (ph_w + gap))
    add_rect(s, x, top, ph_w, ph_h, LIGHT_BG)
    add_rect(s, x, top, ph_w, Inches(1.1), col)
    add_text(s, x + Inches(0.3), top + Inches(0.12),
             ph_w - Inches(0.6), Inches(0.4),
             label, size=14, bold=True, color=WHITE)
    add_text(s, x + Inches(0.3), top + Inches(0.5),
             ph_w - Inches(0.6), Inches(0.5),
             weeks, size=12, color=WHITE)
    add_text(s, x + Inches(0.3), top + Inches(1.25),
             ph_w - Inches(0.6), Inches(0.5),
             name, size=17, bold=True, color=col)
    add_bullets(s, x + Inches(0.3), top + Inches(1.95),
                ph_w - Inches(0.6), Inches(2.3),
                items, size=13, bullet_color=col)


# -------------------- SLIDE 10 --------------------
s = base_slide(GREEN)
add_title(s, "O que define o sucesso do piloto", "MVP", GREEN)
add_bullets(s, Inches(0.75), Inches(2.0), Inches(12), Inches(5.0), [
    "10 fellows selecionados e ativos",
    "2 núcleos operando (Eventos + Conteúdo)",
    "1 evento ou visita internacional com participação dos fellows",
    "1 ciclo completo de 6 meses",
    "1 certificado emitido",
    "1 relatório de impacto para a liderança do PIT",
], size=18, bullet_color=GREEN)


# -------------------- SLIDE 11: INDICADORES --------------------
s = base_slide(BLUE)
add_title(s, "Como mediremos o sucesso", "Indicadores", BLUE)

cols = [
    ("Operacional", MAGENTA,
     ["Nº de eventos cobertos",
      "Horas de suporte geradas",
      "Taxa de entrega por núcleo"]),
    ("Experiência", ORANGE,
     ["Taxa de renovação dos fellows (meta: +70%)",
      "NPS do programa",
      "Posts orgânicos dos fellows"]),
    ("Impacto", GREEN,
     ["Universidades parceiras formalizadas",
      "Conteúdos produzidos e publicados",
      "Alumni com oportunidades geradas pelo programa"]),
]
c_w = Inches(4.0)
c_h = Inches(4.4)
gap = Inches(0.22)
total_w = c_w * 3 + gap * 2
start_x = Emu((SW - total_w) / 2)
top = Inches(2.3)
for i, (name, col, items) in enumerate(cols):
    x = start_x + Emu(i * (c_w + gap))
    add_rect(s, x, top, c_w, c_h, LIGHT_BG)
    add_rect(s, x, top, c_w, Inches(0.8), col)
    add_text(s, x, top + Inches(0.15), c_w, Inches(0.5),
             name, size=20, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, x + Inches(0.3), top + Inches(1.05),
                c_w - Inches(0.6), Inches(3.2),
                items, size=14, bullet_color=col)


# -------------------- SLIDE 12 --------------------
s = base_slide(HOT_PINK)
add_title(s, "O que fazer agora", "Próximos passos", HOT_PINK)

steps_list = [
    "Apresentar o conceito para a liderança do PIT e obter aprovação",
    "Definir nome final e criar identidade visual mínima",
    "Verificar aspectos jurídicos do termo de adesão",
    "Mapear agenda internacional dos próximos 90 dias",
    "Contactar 2-3 universidades para sondar parceria",
    "Criar 1-pager e formulário de candidatura",
    "Abrir inscrições para o piloto",
]
top = Inches(2.05)
row_h = Inches(0.62)
colors_cycle = [MAGENTA, HOT_PINK, ORANGE, BLUE, CYAN_BLUE, GREEN, LILAC]
for i, txt in enumerate(steps_list):
    y = top + Emu(i * row_h)
    col = colors_cycle[i]
    # number circle
    num = s.shapes.add_shape(MSO_SHAPE.OVAL,
                             Inches(0.75), y,
                             Inches(0.5), Inches(0.5))
    num.fill.solid(); num.fill.fore_color.rgb = col
    num.line.fill.background(); num.shadow.inherit = False
    ntf = num.text_frame
    ntf.margin_left = ntf.margin_right = Emu(0)
    ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
    np_ = ntf.paragraphs[0]; np_.alignment = PP_ALIGN.CENTER
    nr = np_.add_run(); nr.text = str(i + 1)
    nr.font.name = FONT; nr.font.size = Pt(15)
    nr.font.bold = True; nr.font.color.rgb = WHITE

    add_text(s, Inches(1.45), y, Inches(11.5), Inches(0.5),
             txt, size=16, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)


# -------------------- SLIDE 13: CLOSING --------------------
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, MAGENTA)
add_rect(s, 0, Inches(6.7), SW, Inches(0.18), HOT_PINK)
add_rect(s, 0, Inches(6.95), SW, Inches(0.08), ORANGE)
add_rect(s, 0, Inches(7.10), SW, Inches(0.05), LILAC)

# decorative left side blocks
add_rect(s, Inches(0.5), Inches(0.5), Inches(0.18), Inches(1.4), HOT_PINK)
add_rect(s, Inches(0.8), Inches(0.5), Inches(0.18), Inches(1.4), ORANGE)
add_rect(s, Inches(1.1), Inches(0.5), Inches(0.18), Inches(1.4), GREEN)
add_rect(s, Inches(1.4), Inches(0.5), Inches(0.18), Inches(1.4), CYAN_BLUE)

add_text(s, Inches(0.8), Inches(2.4), Inches(11.5), Inches(1.4),
         "PIT Global Fellows", size=60, bold=True, color=WHITE)
add_rect(s, Inches(0.8), Inches(3.65), Inches(1.5), Inches(0.08), HOT_PINK)
add_text(s, Inches(0.8), Inches(3.85), Inches(11.5), Inches(1.2),
         "Inovação começa por quem vai transformar o mundo.",
         size=24, color=WHITE)

pill = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                          Inches(0.8), Inches(5.5), Inches(2.5), Inches(0.5))
pill.fill.solid(); pill.fill.fore_color.rgb = HOT_PINK
pill.line.fill.background(); pill.shadow.inherit = False
tf = pill.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.margin_left = tf.margin_right = Emu(0)
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "PIT SJC — 2025"
r.font.name = FONT; r.font.size = Pt(14); r.font.bold = True
r.font.color.rgb = WHITE


out = "PIT_Global_Fellows_Apresentacao.pptx"
prs.save(out)
print(f"Apresentação criada com sucesso: {out}")
print(f"Total de slides: {len(prs.slides)}")
