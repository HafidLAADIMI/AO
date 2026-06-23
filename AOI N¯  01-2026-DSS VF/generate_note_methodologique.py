# -*- coding: utf-8 -*-
"""
Note Méthodologique - Offre Technique
Objet : Conception, développement et déploiement d'une plateforme intégrée
        de suivi-évaluation de la Stratégie Génération Green 2020-2030
Référence : AOOI N°01/2026/DSS
Maître d'ouvrage : DSS - Département de l'Agriculture
Soumissionnaire : ABI CONSULTING
Document type : offre technique
Language : French
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import tempfile
import os
import zipfile
import re
import shutil
from graphviz import Digraph

# ---------------------------------------------------------------
# Constantes de couleurs
# ---------------------------------------------------------------
PRIMARY_HEX = "1B4F35"
LIGHT_BG_HEX = "EAF4EE"
WHITE_HEX = "FFFFFF"
PRIMARY_RGB = RGBColor(0x1B, 0x4F, 0x35)
WHITE_RGB = RGBColor(0xFF, 0xFF, 0xFF)
DARK_RGB = RGBColor(0x1A, 0x1A, 0x1A)
H1_COLOR = RGBColor(0x15, 0x60, 0x82)
H2H3_COLOR = RGBColor(0x0F, 0x47, 0x61)

heading_counter = [0, 0, 0]


# ---------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------

def set_document_font_theme(doc):
    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(11)
    normal.font.color.rgb = DARK_RGB
    style_el = normal.element
    rPr = style_el.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        style_el.append(rPr)
    for ef in rPr.findall(qn('w:rFonts')):
        rPr.remove(ef)
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Calibri')
    rFonts.set(qn('w:hAnsi'), 'Calibri')
    rFonts.set(qn('w:eastAsia'), 'Calibri')
    rFonts.set(qn('w:cs'), 'Calibri')
    rPr.insert(0, rFonts)
    for lvl in range(1, 5):
        try:
            h = doc.styles[f'Heading {lvl}']
            h.font.name = 'Calibri'
        except Exception:
            pass


def set_margins(doc):
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.2)


def add_page_numbers(doc):
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run()
    run.font.name = 'Calibri'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    fc1 = OxmlElement('w:fldChar')
    fc1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    instr.text = ' PAGE '
    fc2 = OxmlElement('w:fldChar')
    fc2.set(qn('w:fldCharType'), 'end')
    run._r.extend([fc1, instr, fc2])


def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for ex in tcPr.findall(qn('w:shd')):
        tcPr.remove(ex)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_text(cell, text, bold=False, color_rgb=None, size_pt=10, center=False):
    cell.text = ''
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    run = para.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color_rgb:
        run.font.color.rgb = color_rgb


def add_heading(doc, text, level, numbered=None):
    global heading_counter
    para = doc.add_paragraph(style=f'Heading {level}')
    if level == 1:
        if numbered is not None:
            n = int(numbered)
            heading_counter[0] = n
            heading_counter[1] = 0
            heading_counter[2] = 0
            display = f"{n}. {text}"
        else:
            heading_counter[1] = 0
            heading_counter[2] = 0
            display = text
    elif level == 2:
        heading_counter[1] += 1
        heading_counter[2] = 0
        display = f"{heading_counter[0]}.{heading_counter[1]}. {text}"
    else:
        heading_counter[2] += 1
        display = f"{heading_counter[0]}.{heading_counter[1]}.{heading_counter[2]}. {text}"
    run = para.add_run(display)
    run.font.name = 'Calibri'
    if level == 1:
        run.font.color.rgb = H1_COLOR
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.italic = False
    elif level == 2:
        run.font.color.rgb = H2H3_COLOR
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.italic = False
    else:
        run.font.color.rgb = H2H3_COLOR
        run.font.size = Pt(11)
        run.font.bold = False
        run.font.italic = True
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return para


def add_body(doc, text, space_before=0):
    para = doc.add_paragraph(style='Normal')
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = para.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    if space_before:
        para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(6)
    return para


def add_table(doc, headers, rows, caption=None, col_widths=None):
    if caption:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cr = cp.add_run(caption)
        cr.font.name = 'Calibri'
        cr.font.size = Pt(9)
        cr.font.italic = True
        cr.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    tbl_el = table._tbl
    tbl_pr = tbl_el.find(qn('w:tblPr'))
    if tbl_pr is None:
        tbl_pr = OxmlElement('w:tblPr')
        tbl_el.insert(0, tbl_pr)
    for ex in tbl_pr.findall(qn('w:tblW')):
        tbl_pr.remove(ex)
    tbl_w = OxmlElement('w:tblW')
    tbl_w.set(qn('w:w'), '9072')
    tbl_w.set(qn('w:type'), 'dxa')
    tbl_pr.append(tbl_w)
    jc_el = OxmlElement('w:jc')
    jc_el.set(qn('w:val'), 'center')
    tbl_pr.append(jc_el)

    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        shade_cell(cell, PRIMARY_HEX)
        set_cell_text(cell, h, bold=True, color_rgb=WHITE_RGB, size_pt=9)

    for ri, row_data in enumerate(rows):
        bg = LIGHT_BG_HEX if ri % 2 == 0 else WHITE_HEX
        row = table.rows[ri + 1]
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            shade_cell(cell, bg)
            set_cell_text(cell, str(val), size_pt=9)

    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(width)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return table


def add_figure(doc, img_path, caption, width_cm=14.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(img_path, width=Cm(width_cm))
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cr = cp.add_run(caption)
    cr.font.name = 'Calibri'
    cr.font.size = Pt(9)
    cr.font.italic = True
    cr.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)


def add_toc_field(doc):
    """
    TOC Word dynamique avec contenu visible immédiatement.
    Structure : begin + instrText + separate + [contenu manuel] + end
    Word rafraichit le contenu à l'ouverture grâce à dirty=1 + updateFields.
    LibreOffice affiche le contenu manuel en attendant.
    """
    # --- Paragraphe d'ouverture du champ (begin + instrText + separate) ---
    p_open = doc.add_paragraph()
    p_open.paragraph_format.space_before = Pt(0)
    p_open.paragraph_format.space_after = Pt(0)

    r_begin = p_open.add_run()
    fc_begin = OxmlElement('w:fldChar')
    fc_begin.set(qn('w:fldCharType'), 'begin')
    fc_begin.set(qn('w:dirty'), '1')
    r_begin._r.append(fc_begin)

    r_instr = p_open.add_run()
    instr = OxmlElement('w:instrText')
    instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r_instr._r.append(instr)

    r_sep = p_open.add_run()
    fc_sep = OxmlElement('w:fldChar')
    fc_sep.set(qn('w:fldCharType'), 'separate')
    r_sep._r.append(fc_sep)

    # --- Contenu manuel visible (entre separate et end) ---
    _add_toc_entries(doc)

    # --- Paragraphe de fermeture du champ (end) ---
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(0)
    p_end.paragraph_format.space_after = Pt(0)
    r_end = p_end.add_run()
    fc_end = OxmlElement('w:fldChar')
    fc_end.set(qn('w:fldCharType'), 'end')
    r_end._r.append(fc_end)


def _add_toc_entries(doc):
    entries = [
        (1, "1",    "Compréhension du contexte",                            4),
        (2, "1.1",  "L'organisation institutionnelle du dispositif",        4),
        (2, "1.2",  "La dimension filière x territoire",                    5),
        (2, "1.3",  "Les limites du dispositif actuel",                     6),
        (1, "2",    "Compréhension de la mission",                          7),
        (2, "2.1",  "Le cadre analytique : la GAR",                         7),
        (2, "2.2",  "Les trois missions et leur logique d'enchaînement",    8),
        (2, "2.3",  "Les indicateurs de succès de la mission",              9),
        (1, "3",    "Approche méthodologique",                              10),
        (2, "3.1",  "Mission I : Conception de la plateforme",              10),
        (3, "3.1.1","Analyse des besoins",                                  10),
        (3, "3.1.2","Conception de la plateforme",                          13),
        (2, "3.2",  "Mission II : Développement de la plateforme",          15),
        (2, "3.3",  "Mission III : Déploiement de la plateforme",           20),
        (2, "3.4",  "Pratiques d'ingénierie et philosophie Agile/DevOps",   23),
        (2, "3.5",  "Communication et collaboration interdisciplinaire",    26),
        (1, "4",    "Architecture technique proposée",                      27),
        (2, "4.1",  "Périmètre fonctionnel - Cas d'utilisation",           28),
        (2, "4.2",  "Structure des données - Diagramme de classes",        29),
        (2, "4.3",  "Circuit de validation - Diagramme de séquence",       30),
        (2, "4.4",  "Justification des choix techniques",                   31),
        (2, "4.5",  "Infrastructure, déploiement et exploitation",          33),
        (2, "4.6",  "Tableau des composants et justifications",             35),
        (1, "5",    "Équipe proposée",                                      36),
        (2, "5.1",  "Organisation et gouvernance de la mission",            36),
        (2, "5.2",  "Profils détaillés des experts clés",                   37),
        (2, "5.3",  "Tableau synthétique de l'équipe",                     38),
        (2, "5.4",  "Conformité aux exigences du CPS",                     39),
        (1, "6",    "Planning d'exécution",                                 40),
        (2, "6.1",  "Déroulement par mission",                              40),
        (2, "6.2",  "Jalons contractuels",                                  42),
        (1, "7",    "Chronogramme d'affectation du personnel",              43),
        (1, "8",    "Engagements qualité et spécificités techniques",       44),
        (2, "8.1",  "Architecture open source et propriété du MO",         44),
        (2, "8.2",  "Pipeline CI/CD avec contrôle de sécurité",            44),
        (2, "8.3",  "Audit log immuable et traçabilité",                    45),
        (2, "8.4",  "Tests de sécurité en préproduction",                   45),
        (2, "8.5",  "Isolation régionale par Row-Level Security",           46),
        (1, "9",    "Gestion des risques",                                  47),
        (2, "9.1",  "Registre des risques",                                 47),
        (2, "9.2",  "Matrice de criticité",                                 48),
        (2, "9.3",  "Niveaux de service pendant la période de garantie",   49),
    ]

    # Tab stop à droite avec points de conduite (position en twips: ~15.5cm = 8789 twips)
    TAB_POS = "8789"

    for level, num, title, page in entries:
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(3 if level == 1 else 1)
        para.paragraph_format.left_indent = Cm(0) if level == 1 else Cm(0.7)

        # Tab stop avec leaders
        pPr = para._p.get_or_add_pPr()
        tabs_el = OxmlElement('w:tabs')
        tab_el = OxmlElement('w:tab')
        tab_el.set(qn('w:val'), 'right')
        tab_el.set(qn('w:leader'), 'dot')
        tab_el.set(qn('w:pos'), TAB_POS)
        tabs_el.append(tab_el)
        pPr.append(tabs_el)

        if level == 1:
            run_num = para.add_run(f"{num}.  ")
            run_num.font.name = 'Calibri'
            run_num.font.size = Pt(11)
            run_num.font.bold = True
            run_num.font.color.rgb = PRIMARY_RGB

            run_title = para.add_run(title)
            run_title.font.name = 'Calibri'
            run_title.font.size = Pt(11)
            run_title.font.bold = True
            run_title.font.color.rgb = DARK_RGB
        else:
            run_num = para.add_run(f"{num}  ")
            run_num.font.name = 'Calibri'
            run_num.font.size = Pt(10)
            run_num.font.color.rgb = RGBColor(0x2E, 0x7D, 0x32)

            run_title = para.add_run(title)
            run_title.font.name = 'Calibri'
            run_title.font.size = Pt(10)
            run_title.font.color.rgb = DARK_RGB

        run_tab = para.add_run('\t')
        run_tab.font.name = 'Calibri'
        run_tab.font.size = Pt(10 if level == 2 else 11)

        run_page = para.add_run(str(page))
        run_page.font.name = 'Calibri'
        run_page.font.size = Pt(10 if level == 2 else 11)
        run_page.font.bold = (level == 1)
        run_page.font.color.rgb = PRIMARY_RGB if level == 1 else DARK_RGB


# ---------------------------------------------------------------
# Génération des figures
# ---------------------------------------------------------------

def generate_gantt():
    activities = [
        ("Lancement et cadrage", 0.0, 0.5, "#1B4F35"),
        ("Analyse des besoins et inventaire des indicateurs", 0.3, 1.8, "#1B4F35"),
        ("Ateliers de validation des indicateurs", 1.5, 0.8, "#1B4F35"),
        ("Rédaction SFG", 1.8, 1.0, "#1B4F35"),
        ("Rédaction SFD et note d'architecture", 2.3, 0.9, "#1B4F35"),
        ("Environnement de développement et CI/CD", 3.0, 0.5, "#2E7D32"),
        ("Module gestion des référentiels", 3.3, 1.5, "#2E7D32"),
        ("Module saisie et circuit de validation", 4.2, 1.8, "#2E7D32"),
        ("Module tableaux de bord et reporting", 5.5, 1.5, "#2E7D32"),
        ("Module alertes, notifications et échéanciers", 6.8, 1.0, "#2E7D32"),
        ("Tests intégration et sécurité DGSSI/OWASP", 7.2, 0.9, "#2E7D32"),
        ("Déploiement infrastructure et configuration", 8.0, 0.6, "#388E3C"),
        ("Migration des données et recette technique", 8.4, 1.0, "#388E3C"),
        ("Formation des utilisateurs par profil", 9.0, 1.0, "#388E3C"),
        ("Tests de recette fonctionnelle", 9.5, 1.0, "#388E3C"),
        ("Mise en production et stabilisation", 10.2, 0.8, "#388E3C"),
    ]

    # Milestones: (x, label, label_y_offset)
    milestones = [
        (3.0,  "Livrable M1 : SFG + SFD validés",   2.8),
        (8.0,  "Livrable M2 : Version bêta validée", 2.1),
        (11.0, "Livrable M3 : Réception provisoire", 1.4),
    ]

    n = len(activities)
    fig, ax = plt.subplots(figsize=(18 / 2.54, 16 / 2.54))

    # Bandes de fond par mission
    ax.axvspan(0,  3,  alpha=0.07, color='#1B4F35', zorder=0)
    ax.axvspan(3,  8,  alpha=0.07, color='#2E7D32', zorder=0)
    ax.axvspan(8, 11,  alpha=0.07, color='#388E3C', zorder=0)

    # Barres d'activités
    for i, (label, start, dur, color) in enumerate(activities):
        y = n - i
        ax.barh(y, dur, left=start, height=0.58, color=color, alpha=0.88,
                edgecolor='white', linewidth=0.5, zorder=2)

    # Lignes et labels des jalons (staggerés verticalement)
    for ms_x, ms_label, ms_y in milestones:
        ax.axvline(x=ms_x, color='#C17F00', linewidth=1.4, linestyle='--',
                   alpha=0.85, zorder=3)
        ax.plot(ms_x, n + ms_y, marker='D', color='#C17F00', markersize=5,
                zorder=5, clip_on=False)
        ax.text(ms_x + 0.15, n + ms_y, ms_label, fontsize=7.5,
                color='#C17F00', fontweight='bold', va='center', ha='left',
                zorder=6, clip_on=False)

    # Labels de missions (en haut, séparés des jalons)
    ax.text(1.5,  n + 3.5, 'MISSION I\nConception',    fontsize=8, ha='center',
            color='#1B4F35', fontweight='bold', va='center')
    ax.text(5.5,  n + 3.5, 'MISSION II\nDéveloppement', fontsize=8, ha='center',
            color='#2E7D32', fontweight='bold', va='center')
    ax.text(9.5,  n + 3.5, 'MISSION III\nDéploiement',  fontsize=8, ha='center',
            color='#388E3C', fontweight='bold', va='center')

    # Séparateurs visuels entre missions
    for x_sep in [3, 8]:
        ax.axvline(x=x_sep, color='#888888', linewidth=0.6, alpha=0.4, zorder=1)

    # Axes
    ax.set_yticks(list(range(1, n + 1)))
    ax.set_yticklabels([a[0] for a in reversed(activities)], fontsize=7.8)
    ax.set_xlim(-0.2, 11.8)
    ax.set_ylim(0.1, n + 4.3)
    ax.set_xticks(range(12))
    ax.set_xticklabels(['Départ'] + [f'M{i}' for i in range(1, 12)], fontsize=8)
    ax.set_xlabel('Mois', fontsize=9, labelpad=6)

    legend_els = [
        mpatches.Patch(facecolor='#1B4F35', alpha=0.88, label='Mission I - Conception (3 mois)'),
        mpatches.Patch(facecolor='#2E7D32', alpha=0.88, label='Mission II - Développement (5 mois)'),
        mpatches.Patch(facecolor='#388E3C', alpha=0.88, label='Mission III - Déploiement (3 mois)'),
    ]
    ax.legend(handles=legend_els, loc='lower right', fontsize=7.5, framealpha=0.9)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='x', alpha=0.2, linewidth=0.4, zorder=0)
    ax.set_title("Planning d'exécution - Plateforme SGG (11 mois)",
                 fontsize=10, fontweight='bold', pad=12)

    plt.subplots_adjust(left=0.32, right=0.97, top=0.93, bottom=0.08)
    path = os.path.join(tempfile.gettempdir(), 'gantt_sgg.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    return path


def generate_architecture_diagram():
    dot = Digraph('architecture_sgg')
    dot.attr(rankdir='TB', nodesep='0.45', ranksep='0.75', dpi='200',
             fontname='Helvetica', bgcolor='white')

    node_box = {'shape': 'box', 'style': 'filled', 'fontname': 'Helvetica',
                'fontsize': '9', 'margin': '0.15,0.1'}
    node_cyl = {'shape': 'cylinder', 'style': 'filled', 'fontname': 'Helvetica',
                'fontsize': '9', 'margin': '0.15,0.1'}

    with dot.subgraph(name='cluster_frontend') as c:
        c.attr(label='Couche Présentation', style='rounded,filled',
               fillcolor='#F1F8E9', color='#558B2F', fontname='Helvetica', fontsize='9')
        c.node('nginx', 'Nginx\n(reverse proxy / TLS)', fillcolor='#DCEDC8', **node_box, width='1.9')
        c.node('react', 'React + Vite\n(SPA)', fillcolor='#DCEDC8', **node_box, width='1.9')
        c.edge('react', 'nginx', label='HTTP')

    with dot.subgraph(name='cluster_gateway') as c:
        c.attr(label='API Gateway', style='rounded,filled',
               fillcolor='#E3F2FD', color='#1565C0', fontname='Helvetica', fontsize='9')
        c.node('gateway', 'Spring Cloud Gateway\n(validation JWT Keycloak)', fillcolor='#BBDEFB',
               **node_box, width='2.5')

    with dot.subgraph(name='cluster_iam') as c:
        c.attr(label='IAM', style='rounded,filled',
               fillcolor='#FFF8E1', color='#F57F17', fontname='Helvetica', fontsize='9')
        c.node('keycloak', 'Keycloak\n(realm sgg-platform\n6 rôles, RBAC)', fillcolor='#FFE082',
               **node_box, width='2.2')

    with dot.subgraph(name='cluster_services') as c:
        c.attr(label='Microservices', style='rounded,filled',
               fillcolor='#F3E5F5', color='#6A1B9A', fontname='Helvetica', fontsize='9',
               rank='same')
        c.node('api', 'API Principale\n(Kotlin / Spring Boot)\nLogique métier & référentiels',
               fillcolor='#E1BEE7', **node_box, width='2.3')
        c.node('audit', 'Service Audit\n(Kotlin / Spring Boot)\nJournal immuable',
               fillcolor='#E1BEE7', **node_box, width='2.3')
        c.node('notif', 'Service Notification\n(Kotlin / Spring Boot)\nAlertes & emails',
               fillcolor='#E1BEE7', **node_box, width='2.3')
        c.node('rapport', 'Service Rapport\n(Python / FastAPI)\nExport PDF / Excel',
               fillcolor='#E1BEE7', **node_box, width='2.3')

    with dot.subgraph(name='cluster_msg') as c:
        c.attr(label='Messagerie asynchrone', style='rounded,filled',
               fillcolor='#FCE4EC', color='#880E4F', fontname='Helvetica', fontsize='9')
        c.node('rabbitmq', 'RabbitMQ\n(broker, 6 événements)', fillcolor='#F8BBD0',
               **node_box, width='2.2')

    with dot.subgraph(name='cluster_data') as c:
        c.attr(label='Couche Données', style='rounded,filled',
               fillcolor='#E8EAF6', color='#1A237E', fontname='Helvetica', fontsize='9')
        c.node('postgres', 'PostgreSQL + TimescaleDB\nRow-Level Security / région',
               fillcolor='#C5CAE9', **node_cyl, width='2.3')
        c.node('mongodb', 'MongoDB\n(audit logs append-only)', fillcolor='#C5CAE9',
               **node_cyl, width='2.3')
        c.node('garage', 'Garage\n(stockage S3, pièces justificatives)', fillcolor='#C5CAE9',
               **node_cyl, width='2.3')
        c.node('reporting', 'Schéma reporting\n(vues PostgreSQL read-only)',
               shape='box', fillcolor='#C5CAE9', style='filled,dashed',
               fontname='Helvetica', fontsize='9', margin='0.15,0.1', width='2.3')

    dot.edge('nginx', 'gateway', label='HTTPS')
    dot.edge('gateway', 'keycloak', label='validation token', style='dashed', color='#F57F17')
    dot.edge('gateway', 'api', label='REST')
    dot.edge('gateway', 'rapport', label='REST')
    dot.edge('api', 'rabbitmq', label='publish')
    dot.edge('rabbitmq', 'audit', label='consume')
    dot.edge('rabbitmq', 'notif', label='consume')
    dot.edge('api', 'postgres', label='R/W')
    dot.edge('api', 'garage', label='S3 API')
    dot.edge('audit', 'mongodb', label='append')
    dot.edge('rapport', 'reporting', label='SELECT')
    dot.edge('reporting', 'postgres', label='views', style='dashed', dir='none')

    path_base = os.path.join(tempfile.gettempdir(), 'architecture_sgg')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_organigramme():
    dot = Digraph('organigramme')
    dot.attr(rankdir='TB', nodesep='0.6', ranksep='0.55', dpi='200',
             fontname='Helvetica', bgcolor='white')

    common = {'shape': 'box', 'style': 'filled', 'fontname': 'Helvetica',
              'fontsize': '9', 'margin': '0.2,0.12'}

    dot.node('cdp', 'Brahim EL ORF\nChef de projet',
             fillcolor='#1B4F35', fontcolor='white', width='2.4', **common)

    with dot.subgraph() as s:
        s.attr(rank='same')
        s.node('se', 'Hassan KAMIL\nExpert Suivi-évaluation',
               fillcolor='#2E7D32', fontcolor='white', width='2.1', **common)
        s.node('bi', 'Hassan EL BAHI\nExpert BI / Data Analytics',
               fillcolor='#2E7D32', fontcolor='white', width='2.1', **common)
        s.node('arch', 'Abdelouahab AZIZ\nArchitecte SI',
               fillcolor='#388E3C', fontcolor='white', width='2.1', **common)

    dot.node('dev1', 'Mohammed OUAZZANE\nDéveloppeur Full-stack 1',
             fillcolor='#689F38', fontcolor='white', width='2.2', **common)
    dot.node('dev2', 'Noureddine AMENZOU\nDéveloppeur Full-stack 2',
             fillcolor='#689F38', fontcolor='white', width='2.2', **common)

    for n in ['se', 'bi', 'arch']:
        dot.edge('cdp', n)
    dot.edge('arch', 'dev1')
    dot.edge('arch', 'dev2')

    path_base = os.path.join(tempfile.gettempdir(), 'organigramme_sgg')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_usecase_diagram():
    dot = Digraph('usecase')
    dot.attr(rankdir='LR', nodesep='0.35', ranksep='1.1', dpi='200',
             fontname='Helvetica', bgcolor='white')

    actors = [
        ('saisisseur',     'SAISISSEUR',           '#1B4F35'),
        ('val_reg',        'VALIDATEUR\nRÉGIONAL', '#2E7D32'),
        ('val_cen',        'VALIDATEUR\nCENTRAL',  '#388E3C'),
        ('decideur',       'DÉCIDEUR',              '#1565C0'),
        ('admin',          'ADMIN',                 '#37474F'),
    ]
    for aid, alabel, acolor in actors:
        dot.node(aid, alabel, shape='box', style='filled',
                 fillcolor=acolor, fontcolor='white',
                 fontname='Helvetica', fontsize='9', width='1.3', height='0.45')

    with dot.subgraph(name='cluster_system') as s:
        s.attr(label='Plateforme de Suivi SGG', style='rounded',
               color='#1B4F35', fontname='Helvetica', fontsize='10', bgcolor='#F4FAF5')
        ucs = [
            ('uc_saisir',    'Saisir les données\nd\'indicateurs'),
            ('uc_importer',  'Importer des données\n(CSV / Excel)'),
            ('uc_soumettre', 'Soumettre pour\nvalidation'),
            ('uc_valider',   'Valider les données'),
            ('uc_rejeter',   'Rejeter avec motif'),
            ('uc_tdb',       'Consulter tableaux\nde bord'),
            ('uc_export',    'Exporter rapport\nPDF / Excel'),
            ('uc_ref',       'Gérer le référentiel\ndes indicateurs'),
            ('uc_users',     'Gérer utilisateurs\net droits'),
            ('uc_circuit',   'Paramétrer circuit\nde validation'),
            ('uc_audit',     'Consulter journal\nd\'audit'),
        ]
        for uid, ulabel in ucs:
            s.node(uid, ulabel, shape='ellipse', style='filled',
                   fillcolor='#EAF4EE', color='#1B4F35',
                   fontname='Helvetica', fontsize='8')

    assoc = [
        ('saisisseur', 'uc_saisir'), ('saisisseur', 'uc_importer'), ('saisisseur', 'uc_soumettre'),
        ('val_reg',    'uc_valider'), ('val_reg', 'uc_rejeter'), ('val_reg', 'uc_tdb'),
        ('val_cen',    'uc_valider'), ('val_cen', 'uc_tdb'), ('val_cen', 'uc_export'),
        ('decideur',   'uc_tdb'), ('decideur', 'uc_export'),
        ('admin',      'uc_ref'), ('admin', 'uc_users'), ('admin', 'uc_circuit'), ('admin', 'uc_audit'),
    ]
    for src, tgt in assoc:
        dot.edge(src, tgt, arrowhead='none')

    path_base = os.path.join(tempfile.gettempdir(), 'usecase_sgg')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_class_diagram():
    dot = Digraph('classes')
    dot.attr(rankdir='TB', nodesep='0.55', ranksep='0.65', dpi='200',
             fontname='Helvetica', bgcolor='white')

    rec = {'shape': 'record', 'style': 'filled', 'fillcolor': '#EAF4EE',
           'fontname': 'Helvetica', 'fontsize': '8', 'color': '#1B4F35'}

    dot.node('Indicateur',
        '{Indicateur|id: Long\\lcode: String\\llibellé: String\\lformule: String\\lunité: String\\lfrequence: Enum\\lniveauDesagreg.: Enum\\l|+ getValeurCible(): Decimal\\l+ estActif(): Boolean\\l}',
        **rec)
    dot.node('Programme',
        '{Programme|id: Long\\lcode: String\\llibellé: String\\lfondement: Enum\\l\\{GÉNÉRATION, GREEN\\}\\l}',
        **rec)
    dot.node('Filiere',
        '{Filière|id: Long\\lcode: String\\llibellé: String\\l}',
        **rec)
    dot.node('Region',
        '{Région|id: Long\\lcode: String\\llibellé: String\\l}',
        **rec)
    dot.node('SaisieDonnee',
        '{SaisieDonnée|id: Long\\lvaleur: Decimal\\lpériode: String\\lstatut: Enum\\ldateCreation: DateTime\\ldateSoumission: DateTime\\l|+ soumettre()\\l+ retirer()\\l}',
        **rec)
    dot.node('ValidationCircuit',
        '{ValidationCircuit|id: Long\\lniveauxRequis: Integer\\lniveauActuel: Integer\\lstatut: Enum\\l\\{EN_ATTENTE, VALIDÉ, REJETÉ\\}\\l}',
        **rec)
    dot.node('EntreeAudit',
        '{EntréeAudit|id: Long\\lutilisateurId: String\\lrôle: String\\laction: String\\lancienneValeur: JSON\\lnouvelleValeur: JSON\\lhorodatage: DateTime\\l}',
        **rec)
    dot.node('Alerte',
        '{Alerte|id: Long\\ltype: Enum\\lmessage: String\\ldateEchéance: Date\\lstatut: Enum\\l}',
        **rec)

    dot.edge('Indicateur', 'Programme',        label='appartient à (N-N)', dir='none')
    dot.edge('Indicateur', 'Filiere',          label='filière (0..1)',     dir='none', style='dashed')
    dot.edge('SaisieDonnee', 'Indicateur',     label='1',   arrowhead='open', arrowtail='none')
    dot.edge('SaisieDonnee', 'Region',         label='1',   arrowhead='open', arrowtail='none')
    dot.edge('SaisieDonnee', 'ValidationCircuit', label='1..1', arrowhead='none')
    dot.edge('SaisieDonnee', 'EntreeAudit',    label='génère', style='dashed', arrowhead='open')
    dot.edge('SaisieDonnee', 'Alerte',         label='déclenche', style='dashed', arrowhead='open')

    path_base = os.path.join(tempfile.gettempdir(), 'classes_sgg')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_sequence_diagram():
    participants = [
        ('SAISISSEUR',      '#1B4F35'),
        ('API Principale',  '#2E7D32'),
        ('PostgreSQL',      '#1565C0'),
        ('RabbitMQ',        '#880E4F'),
        ('Svc Notification','#E65100'),
        ('VALIDATEUR',      '#4A148C'),
        ('Svc Audit',       '#37474F'),
    ]
    messages = [
        (0, 1, 'POST /saisies  {valeur, période, indicateurId}'),
        (1, 2, 'INSERT saisie  (statut = BROUILLON)'),
        (0, 1, 'POST /saisies/{id}/soumettre'),
        (1, 2, 'UPDATE statut = EN_ATTENTE'),
        (1, 3, 'publish  data.submitted'),
        (3, 4, 'consume  data.submitted'),
        (4, 5, 'Email: saisie en attente de validation'),
        (5, 1, 'POST /saisies/{id}/valider  (ou /rejeter)'),
        (1, 2, 'UPDATE statut = VALIDÉ  (ou REJETÉ)'),
        (1, 3, 'publish  data.validated  (ou data.rejected)'),
        (3, 6, 'consume → INSERT log MongoDB  (append-only)'),
    ]

    np_ = len(participants)
    nm  = len(messages)
    COL = 2.3
    ROW = 1.05

    fig_w = (np_ * COL + 0.8) / 2.54
    fig_h = (nm  * ROW + 3.2) / 2.54
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))

    X = [i * COL for i in range(np_)]
    ax.set_xlim(-1.4, X[-1] + 1.4)
    ax.set_ylim(-(nm * ROW + 0.8), 2.1)
    ax.axis('off')

    for i, (name, color) in enumerate(participants):
        ax.add_patch(plt.Rectangle((X[i] - 0.95, 0.65), 1.90, 0.85,
                     facecolor=color, edgecolor='white', linewidth=0.5, zorder=2))
        ax.text(X[i], 1.08, name, ha='center', va='center',
                fontsize=7, color='white', fontweight='bold', zorder=3)
        ax.plot([X[i], X[i]], [0.65, -(nm * ROW + 0.5)],
                color='#BBBBBB', linewidth=0.7, linestyle=(0, (5, 3)), zorder=1)

    for mi, (src, tgt, label) in enumerate(messages):
        y  = -(mi + 1) * ROW
        xs, xt = X[src], X[tgt]
        ax.annotate('', xy=(xt, y), xytext=(xs, y),
                    arrowprops=dict(arrowstyle='->', color='#1A3A2A', lw=0.85))
        ax.text((xs + xt) / 2, y + 0.13, f'{mi+1}. {label}',
                ha='center', va='bottom', fontsize=6.2, color='#1A3A2A')

    ax.set_title("Circuit de validation d'une saisie d'indicateur",
                 fontsize=9, fontweight='bold', y=0.99)
    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'sequence_sgg.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    return path


def generate_sgg_structure():
    """Diagramme hiérarchique de la structure SGG : 2 fondements x programmes."""
    g = Digraph('sgg', format='png')
    g.attr(rankdir='TB', dpi='250', bgcolor='white',
           nodesep='0.35', ranksep='0.6', fontname='Calibri')

    g.node('SGG',
           label='Stratégie Génération Green\n2020-2030',
           shape='rectangle', style='filled',
           fillcolor='#1B4F35', fontcolor='white',
           fontsize='12', fontname='Calibri Bold',
           width='4.5', height='0.7')

    g.node('GEN', label='Fondement I\nGÉNÉRATION',
           shape='rectangle', style='filled',
           fillcolor='#2E7D32', fontcolor='white',
           fontsize='11', fontname='Calibri Bold', width='3.2', height='0.65')
    g.node('GREEN', label='Fondement II\nGREEN',
           shape='rectangle', style='filled',
           fillcolor='#388E3C', fontcolor='white',
           fontsize='11', fontname='Calibri Bold', width='3.2', height='0.65')

    g.edge('SGG', 'GEN', color='#1B4F35', penwidth='1.8')
    g.edge('SGG', 'GREEN', color='#1B4F35', penwidth='1.8')

    gen_progs = [
        ('G1', 'Classe Moyenne Agricole\n(Assurance, Protection sociale, SMAG)'),
        ('G2', 'Jeunes Entrepreneurs Agricoles\n(Terres collectives, FDA, Formation)'),
        ('G3', 'Organisations Agricoles\n(Coopératives, Agrégation)'),
        ('G4', 'Mécanismes d\'Accompagnement\n(Conseil agricole, Agriculture solidaire)'),
    ]
    for nid, label in gen_progs:
        g.node(nid, label=label,
               shape='rectangle', style='filled,rounded',
               fillcolor='#EAF4EE', fontcolor='#1B4F35',
               fontsize='10', fontname='Calibri', width='3.6')
        g.edge('GEN', nid, color='#2E7D32', penwidth='1.2')

    grn_progs = [
        ('V1', 'Consolidation des Filières\n(Contrats-programmes, chaîne de valeur)'),
        ('V2', 'Chaînes de Distribution\n(Marchés de gros, souks, agropoles)'),
        ('V3', 'Qualité, Innovation et Green Tech\n(R&D, abattoirs, contrôle sanitaire)'),
        ('V4', 'Agriculture Résiliente\n(Irrigation, énergies renouvelables, conservation)'),
    ]
    for nid, label in grn_progs:
        g.node(nid, label=label,
               shape='rectangle', style='filled,rounded',
               fillcolor='#EAF4EE', fontcolor='#1B4F35',
               fontsize='10', fontname='Calibri', width='3.6')
        g.edge('GREEN', nid, color='#388E3C', penwidth='1.2')

    g.node('TRANS',
           label='Programmes Transverses : Systèmes d\'aide | Financement | Gouvernance | Fiscalité | Transformation digitale',
           shape='rectangle', style='filled,rounded',
           fillcolor='#F5F5F5', fontcolor='#37474F',
           fontsize='9', fontname='Calibri', width='10.0')
    g.edge('SGG', 'TRANS', color='#90A4AE', penwidth='1.0', style='dashed')

    path = os.path.join(tempfile.gettempdir(), 'sgg_structure.png')
    g.render(path.replace('.png', ''), format='png', cleanup=True)
    return path


def generate_circuit_collecte():
    """Diagramme du circuit de collecte et de validation DPA-DRA-DSS."""
    g = Digraph('circuit', format='png')
    g.attr(rankdir='LR', dpi='250', bgcolor='white',
           nodesep='0.8', ranksep='1.2', fontname='Calibri')

    levels = [
        ('DPA', 'Niveau Provincial\n(DPA)',
         'Saisie des données\npar programme et filière\n75+ Directions Provinciales',
         '#1565C0', 'white'),
        ('DRA', 'Niveau Régional\n(DRA)',
         'Consolidation régionale\nValidation des remontées\n12 Directions Régionales',
         '#1B4F35', 'white'),
        ('DSS', 'Niveau Central\n(DSS)',
         'Consolidation nationale\nValidation finale\nTableaux de bord décideurs',
         '#880E4F', 'white'),
    ]

    for nid, title, desc, color, fc in levels:
        label = f'<<B>{title}</B><BR/><FONT POINT-SIZE="9">{desc.replace(chr(10), "<BR/>")}</FONT>>'
        g.node(nid, label=label,
               shape='rectangle', style='filled',
               fillcolor=color, fontcolor=fc,
               fontsize='11', fontname='Calibri',
               width='3.4', height='1.6')

    g.edge('DPA', 'DRA',
           label='  Soumission\n  (mensuelle/\n  trimestrielle)',
           color='#1B4F35', penwidth='2.0',
           fontsize='9', fontname='Calibri', fontcolor='#1B4F35')
    g.edge('DRA', 'DSS',
           label='  Validation\n  régionale\n  et agrégation',
           color='#1B4F35', penwidth='2.0',
           fontsize='9', fontname='Calibri', fontcolor='#1B4F35')
    g.edge('DRA', 'DPA',
           label='  Rejet avec motif\n  ou demande\n  de clarification',
           color='#C62828', penwidth='1.2', style='dashed',
           fontsize='8', fontname='Calibri', fontcolor='#C62828')
    g.edge('DSS', 'DRA',
           label='  Rejet ou\n  demande\n  de clarification',
           color='#C62828', penwidth='1.2', style='dashed',
           fontsize='8', fontname='Calibri', fontcolor='#C62828')

    path = os.path.join(tempfile.gettempdir(), 'circuit_collecte.png')
    g.render(path.replace('.png', ''), format='png', cleanup=True)
    return path


def generate_gar_chain():
    """Chaîne de résultats GAR adaptée à la SGG (matplotlib horizontal)."""
    fig, ax = plt.subplots(figsize=(22 / 2.54, 11 / 2.54))
    ax.set_xlim(0, 22)
    ax.set_ylim(0, 11)
    ax.axis('off')

    levels = [
        (0.5,  '#1565C0', 'INTRANTS',
         'Ressources mobilisées',
         ['Budgets FDA et programmes', 'Conseillers agricoles', 'Superficies aménagées', 'Financements publics']),
        (5.9,  '#1B4F35', 'EXTRANTS',
         'Livrables directs',
         ['Coopératives créées', 'Jeunes accompagnés', 'Superficies assurées', 'Agropoles opérationnels']),
        (11.3, '#2E7D32', 'EFFETS',
         'Changements à moyen terme',
         ['Revenus agricoles accrus', "Taux d'adoption irrigation", 'Regroupement agriculteurs', 'Emplois créés']),
        (16.7, '#880E4F', 'IMPACTS',
         'Transformations à long terme',
         ['Contribution PIB agricole', 'Réduction pauvreté rurale', 'Sécurité alimentaire', 'Balance commerciale']),
    ]

    box_w, box_h = 4.8, 9.5
    for x, color, title, subtitle, items in levels:
        ax.add_patch(mpatches.FancyBboxPatch(
            (x, 0.8), box_w, box_h,
            boxstyle='round,pad=0.12',
            facecolor=color, edgecolor='white', linewidth=1.8, zorder=2))
        ax.text(x + box_w / 2, 9.7, title,
                ha='center', va='center', fontsize=11, fontweight='bold',
                color='white', zorder=3)
        ax.text(x + box_w / 2, 8.9, subtitle,
                ha='center', va='center', fontsize=8.5,
                color='#E8F5E9', zorder=3, style='italic')
        ax.plot([x + 0.4, x + box_w - 0.4], [8.4, 8.4],
                color='white', linewidth=0.9, alpha=0.5, zorder=3)
        for i, item in enumerate(items):
            ax.text(x + 0.5, 7.7 - i * 1.5, f'- {item}',
                    ha='left', va='center', fontsize=8.5,
                    color='white', zorder=3)

    for x, _, _, _, _ in levels[:-1]:
        ax.annotate('', xy=(x + box_w + 0.35, 5.5),
                    xytext=(x + box_w + 0.08, 5.5),
                    arrowprops=dict(arrowstyle='->', color='#455A64',
                                   lw=2.2, mutation_scale=18))

    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'gar_chain.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    return path


# ---------------------------------------------------------------
# Sections du document
# ---------------------------------------------------------------

def page_de_garde(doc):
    # La page de garde est injectée depuis le template via inject_template_cover()
    pass


def section_sommaire(doc):
    add_heading(doc, "Sommaire", 1)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    add_toc_field(doc)
    doc.add_page_break()


def section_presentation_cabinet(doc):
    add_heading(doc, "Présentation du cabinet ABI CONSULTING", 1, numbered="1")

    add_body(doc,
        "ABI CONSULTING est une société à responsabilité limitée basée à Temara, fondée et dirigée "
        "par Ahmed Ben Hammou. Le cabinet intervient dans trois domaines: la conception et le déploiement "
        "de systèmes d'information de suivi-évaluation pour des politiques publiques, l'appui institutionnel "
        "aux maîtres d'ouvrage publics et parapublics, et la gestion de projets de développement en milieu "
        "rural et agricole. Son réseau opérationnel compte plus de trente experts permanents et associés, "
        "couvrant les profils techniques et thématiques requis par ce marché."
    )

    add_body(doc,
        "Le cabinet intervient au Maroc sur l'ensemble des douze régions, ainsi qu'en Afrique subsaharienne "
        "et dans la région MENA. Ses missions sont réalisées pour des donneurs d'ordre internationaux "
        "(Expertise France, GIZ-Maroc) et des institutions nationales (agences urbaines, opérateurs publics, "
        "organisations professionnelles agricoles)."
    )

    add_body(doc,
        "Le profil d'ABI CONSULTING répond directement aux trois composantes du présent marché. "
        "La Mission I mobilise l'expérience du cabinet dans la modélisation de systèmes de suivi "
        "(projets SABIL et TAWJIH MAROC). La Mission II s'appuie sur des réalisations comparables "
        "en développement de plateformes numériques (DigiTPME pour GIZ-Maroc). La Mission III est "
        "documentée par des réceptions provisoires réussies sur des systèmes à périmètre similaire."
    )

    add_heading(doc, "Références de missions comparables", 2)

    headers = ["Client", "Objet de la mission", "Année", "Livrable principal"]
    rows = [
        ["Expertise France", "Développement du système SABIL de suivi-évaluation", "2023",
         "Plateforme S&E opérationnelle, manuel d'utilisation"],
        ["GIZ-Maroc", "Conception et déploiement de la plateforme DigiTPME", "2022",
         "Plateforme web, module de formation en ligne"],
        ["Agence Urbaine Rabat-Salé", "Développement du système d'information géographique", "2021",
         "SIG opérationnel, base de données spatiales"],
        ["TAWJIH MAROC", "Système de suivi-évaluation des programmes d'orientation", "2021",
         "Plateforme S&E, tableaux de bord de pilotage"],
    ]
    add_table(doc, headers, rows,
              caption="Tableau 1. Références pertinentes d'ABI CONSULTING",
              col_widths=[3.5, 5.0, 1.5, 5.0])


def section_comprehension_contexte(doc, sgg_path, circuit_path):
    add_heading(doc, "Compréhension du contexte", 1, numbered="1")

    add_body(doc,
        "Le Maroc a lancé la Stratégie Génération Green 2020-2030 comme successeur direct du Plan "
        "Maroc Vert, dont elle consolide les acquis tout en élargissant l'ambition. Le Plan Maroc "
        "Vert avait démontré que la contractualisation avec les filières agricoles et le soutien "
        "à l'agrégation produisaient des résultats mesurables en matière de production et d'export. "
        "La Stratégie Génération Green va plus loin : elle intègre une dimension de développement "
        "humain que le PMV n'avait pas traitée avec la même priorité, en plaçant l'émergence "
        "d'une classe moyenne agricole et l'inclusion des jeunes ruraux au coeur de ses objectifs. "
        "C'est cette double ambition, à la fois économique et sociale, qui rend le suivi "
        "de la SGG plus complexe que celui de son prédécesseur."
    )
    add_body(doc,
        "La stratégie s'articule autour de deux fondements complémentaires, chacun décliné en "
        "quatre programmes, auxquels s'ajoutent six programmes transverses couvrant les mécanismes "
        "d'aide, le financement, la gouvernance, la fiscalité et la transformation digitale du secteur. "
        "Chaque programme est formalisé par des contrats-programmes signés avec les interprofessions "
        "et l'industrie agroalimentaire, ce qui fait de la stratégie un cadre contractuel autant "
        "qu'une politique sectorielle. Le suivi doit donc rendre compte non seulement des indicateurs "
        "de résultat nationaux, mais aussi du degré de réalisation des engagements contractuels "
        "par filière et par région."
    )
    add_figure(doc, sgg_path,
               "Figure 1. Structure programmatique de la Stratégie Génération Green 2020-2030",
               width_cm=15.0)

    add_body(doc,
        "Le Fondement I, dit Génération, cible le développement du capital humain agricole. "
        "Il regroupe quatre programmes : l'émergence d'une classe moyenne agricole, au travers "
        "de l'assurance agricole, de la protection sociale et de l'amélioration du statut de "
        "l'agriculteur ; la nouvelle génération de jeunes entrepreneurs agricoles, en mobilisant "
        "les terres collectives, les mécanismes de financement FDA et les structures de formation ; "
        "la nouvelle génération d'organisations agricoles, en renforçant les coopératives et "
        "les formes d'agrégation ; et les mécanismes d'accompagnement, au travers du conseil "
        "agricole public et de l'agriculture solidaire. Ces quatre programmes génèrent des "
        "indicateurs de nature très différente : financiers, sociaux, démographiques et "
        "institutionnels, souvent collectés par des administrations différentes."
    )
    add_body(doc,
        "Le Fondement II, dit Green, vise la compétitivité et la résilience du secteur agricole. "
        "Il couvre la consolidation des filières agricoles dans le cadre des contrats-programmes, "
        "la modernisation des chaînes de distribution et de commercialisation, le développement "
        "de la qualité, de l'innovation et de l'agriculture de précision, et la transition vers "
        "une agriculture résiliente et éco-efficiente face aux contraintes hydriques et climatiques. "
        "Ces programmes produisent des indicateurs agronomiques, économiques et environnementaux "
        "dont les sources de données sont dispersées entre la DSS, l'ONSSA, les interprofessions "
        "et les organismes de bassin hydraulique."
    )

    add_heading(doc, "L'organisation institutionnelle du dispositif de suivi", 2)

    add_body(doc,
        "Le dispositif de collecte des données de suivi repose sur trois niveaux administratifs "
        "imbriqués. Chaque niveau a un rôle distinct dans la chaîne de production des données, "
        "et les interactions entre ces niveaux définissent les exigences fonctionnelles les "
        "plus critiques de la plateforme."
    )
    add_body(doc,
        "Au niveau provincial, les Directions Provinciales de l'Agriculture (DPA) sont les "
        "points de collecte primaires. Plus de soixante-quinze DPA couvrent l'ensemble du "
        "territoire national. Ce sont leurs agents qui saisissent les données de terrain : "
        "superficies emblavées, nombres de bénéficiaires des programmes, volumes de production "
        "par filière, états d'avancement des contrats-programmes locaux. La capacité de ces "
        "agents en matière de saisie numérique est variable selon les provinces, et les "
        "contraintes de connectivité dans certaines zones rurales doivent être prises en compte "
        "dans la conception des interfaces."
    )
    add_body(doc,
        "Au niveau régional, les Directions Régionales de l'Agriculture (DRA) assurent la "
        "consolidation des remontées provinciales et exercent un premier niveau de validation. "
        "Les douze DRA vérifient la cohérence des données soumises par les DPA de leur ressort, "
        "identifient les anomalies, sollicitent les corrections nécessaires et transmettent les "
        "données validées au niveau central. Elles produisent également leurs propres données "
        "régionales, notamment sur les indicateurs de performance des filières à l'échelle régionale."
    )
    add_body(doc,
        "Au niveau central, la Direction de la Stratégie et des Statistiques (DSS) "
        "assure la consolidation nationale, la validation finale, et la production des "
        "tableaux de bord destinés aux décideurs du Département de l'Agriculture et aux "
        "instances de gouvernance nationale. La DSS est également responsable de la "
        "définition du référentiel des indicateurs, de l'arbitrage des méthodes de calcul "
        "et de la reddition de comptes auprès des partenaires institutionnels."
    )
    add_figure(doc, circuit_path,
               "Figure 2. Circuit de collecte et de validation des données de suivi SGG",
               width_cm=15.0)

    add_heading(doc, "La dimension filière x territoire : le coeur de la complexité analytique", 2)

    add_body(doc,
        "La complexité analytique de la Stratégie Génération Green tient en grande partie "
        "à l'intersection de deux dimensions de suivi : la dimension filière et la dimension "
        "territoriale. Vingt et une filières agricoles font l'objet de contrats-programmes "
        "dans le cadre de la stratégie, couvrant aussi bien les filières végétales "
        "(arboricole, céréalière, oléicole, maraîchère de primeurs, palmier dattier, "
        "safran, arganier, rose à parfum, rizicole, semencière, sucrière) que les filières "
        "animales (bovines, ovines, avicole, laitière, apicole, cameline, viandes rouges). "
        "Chacune de ces filières a sa propre chaîne de valeur, ses propres indicateurs "
        "de production, de commercialisation et de valorisation, et ses propres engagements "
        "contractuels vis-à-vis du Département de l'Agriculture."
    )
    add_body(doc,
        "Cette dimension filière se croise avec la dimension territoriale des douze régions "
        "du Maroc, dont les profils agro-climatiques, les spécialisations agricoles et les "
        "capacités institutionnelles sont très différents. La filière céréalière dans la "
        "région de Meknès-Fès n'a pas les mêmes caractéristiques de collecte que dans "
        "la région de Souss-Massa, et les indicateurs de rendement plausibles varient "
        "significativement d'une zone agro-écologique à l'autre. La plateforme doit "
        "intégrer ces différences dans ses règles de validation automatique, sous peine "
        "de générer des faux positifs ou de laisser passer des anomalies réelles."
    )

    perim_headers = ["Dimension", "Périmètre", "Implications pour la plateforme"]
    perim_rows = [
        ["Fondements", "2 fondements (Génération et Green)",
         "Deux logiques de suivi différentes : développement humain versus compétitivité économique"],
        ["Programmes", "8 programmes thématiques + 6 transverses",
         "14 contextes de collecte avec des indicateurs, des sources et des fréquences distincts"],
        ["Filières", "21 filières agricoles (végétales et animales)",
         "21 chaînes de valeur à suivre, avec des contrats-programmes et des seuils agronomiques spécifiques"],
        ["Territoire", "12 régions du Royaume",
         "12 DRA, 75+ DPA, des profils agro-climatiques différents, des capacités de collecte variables"],
        ["Niveaux de collecte", "Provincial, Régional, Central",
         "3 niveaux de saisie et de validation, avec des droits distincts et des circuits paramétrables par programme"],
        ["Fréquences", "Mensuelle, trimestrielle, annuelle",
         "Des échéanciers différents selon l'indicateur, gérés automatiquement par la plateforme"],
        ["Horizon temporel", "2020-2030 (10 ans)",
         "6 ans de données historiques à migrer, 4 ans de collecte prospective jusqu'à l'horizon 2030"],
    ]
    add_table(doc, perim_headers, perim_rows,
              caption="Tableau 2. Périmètre de suivi de la Stratégie Génération Green et implications pour la plateforme",
              col_widths=[3.0, 4.5, 7.0])

    add_heading(doc, "Les limites du dispositif actuel et les enjeux de la plateforme", 2)

    add_body(doc,
        "Le dispositif de suivi actuel de la SGG repose en grande partie sur des fichiers "
        "tableurs renseignés manuellement à chaque niveau de l'administration, puis transmis "
        "par courrier électronique pour consolidation. Ce mode opératoire, hérité du PMV, "
        "présente plusieurs limites structurelles qui réduisent la valeur du suivi pour "
        "les décideurs."
    )
    add_bullet(doc, [
        "Délais de consolidation : le circuit de remontée manuelle des données provinciales vers le niveau central prend plusieurs semaines, ce qui rend les tableaux de bord toujours décalés par rapport à la réalité du terrain.",
        "Risques d'incohérence : plusieurs versions d'un même tableau de bord peuvent coexister selon l'échelon consulté, sans qu'il soit possible d'identifier la version faisant foi.",
        "Absence de traçabilité : il est difficile de savoir qui a modifié quelle valeur et pour quelle raison, ce qui affaiblit la crédibilité des données pour la reddition de comptes.",
        "Pas de détection automatique des anomalies : un rendement agronomique anormal passe inaperçu si aucun agent ne le signale manuellement.",
        "Impossibilité de croiser filière et territoire automatiquement : les tableaux de bord par filière et par région sont produits manuellement, avec des risques d'erreur élevés sur de gros volumes.",
        "Aucune alerte sur les retards de saisie : les responsables de programme n'ont pas de visibilité sur les DPA qui n'ont pas renseigné leurs données à l'échéance.",
    ])
    add_body(doc,
        "La plateforme demandée dans le cadre de ce marché doit résoudre l'ensemble de ces "
        "limites en fournissant un système intégré, traçable et automatisé, accessible à "
        "chaque niveau de l'administration selon des droits différenciés. Elle doit également "
        "assurer l'interopérabilité avec les systèmes d'information existants au sein du "
        "MAPMDREF, pour éviter toute double saisie et garantir la cohérence des données "
        "entre la plateforme SGG et les autres bases de données sectorielles. Sa mise en "
        "service marquera un changement de mode de pilotage de la stratégie : d'un suivi "
        "rétrospectif et manuel à un pilotage prospectif et outillé, capable d'anticiper "
        "les retards et les écarts avant qu'ils ne deviennent des problèmes irréversibles."
    )


def section_comprehension_mission(doc, gar_path):
    add_heading(doc, "Compréhension de la mission", 1, numbered="2")

    add_body(doc,
        "La présente mission ne se résume pas au développement d'un logiciel de gestion "
        "de données. Son objectif central est de doter la Direction de la Stratégie et des "
        "Statistiques d'un instrument de pilotage stratégique capable de répondre à des "
        "questions précises : quel programme accuse un retard sur ses objectifs 2030 ? "
        "Quelle région remonte des données incomplètes depuis plusieurs cycles ? "
        "Quelle filière dépasse ses cibles de production dans le Souss mais stagne dans "
        "l'Oriental ? La plateforme est l'interface entre les données de terrain et "
        "la décision au niveau central. Sa valeur se mesure à la qualité des décisions "
        "qu'elle rend possible, pas au nombre de fonctionnalités qu'elle contient."
    )
    add_body(doc,
        "Cette mission présente des difficultés spécifiques qui la distinguent d'un projet "
        "informatique standard. La première tient à la nature hétérogène des données à "
        "collecter : des indicateurs agronomiques en tonnes par hectare, des indicateurs "
        "sociaux en nombre de ménages, des indicateurs financiers en millions de dirhams "
        "et des indicateurs institutionnels en nombre de coopératives coexistent dans le "
        "même référentiel, avec des formules de calcul, des sources et des fréquences "
        "différentes. La deuxième difficulté tient à la multiplicité des acteurs : "
        "des agents de terrain dans soixante-quinze provinces, des directeurs régionaux "
        "dans douze DRA et des responsables de programme à la DSS doivent utiliser le "
        "même système avec des interfaces et des droits adaptés à leurs rôles respectifs. "
        "La troisième difficulté est temporelle : la plateforme doit intégrer six ans "
        "de données historiques (2020-2026) et produire des projections comparatives "
        "jusqu'à l'horizon 2030."
    )

    add_heading(doc, "Le cadre analytique : la Gestion Axée sur les Résultats (GAR)", 2)

    add_body(doc,
        "Le cadre analytique retenu pour structurer le référentiel des indicateurs est "
        "la Gestion Axée sur les Résultats (GAR), standard international adopté par les "
        "grandes agences de développement et par le Maroc dans ses politiques publiques. "
        "Ce cadre organise les indicateurs en une chaîne de résultats à quatre niveaux, "
        "chacun correspondant à une question de pilotage différente pour les décideurs."
    )
    add_figure(doc, gar_path,
               "Figure 3. Chaîne de résultats GAR appliquée à la Stratégie Génération Green",
               width_cm=16.0)
    add_body(doc,
        "Les intrants désignent les ressources mobilisées pour mettre en oeuvre les programmes "
        ": budgets FDA, effectifs de conseillers agricoles, superficies équipées en irrigation, "
        "financements accordés aux jeunes exploitants. Leur suivi est nécessaire pour mesurer "
        "l'effort public consenti, mais insuffisant pour évaluer l'efficacité de la stratégie."
    )
    add_body(doc,
        "Les extrants sont les livrables directs et quantifiables des programmes : nombre de "
        "coopératives créées, superficie assurée par programme, nombre de jeunes exploitants "
        "accompagnés, agropoles mis en service. Ils permettent de vérifier que les ressources "
        "mobilisées ont bien produit les réalisations attendues dans les délais contractuels."
    )
    add_body(doc,
        "Les effets mesurent les changements à moyen terme sur les bénéficiaires des programmes "
        ": évolution des revenus agricoles, adoption de pratiques d'irrigation localisée, "
        "taux de regroupement des agriculteurs en coopératives, évolution de l'emploi dans "
        "les services agricoles et paragricoles. Ces indicateurs sont les plus difficiles "
        "à collecter mais les plus informatifs pour les décideurs stratégiques."
    )
    add_body(doc,
        "Les impacts désignent les transformations structurelles à long terme attendues de la "
        "stratégie à l'horizon 2030 : contribution du secteur agricole au PIB national, "
        "réduction de la pauvreté rurale, amélioration de la sécurité alimentaire, "
        "évolution de la balance commerciale agro-alimentaire. Ces indicateurs sont suivis "
        "au niveau national et agrégés à partir des effets mesurés au niveau régional "
        "et provincial. Le tableau de bord d'impact est la vue principale pour les "
        "décideurs de haut niveau."
    )

    add_heading(doc, "Les trois missions et leur logique d'enchaînement", 2)

    add_body(doc,
        "Le CPS structure les prestations en trois missions successives dont l'enchaînement "
        "est contractuellement conditionné : chaque mission ne peut démarrer qu'après "
        "approbation formelle des livrables de la précédente par le maître d'ouvrage, "
        "conformément à l'article 46 du CCAG-EMO. Cette structure protège la DSS contre "
        "les dérives de périmètre en développement : si les spécifications produites en "
        "Mission I ne sont pas satisfaisantes, elles sont corrigées avant que le "
        "développement soit engagé, pas après."
    )

    miss_headers = ["Mission", "Durée", "Objectif", "Condition de démarrage", "Livrables clés"]
    miss_rows = [
        ["Mission I\nConception", "3 mois",
         "Comprendre l'organisation, modéliser les indicateurs, spécifier la plateforme et définir l'architecture",
         "Notification du marché",
         "SFG, SFD, dossiers d'architecture (applicative, technique, sécurité), fiches indicateurs, maquettes UI"],
        ["Mission II\nDéveloppement", "5 mois",
         "Développer les modules fonctionnels, tester la plateforme et produire une version bêta validée",
         "Approbation des livrables Mission I par la DSS",
         "Note technologique, Plan d'intégration des données, version bêta, rapports de tests, recommandations"],
        ["Mission III\nDéploiement", "3 mois",
         "Déployer en production, migrer les données historiques, former les utilisateurs et assurer le support",
         "Validation formelle de la version bêta par la DSS",
         "Plateforme opérationnelle, manuels d'utilisation et d'administration, codes sources, plan de maintenance"],
    ]
    add_table(doc, miss_headers, miss_rows,
              caption="Tableau 3. Les trois missions du marché : objectifs, conditions et livrables",
              col_widths=[2.2, 1.3, 4.5, 3.5, 4.0])

    add_heading(doc, "Les indicateurs de succès de la mission", 2)

    add_body(doc,
        "Au-delà du respect des délais et du contenu des livrables contractuels, le succès "
        "de cette mission se mesurera à plusieurs indicateurs concrets. Le premier est "
        "l'adoption effective de la plateforme par les agents de terrain : une plateforme "
        "qui n'est pas utilisée par les DPA n'a pas rempli son objectif, quelle que soit "
        "sa sophistication technique. Le second est la qualité des données produites : "
        "un taux faible d'anomalies détectées après validation et un taux élevé de "
        "renseignement dans les délais attendus indiquent que le système fonctionne "
        "correctement. Le troisième est l'autonomie des équipes DSS à l'issue de la "
        "période de garantie : la plateforme doit pouvoir être administrée, étendue "
        "et exploitée sans dépendance permanente au titulaire."
    )

    liv_headers = ["Mission", "Livrable", "Format", "Délai", "Exemplaires"]
    liv_rows = [
        ["I", "Compte rendu de l'atelier de démarrage", "Word", "J+3 après atelier", "5 USB + 5 papier"],
        ["I", "Rapport d'audit de l'existant + entretiens", "Word + annexes", "Fin semaine 4", "5 USB + 5 papier"],
        ["I", "Référentiel des indicateurs (fiches SMART+)", "Word structuré", "Fin mois 2", "5 USB + 5 papier"],
        ["I", "SFG validé", "Word", "Fin mois 2", "5 USB + 5 papier"],
        ["I", "SFD + maquettes + dossiers d'architecture", "Word + schémas", "Fin mois 3", "5 USB + 5 papier"],
        ["II", "Note de choix technologique", "Word", "Fin mois 4", "5 USB"],
        ["II", "Plan d'intégration des données (PID)", "Word", "Fin mois 6", "5 USB"],
        ["II", "Note méthodologique sécurité", "Word", "Fin mois 8", "5 USB"],
        ["II", "Version bêta + rapports de tests", "Code + Word", "Fin mois 8", "5 USB"],
        ["III", "Plateforme finale opérationnelle", "Déploiement", "Fin mois 9", "5 USB"],
        ["III", "Manuels d'administration et d'utilisation", "Word", "Fin mois 10", "5 USB + 5 papier"],
        ["III", "Documentation technique + codes sources", "Code + Word", "Fin mois 11", "5 USB + 5 papier"],
        ["III", "Rapport de déploiement + plan de maintenance", "Word", "Fin mois 11", "5 USB + 5 papier"],
    ]
    add_table(doc, liv_headers, liv_rows,
              caption="Tableau 4. Livrables contractuels par mission",
              col_widths=[1.5, 6.0, 2.0, 3.0, 2.0])


def add_bullet(doc, items):
    """Liste à tirets - pour énumérations secondaires ou items courts."""
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.first_line_indent = Cm(-0.5)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        run = p.add_run(f"–  {item}")
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.color.rgb = DARK_RGB
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def add_bullet_dot(doc, items):
    """Liste à points ronds - pour énumérations primaires de 3-6 éléments parallèles."""
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.9)
        p.paragraph_format.first_line_indent = Cm(-0.55)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(3)
        run = p.add_run(f"•  {item}")
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.color.rgb = DARK_RGB
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def add_bullet_square(doc, items):
    """Liste à carrés pleins - pour spécifications techniques, livrables, checklist."""
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.9)
        p.paragraph_format.first_line_indent = Cm(-0.55)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(3)
        run = p.add_run(f"▪  {item}")
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.color.rgb = DARK_RGB
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def section_approche_methodologique(doc):
    add_heading(doc, "Approche méthodologique", 1, numbered="3")

    add_body(doc,
        "La stratégie Génération Green 2020-2030 est un programme de transformation structurelle "
        "du secteur agricole marocain, articulé autour de deux fondements complémentaires et d'un "
        "ensemble de programmes transverses. Son suivi exige un dispositif de mesure rigoureux, "
        "capable de collecter, de consolider et d'analyser des données provenant de douze régions, "
        "de vingt et une filières agricoles et de plusieurs niveaux de l'administration, de la "
        "province jusqu'au niveau central. La Direction de la Stratégie et des Statistiques pilote "
        "ce dispositif et a besoin d'un outil numérique qui traduit cette complexité en tableaux "
        "de bord lisibles, en alertes actionnables et en rapports structurés à destination des décideurs."
    )
    add_body(doc,
        "L'approche proposée repose sur trois principes fondamentaux. Le premier est la co-construction "
        ": aucune spécification ne sera rédigée sans validation préalable par les équipes DSS, les "
        "directeurs régionaux et les responsables de programme. Le second est la progressivité "
        ": chaque mission produit des livrables contrôlables avant engagement de la suivante, "
        "conformément à l'article 4 du CPS. Le troisième est l'ancrage terrain : les entretiens "
        "conduits auprès des DPA et des DRA permettent de concevoir une plateforme qui répond aux "
        "contraintes réelles de collecte au niveau provincial, pas seulement aux exigences centrales."
    )
    add_body(doc,
        "Les prestations s'organisent en trois missions successives d'une durée totale de onze mois. "
        "La Mission I (trois mois) couvre la conception complète : analyse des besoins, inventaire "
        "des indicateurs, spécifications fonctionnelles et architecture technique. La Mission II "
        "(cinq mois) développe la plateforme module par module, en cycles bimensuels avec "
        "démonstrations régulières au maître d'ouvrage. La Mission III (trois mois) déploie la "
        "plateforme en production, forme les utilisateurs et assure le support post-déploiement "
        "pendant toute la période de garantie."
    )

    # ===================================================================
    # MISSION I
    # ===================================================================
    add_heading(doc,
        "Mission I : Conception de la plateforme de suivi-évaluation de la stratégie Génération Green",
        2)

    add_body(doc,
        "La Mission I est la mission fondatrice. Toute décision prise en développement et en "
        "déploiement découlera directement de ce qui aura été conçu, validé et documenté ici. "
        "Son enjeu central n'est pas de rédiger des spécifications : c'est de comprendre "
        "suffisamment l'organisation de la DSS, les pratiques de collecte des DRA et DPA, "
        "et la logique des indicateurs SGG pour que le système conçu réponde aux vrais besoins "
        "opérationnels, pas à une interprétation abstraite du CPS."
    )
    add_body(doc,
        "Cette mission se déroule sur trois mois. Elle produit les Spécifications Fonctionnelles "
        "Générales (SFG), les Spécifications Fonctionnelles et Techniques Détaillées (SFD), et "
        "les dossiers d'architecture applicative, technique et sécurité. Une maquette des interfaces "
        "clés est co-construite avec les équipes DSS pour valider l'adéquation du système avant "
        "tout engagement de développement. Ces livrables sont soumis à approbation du maître "
        "d'ouvrage avant l'engagement de la Mission II."
    )

    # -------------------------------------------------------------------
    # 1. Analyse des besoins
    # -------------------------------------------------------------------
    add_heading(doc, "Analyse des besoins", 3)

    add_body(doc,
        "L'analyse des besoins ne consiste pas à traduire le CPS en liste de fonctionnalités. "
        "Son objet est de comprendre comment la DSS et ses directions régionales et provinciales "
        "produisent, transmettent et utilisent réellement les données de suivi de la SGG. "
        "Sans cette compréhension préalable, toute spécification fonctionnelle reste abstraite "
        "et expose le projet à des écarts coûteux entre ce qui a été conçu et ce qui sera utilisé."
    )

    add_body(doc,
        "La DSS pilote vingt et un programmes regroupés en deux fondements : Génération et Green. "
        "Ces programmes mobilisent douze Directions Régionales de l'Agriculture et plus de "
        "soixante-dix Directions Provinciales, chacune avec ses pratiques de collecte, ses outils "
        "bureautiques, ses délais de remontée et ses contraintes de connectivité. "
        "Un agent de DPA en zone de montagne ne renseigne pas ses indicateurs dans les mêmes "
        "conditions qu'un responsable de programme au siège de la DSS à Rabat. "
        "La plateforme doit fonctionner pour les deux, sans imposer une charge supplémentaire "
        "à ceux qui disposent du moins de moyens."
    )

    add_body(doc,
        "L'analyse s'organise autour de deux activités complémentaires. "
        "La première porte sur l'optimisation de l'approche méthodologique : elle consiste à "
        "cartographier les flux d'information existants, identifier les indicateurs déjà renseignés "
        "et ceux qui manquent, et qualifier le niveau de confiance des données disponibles. "
        "La seconde porte sur la mise en place d'un protocole de suivi et d'évaluation : "
        "elle traduit cette cartographie en spécifications formelles, en arbitrant entre ce que "
        "le CPS demande, ce que les équipes terrain peuvent produire, et ce que les décideurs "
        "doivent être en mesure de lire."
    )

    add_body(doc,
        "Les livrables de cette étape alimentent directement les Spécifications Fonctionnelles "
        "Générales (SFG). Aucune ligne de code ne sera écrite avant leur approbation par le "
        "maître d'ouvrage. Ce séquençage est une condition de réussite : il évite de construire "
        "un système techniquement correct mais inadapté aux besoins opérationnels réels."
    )

    add_heading(doc, "Optimisation de l'approche méthodologique", 3)

    add_body(doc,
        "La première semaine de la mission est consacrée à l'organisation du projet et à "
        "l'appropriation du contexte. Dès la notification du marché, le chef de projet "
        "convoque l'atelier de démarrage et engage simultanément l'audit de l'existant. "
        "Ces deux activités se complètent : l'atelier recueille les attentes et les contraintes "
        "des parties prenantes, l'audit recense les données et les systèmes déjà en place."
    )

    add_body(doc,
        "L'atelier de démarrage se tient sur deux jours dans les locaux de la DSS à Rabat. "
        "Il réunit les équipes centrales de la DSS, des représentants des Directions Régionales "
        "de l'Agriculture (DRA), des Directions Provinciales de l'Agriculture (DPA) et des "
        "entités sous tutelle directement concernées par la stratégie Génération Green. "
        "La présence de profils régionaux et provinciaux dès le premier jour est délibérée : "
        "les contraintes de connectivité, de capacité des agents et de disponibilité des "
        "données varient fortement entre une DPA de plaine irriguée et une DPA de zone "
        "montagnarde. Ignorer ces différences en phase de conception conduit à développer "
        "un système inadapté aux conditions réelles de collecte."
    )

    # Tableau agenda atelier
    agenda_headers = ["Jour", "Demi-journée", "Contenu"]
    agenda_rows = [
        ["Jour 1", "Matin",
         "Présentation des enjeux et objectifs du marché. Exposé de l'approche méthodologique proposée. "
         "Cartographie des programmes SGG par fondement et identification des responsables de programme présents."],
        ["Jour 1", "Après-midi",
         "Travaux en groupes thématiques : Groupe 1 (Fondement Génération), Groupe 2 (Fondement Green), "
         "Groupe 3 (Programmes transverses). Chaque groupe recense les indicateurs actuellement suivis, "
         "les sources disponibles et les lacunes de données."],
        ["Jour 2", "Matin",
         "Restitution des travaux de groupe. Identification des synergies et complémentarités avec les "
         "systèmes d'information existants au niveau du MAPMDREF. Discussion sur les flux d'information "
         "entre niveaux provincial, régional et central."],
        ["Jour 2", "Après-midi",
         "Validation de l'approche méthodologique enrichie. Définition du calendrier détaillé de la "
         "Mission I. Identification des interlocuteurs référents par programme et par région. "
         "Fixation du calendrier des entretiens individuels et des ateliers de validation à venir."],
    ]
    add_table(doc, agenda_headers, agenda_rows,
              caption="Tableau 3. Agenda de l'atelier de démarrage (2 jours)",
              col_widths=[1.5, 2.5, 10.5])

    add_body(doc,
        "À l'issue de l'atelier, un compte rendu détaillé est rédigé et transmis au maître "
        "d'ouvrage dans les trois jours ouvrables. Ce document consigne les consensus atteints, "
        "les points de divergence à trancher, et les décisions prises sur le périmètre de la "
        "plateforme. Il sert de base contractuelle pour les entretiens individuels qui suivent."
    )

    add_body(doc,
        "L'audit de l'existant est conduit en parallèle par l'Architecte des systèmes "
        "d'information. Il couvre trois dimensions. La première est le recensement des outils "
        "de suivi actuels : fichiers Excel de la DSS, bases de données provinciales, états "
        "de renseignement existants, tableaux de bord manuels. La deuxième est l'identification "
        "des données historiques disponibles pour la période 2020-2026, programme par programme. "
        "La troisième est l'évaluation des systèmes d'information en place au sein du MAPMDREF "
        "qui alimentent ou pourraient alimenter la plateforme. Cet audit conditionne directement "
        "la conception du module d'interopérabilité et le protocole de migration des données historiques."
    )

    add_body(doc,
        "Les entretiens individuels structurés complètent l'atelier de démarrage. Chaque profil "
        "est rencontré séparément selon une grille adaptée à ses responsabilités. L'objectif "
        "n'est pas de recueillir une liste d'exigences génériques, mais de comprendre comment "
        "chaque acteur produit, valide et utilise les données de suivi dans son contexte quotidien."
    )

    # Tableau grille d'entretiens
    ent_headers = ["Profil interviewé", "Thèmes couverts", "Durée"]
    ent_rows = [
        ["Responsables de programme DSS (par fondement)",
         "Indicateurs prioritaires par programme. Fréquences de collecte souhaitées versus réelles. "
         "Sources de données utilisées actuellement. Besoins de visualisation pour le pilotage central. "
         "Points de blocage dans les remontées actuelles.",
         "90 min"],
        ["Directeurs régionaux DRA (échantillon de 4 régions)",
         "Organisation de la collecte régionale. Circuit de consolidation des données provinciales. "
         "Délais de remontée actuels. Attentes sur le circuit de validation et les notifications. "
         "Ressources humaines disponibles pour la saisie régionale.",
         "75 min"],
        ["Directeurs provinciaux DPA (échantillon de 6 provinces)",
         "Modalités concrètes de collecte au niveau provincial. Capacités des agents de saisie. "
         "Contraintes de connectivité et de disponibilité informatique. Indicateurs impossibles "
         "à renseigner dans les délais actuels et pourquoi.",
         "60 min"],
        ["Responsables d'entités sous tutelle et interprofessions",
         "Données filières disponibles. Contrats-programmes en cours et indicateurs contractuels. "
         "Indicateurs de chaîne de valeur non couverts par les remontées administratives. "
         "Format et fréquence de partage des données actuelles.",
         "60 min"],
    ]
    add_table(doc, ent_headers, ent_rows,
              caption="Tableau 4. Grille d'entretiens structurés par profil",
              col_widths=[4.0, 9.0, 1.5])

    add_body(doc,
        "À l'issue des entretiens, des ateliers de prototypage d'interfaces utilisateur sont "
        "organisés avec les équipes DSS. Ces ateliers suivent une méthode participative : "
        "l'équipe projet présente plusieurs variantes de maquettes pour chaque écran critique "
        "et les participants choisissent, annotent et modifient en séance. Les écrans traités "
        "sont au minimum : le formulaire de saisie d'un indicateur avec ses règles de validation, "
        "le tableau de bord de suivi d'un programme par région, la vue de consolidation nationale "
        "par fondement, et l'interface de gestion du circuit de validation. Les maquettes "
        "validées sont intégrées au SFD comme référence contractuelle pour le développement."
    )

    # 1.2
    add_heading(doc, "Mise en place d'un protocole de suivi et d'évaluation", 3)

    add_body(doc,
        "La mise en place du protocole de suivi est l'activité la plus exigeante de la Mission I. "
        "Elle consiste à transformer la liste indicative d'indicateurs du CPS en un référentiel "
        "opérationnel complet : chaque indicateur est défini, sourcé, formulé mathématiquement, "
        "rattaché à un programme et à un niveau de collecte, et associé à une valeur de "
        "référence et à une cible 2030. Ce travail se fait en étroite collaboration avec les "
        "responsables de programme DSS et les experts en suivi-évaluation de l'équipe."
    )
    add_body(doc,
        "La méthode SMART+ guide la sélection et la validation de chaque indicateur. Pour être "
        "retenu dans le référentiel, un indicateur doit satisfaire les six critères suivants :"
    )
    add_bullet(doc, [
        "Spécifique : il mesure un aspect précis et délimité du programme concerné, sans ambiguïté d'interprétation entre deux collecteurs différents.",
        "Mesurable : les données nécessaires à son calcul existent ou peuvent être produites par les DPA et DRA dans des délais raisonnables.",
        "Attribuable : l'évolution de sa valeur peut être reliée aux actions de la stratégie Génération Green, pas seulement à des facteurs extérieurs.",
        "Réaliste : sa collecte est faisable au vu des capacités humaines et techniques des agents de terrain, sans alourdir indûment la charge administrative.",
        "Temporel : une fréquence de collecte est fixée (mensuelle, trimestrielle ou annuelle) et respectée de façon cohérente entre régions.",
        "Vérifiable (+) : une source tierce identifiable peut confirmer ou infirmer la valeur déclarée, garantissant la fiabilité du suivi.",
    ])
    add_body(doc,
        "Les indicateurs qui ne satisfont pas ces critères ne sont pas supprimés du référentiel : "
        "ils sont classés comme indicateurs en développement, avec une note de limitation "
        "transmise au maître d'ouvrage et une proposition d'indicateur de substitution. "
        "Cette transparence évite de créer de fausses certitudes dans les tableaux de bord."
    )
    add_body(doc,
        "Le cadre GAR (Gestion Axée sur les Résultats) structure la classification des indicateurs "
        "en quatre niveaux de la chaîne de résultats. Les intrants sont les ressources mobilisées "
        "pour mettre en oeuvre les programmes (budgets FDA, nombres de conseillers agricoles, "
        "superficies aménagées). Les extrants sont les livrables directs des programmes (nombre "
        "de coopératives créées, superficies assurées, jeunes exploitants accompagnés). Les effets "
        "sont les changements à moyen terme sur les bénéficiaires (évolution des revenus agricoles, "
        "taux de regroupement des agriculteurs, adoption de pratiques d'irrigation localisée). "
        "Les impacts sont les transformations structurelles à long terme sur le secteur (contribution "
        "du secteur agricole au PIB, évolution de la balance commerciale agro-alimentaire, réduction "
        "de la pauvreté rurale). Cette classification est intégrée dans le référentiel de la "
        "plateforme et structure les tableaux de bord par niveau de décision."
    )

    # Tableau fiche indicateur
    fiche_headers = ["Champ de la fiche", "Description et utilisation dans la plateforme"]
    fiche_rows = [
        ["Code indicateur",
         "Identifiant unique dans le référentiel. Format : [FONDEMENT]-[PROG]-[NUM]. Ex. GEN-CMA-001."],
        ["Intitulé",
         "Libellé court (50 caractères max) affiché dans les interfaces et les rapports exportés."],
        ["Définition",
         "Description complète de ce que mesure l'indicateur, rédigée pour éviter toute ambiguïté entre deux collecteurs différents."],
        ["Démarche de calcul",
         "Formule ou algorithme précis, avec unités, arrondis et modalités de traitement des valeurs manquantes."],
        ["Source de données",
         "Organisme producteur de la donnée brute, document source et modalité d'accès pour les agents de saisie."],
        ["Fréquence de collecte",
         "Mensuelle, trimestrielle ou annuelle. Détermine l'échéancier de saisie généré automatiquement."],
        ["Niveau de collecte",
         "Provincial (DPA), Régional (DRA) ou Central (DSS). Conditionne les droits de saisie dans la plateforme."],
        ["Valeur de référence",
         "Valeur constatée à l'année de référence (situation de départ), renseignée lors de la migration des données 2020-2026."],
        ["Cible 2030",
         "Objectif fixé dans le contrat-programme ou la stratégie SGG. Affiché sur les graphiques d'évolution."],
        ["Niveau de désagrégation",
         "National, Régional, Provincial, Filière ou combinaison. Détermine les axes de ventilation dans les tableaux de bord."],
        ["Fondement et programme",
         "Rattachement dans la structure des deux fondements SGG et des programmes transverses."],
        ["Niveau GAR",
         "Intrant, Extrant, Effet ou Impact. Permet le filtrage par chaîne de résultats dans les rapports."],
    ]
    add_table(doc, fiche_headers, fiche_rows,
              caption="Tableau 5. Structure type d'une fiche indicateur standardisée",
              col_widths=[4.0, 10.5])

    add_body(doc,
        "La matrice filière x territoire est la colonne vertébrale du référentiel de saisie. "
        "Elle croise les 21 filières agricoles de la stratégie Génération Green avec les "
        "12 régions du Royaume. Les filières couvertes sont : arboricole, arganier, céréalière, "
        "oléagineuse, oléicole, maraîchère de primeurs, palmier dattier, rose à parfum, "
        "rizicole, safran, semencière, sucrière, apicole, avicole, cameline, laitière et "
        "viandes rouges, auxquelles s'ajoutent les filières animales transversales. Cette "
        "structure garantit que chaque donnée collectée au niveau d'une DPA est automatiquement "
        "agrégée vers le bon échelon régional et national, par filière et par programme, "
        "sans traitement manuel intermédiaire."
    )
    add_body(doc,
        "À l'issue de cette étape, l'équipe dispose d'un référentiel complet couvrant "
        "l'ensemble des programmes SGG, avec une estimation de 150 à 200 indicateurs validés "
        "selon la méthode SMART+. La structure et le format du reporting de suivi des "
        "réalisations, des résultats et des impacts par programme sont définis et approuvés "
        "par le maître d'ouvrage avant engagement de la conception technique."
    )

    # -------------------------------------------------------------------
    # 2. Conception de la plateforme
    # -------------------------------------------------------------------
    add_heading(doc, "Conception de la plateforme", 3)

    add_body(doc,
        "La conception couvre deux volets distincts : les exigences fonctionnelles générales "
        "communes à tous les modules, et la description détaillée des prestations attendues "
        "pour chaque composant de la plateforme."
    )

    add_heading(doc, "Exigences fonctionnelles générales", 3)

    add_body(doc,
        "La solution proposée est modulaire et évolutive. Elle répond aux quatre exigences "
        "fonctionnelles générales définies dans le CPS :"
    )
    add_bullet(doc, [
        "L'intégration, la consolidation et la migration des bases de données existantes, couvrant les données historiques de la période 2020-2026 des DPA, DRA et DSS.",
        "L'alimentation continue de la plateforme par des données et des documents spécifiques à chaque indicateur, émanant des Directions centrales, des DRA et des entités sous tutelle.",
        "La prise en charge de la diversité des indicateurs et des différentes parties prenantes : Directions centrales, DPA, DRA, organismes sous tutelle et interprofessions.",
        "La génération automatisée de tableaux de bord dynamiques et de reportings de suivi à différentes fréquences et niveaux de désagrégation.",
    ])
    add_body(doc,
        "Les droits d'accès de chaque acteur sont définis au niveau de chaque indicateur. "
        "Un agent DPA a un droit de saisie sur les indicateurs de collecte provinciale pour "
        "son territoire uniquement. Un directeur DRA a un droit de validation et de "
        "visualisation pour sa région. Un responsable de programme DSS a un droit de "
        "validation centrale et de visualisation nationale. Un décideur a un droit de lecture "
        "sur les tableaux de bord agrégés. Cette granularité est implémentée par le mécanisme "
        "d'isolation des données par région présenté dans la section Architecture technique."
    )

    add_heading(doc, "Description détaillée des prestations attendues", 3)

    add_body(doc,
        "Les Spécifications Fonctionnelles Générales (SFG) décrivent l'ensemble des "
        "fonctionnalités attendues sous forme de cas d'usage, sans entrer dans les détails "
        "d'implémentation. Chaque cas d'usage précise l'acteur concerné, le déclencheur, "
        "le flux principal, les flux alternatifs et les règles de gestion associées. "
        "Le SFG est soumis au maître d'ouvrage pour validation formelle avant l'engagement "
        "de la rédaction du SFD. Cette validation intermédiaire protège les deux parties : "
        "elle confirme que l'équipe a compris le périmètre fonctionnel attendu avant d'investir "
        "dans des spécifications détaillées."
    )
    add_body(doc,
        "Les Spécifications Fonctionnelles et Techniques Détaillées (SFD) traduisent chaque "
        "cas d'usage du SFG en spécifications précises et opérationnelles :"
    )
    add_bullet(doc, [
        "Règles de gestion métier : conditions de saisie, règles de validation automatique, formules de calcul des indicateurs dérivés, conditions de déclenchement des alertes.",
        "Contraintes de format : types de données, limites de saisie, formats d'affichage, règles d'arrondi pour chaque indicateur.",
        "Maquettes d'interfaces utilisateur pour les fonctionnalités critiques, co-validées lors des ateliers de prototypage.",
        "Diagrammes UML : cas d'utilisation, classes et séquence, présentés et commentés dans la section Architecture technique.",
        "Règles de sécurité et de contrôle d'accès spécifiques à chaque module.",
    ])
    add_body(doc,
        "Les dossiers d'architecture complètent le SFD. Le dossier d'architecture applicative "
        "décrit l'organisation en microservices, les interfaces entre services et les patterns "
        "de communication synchrone et asynchrone. Le dossier d'architecture technique précise "
        "l'environnement de déploiement, le système de gestion de base de données et la "
        "stratégie de sauvegarde. Le dossier d'architecture sécurité décrit le modèle de "
        "contrôle d'accès, la politique de gestion des secrets et les mécanismes de "
        "journalisation. Ces trois dossiers sont soumis à validation avant engagement de la "
        "Mission II et constituent la référence contractuelle pour toute évolution de périmètre."
    )

    # Tableau livrables Mission I
    liv1_headers = ["Livrable", "Format", "Soumis au MO"]
    liv1_rows = [
        ["Compte rendu de l'atelier de démarrage",
         "Document Word, 5 à 10 pages", "J+3 après l'atelier"],
        ["Rapport d'audit de l'existant",
         "Document Word avec annexes des fichiers inventoriés", "Fin semaine 3"],
        ["Rapports d'entretiens par profil",
         "Fiches synthétiques par acteur interviewé", "Fin semaine 4"],
        ["Référentiel des indicateurs (150 à 200 fiches)",
         "Fichier structuré + fiches indicateurs individuelles", "Fin mois 2"],
        ["Spécifications Fonctionnelles Générales (SFG)",
         "Document structuré par module et cas d'usage", "Fin mois 2"],
        ["Spécifications Fonctionnelles et Techniques Détaillées (SFD)",
         "Document détaillé avec maquettes et diagrammes UML", "Fin mois 3"],
        ["Dossiers d'architecture (applicative, technique, sécurité)",
         "Trois documents distincts, validés par l'Architecte SI", "Fin mois 3"],
    ]
    add_table(doc, liv1_headers, liv1_rows,
              caption="Tableau 6. Livrables de la Mission I et calendrier de soumission",
              col_widths=[6.0, 5.0, 3.5])

    # ===================================================================
    # MISSION II
    # ===================================================================
    add_heading(doc,
        "Mission II : Développement de la plateforme de suivi-évaluation de la stratégie Génération Green",
        2)

    add_body(doc,
        "La Mission II traduit les spécifications validées en Mission I en une plateforme "
        "fonctionnelle, testée et approuvée par le maître d'ouvrage. Elle dure cinq mois "
        "et s'appuie sur une organisation en cycles bimensuels : chaque cycle produit des "
        "fonctionnalités démontrables, validées par les équipes DSS avant engagement du "
        "cycle suivant. Cette discipline protège le maître d'ouvrage contre les dérives de "
        "périmètre et lui donne une visibilité continue sur l'état réel du développement."
    )
    add_body(doc,
        "Le premier mois de la Mission II est consacré à la mise en place de l'environnement "
        "de développement : conteneurisation des composants applicatifs, pipeline d'intégration "
        "continue, déploiement du système de gestion des identités avec les six rôles définis "
        "dans l'architecture, et configuration des trois environnements (développement, staging "
        "et production). À partir du mois cinq, les développements des modules fonctionnels "
        "démarrent dans l'ordre des dépendances : le référentiel en premier, la saisie en "
        "deuxième, la validation en troisième, le reporting en quatrième."
    )

    # --- Interfaces web
    add_heading(doc, "Développement des interfaces web ergonomiques permettant :", 3)

    add_body(doc,
        "Les interfaces de saisie sont le point de contact quotidien des agents DPA avec la "
        "plateforme. Leur ergonomie conditionne directement la qualité des données collectées. "
        "Un formulaire mal conçu génère des erreurs de saisie et des retards de validation ; "
        "un formulaire bien conçu réduit le temps de saisie et améliore la fiabilité des données. "
        "C'est pourquoi les maquettes co-construites en Mission I servent de référence directe "
        "pour le développement des interfaces."
    )
    add_body(doc,
        "Chaque formulaire intègre plusieurs niveaux de contrôle avant soumission. Le premier "
        "niveau vérifie le type et le format de la donnée saisie. Le deuxième niveau applique "
        "des règles de plage agronomique : un rendement de blé tendre déclaré supérieur au "
        "seuil de plausibilité pour la région concernée déclenche un avertissement visible "
        "avant soumission. Le troisième niveau vérifie la cohérence entre indicateurs liés "
        "au sein du même formulaire : si la superficie emblavée déclarée est inférieure à la "
        "superficie récoltée du même programme, l'anomalie est signalée. Le quatrième niveau "
        "calcule automatiquement les indicateurs dérivés à partir des valeurs saisies. "
        "Une aide contextuelle par champ affiche la définition de l'indicateur et sa démarche "
        "de calcul directement dans le formulaire. La sauvegarde automatique en mode brouillon "
        "préserve les données saisies en cas de coupure de connexion."
    )
    add_body(doc,
        "Les circuits de validation sont paramétrables par programme. Un programme peut nécessiter "
        "une validation régionale uniquement, une autre une double validation régionale puis "
        "centrale. Le circuit est configuré dans le référentiel sans modification du code. "
        "Lorsqu'un validateur identifie une anomalie dans les données soumises, il peut formuler "
        "une demande de clarification directement depuis la plateforme, sans recourir à des "
        "échanges par courrier électronique hors système. La demande est notifiée à l'agent "
        "concerné, sa réponse est enregistrée avec son auteur et sa date, et un tableau de bord "
        "des échanges en cours est visible par les validateurs centraux. Ce mécanisme de feedback "
        "bidirectionnel DSS/DRA/DPA élimine les zones grises et assure une traçabilité complète "
        "des corrections apportées aux données."
    )
    add_body(doc,
        "La plateforme respecte les normes d'accessibilité WCAG 2.1 niveau AA. Les interfaces "
        "de saisie sont disponibles en français et en arabe, pour répondre aux besoins des "
        "agents de collecte au niveau provincial dont la langue de travail est l'arabe. "
        "La configuration des bases de données est accessible aux administrateurs DSS depuis "
        "une interface dédiée, sans intervention technique du titulaire."
    )

    # --- Reporting
    add_heading(doc, "Reporting, visualisation et analyse", 3)

    add_body(doc,
        "Les tableaux de bord sont le produit final visible pour les décideurs. Leur conception "
        "répond à un principe simple : chaque niveau hiérarchique doit pouvoir répondre à ses "
        "questions sans demander une extraction de données à un technicien. Un directeur de "
        "programme DSS doit voir en quelques secondes l'avancement de son programme par région, "
        "l'écart entre réalisations et cible 2030, et les indicateurs en retard de renseignement. "
        "Un décideur doit voir la contribution globale de la SGG à ses indicateurs d'impact "
        "national."
    )
    add_body(doc,
        "Les tableaux de bord proposent les axes d'analyse suivants :"
    )
    add_bullet(doc, [
        "Par fondement : Fondement I Génération et Fondement II Green, avec vue consolidée des deux.",
        "Par programme : chaque programme de chaque fondement, avec ses indicateurs de résultat, de performance et d'impact.",
        "Par filière : suivi des indicateurs de chaîne de valeur pour chacune des 21 filières agricoles de la stratégie.",
        "Par région : performance de chaque région sur les objectifs 2030, avec comparatif inter-régional.",
        "Comparatif réalisé/cible : graphiques d'évolution avec la trajectoire cible 2030 superposée aux réalisations annuelles.",
        "Vue temporelle : évolution mensuelle, trimestrielle ou annuelle selon la fréquence de l'indicateur.",
        "Impact global : indicateurs d'impact national (contribution au PIB agricole, emploi, sécurité alimentaire) à destination des décideurs.",
    ])
    add_body(doc,
        "Des interfaces dédiées permettent la génération de rapports d'activité périodiques, "
        "de notes de reporting thématiques et de présentations destinées aux instances de pilotage. "
        "Les exports sont disponibles en format PDF et Excel pour chaque tableau de bord et rapport, "
        "avec mise en page adaptée à une impression directe ou à une insertion dans un document officiel."
    )

    # --- Suivi des échéances
    add_heading(doc, "Suivi des échéances et système d'alerte", 3)

    add_body(doc,
        "La plateforme gère les échéanciers de renseignement de tous les indicateurs du "
        "référentiel. Pour chaque indicateur, la date limite de saisie est calculée "
        "automatiquement à partir de sa fréquence de collecte et de la date de clôture "
        "de la période. Le système génère des notifications automatiques selon la séquence "
        "suivante : un premier rappel à J-5 avant l'échéance, adressé à l'agent de saisie "
        "responsable ; un second rappel le jour de l'échéance si la saisie n'est pas effectuée ; "
        "une alerte d'escalade au responsable hiérarchique si l'indicateur n'est toujours pas "
        "renseigné 48 heures après l'échéance. Cette séquence est paramétrable par programme "
        "et par niveau de collecte."
    )
    add_body(doc,
        "Un tableau de bord des retards est accessible aux responsables de programme DSS "
        "et aux directeurs DRA. Il affiche, par programme et par région, la liste des "
        "indicateurs en retard, le nombre de jours de dépassement et l'identité de "
        "l'agent responsable. Ce tableau de bord transforme la gestion des retards d'une "
        "activité de relance manuelle en un pilotage visuel et proactif."
    )

    # --- Archivage
    add_heading(doc, "Archivage, sauvegarde et traçabilité", 3)

    add_body(doc,
        "La fiabilité d'une plateforme de suivi-évaluation repose autant sur la qualité "
        "des données que sur leur traçabilité dans le temps. La DSS doit pouvoir répondre "
        "à deux questions : quelle était la valeur déclarée d'un indicateur à une date donnée, "
        "et qui l'a modifiée et pourquoi. Ces deux exigences sont satisfaites par le "
        "mécanisme de versioning et de non-répudiation."
    )
    add_body(doc,
        "Chaque version d'un indicateur pour une période donnée est conservée intégralement "
        "dans la base de données. Toute modification d'une valeur déjà soumise génère une "
        "nouvelle version (N+1), associée obligatoirement à un motif de modification "
        "sélectionné dans une liste contrôlée et à un commentaire libre de l'auteur. "
        "Il est impossible de modifier une valeur sans renseigner ce motif. "
        "L'historique complet de toutes les versions est accessible aux validateurs "
        "centraux sous forme d'un journal de modifications horodaté et non modifiable."
    )
    add_body(doc,
        "L'audit log applicatif enregistre toutes les actions réalisées sur la plateforme : "
        "connexions, saisies, modifications, validations, rejets, exports et changements "
        "de configuration. Il est immuable : aucun utilisateur, y compris l'administrateur "
        "système, ne peut supprimer ou modifier une entrée du journal. Les sauvegardes "
        "automatiques quotidiennes de la base de données et du journal d'audit sont "
        "stockées de façon sécurisée et vérifiées de façon hebdomadaire par un script "
        "de contrôle d'intégrité."
    )

    # --- Interopérabilité
    add_heading(doc, "Interopérabilité et intégration des données", 3)

    add_body(doc,
        "La plateforme SGG ne fonctionnera pas en silo. Le MAPMDREF dispose déjà de systèmes "
        "d'information produisant des données utiles au suivi de la stratégie Génération Green. "
        "Ignorer ces systèmes imposerait une double saisie aux agents de terrain et "
        "créerait des risques d'incohérence entre les données officielles du ministère "
        "et celles de la plateforme. L'interopérabilité est donc traitée comme une "
        "contrainte de conception, pas comme une option."
    )
    add_body(doc,
        "Un Plan d'Intégration des Données (PID) est élaboré par le titulaire et soumis "
        "au maître d'ouvrage lors de la Mission II. Ce plan précise, pour chaque flux "
        "d'échange identifié :"
    )
    add_bullet(doc, [
        "Les flux d'importation et d'exportation : quelles données circulent, dans quel sens, à quelle fréquence.",
        "Les formats d'échange : formats standards ouverts (JSON, CSV, XML) privilégiés pour garantir l'évolutivité.",
        "Les protocoles d'interconnexion : API REST sécurisées avec authentification par jeton à durée de vie limitée.",
        "Les règles de mise à jour et de synchronisation : gestion des conflits en cas de modification simultanée par deux systèmes.",
        "Les responsabilités des parties prenantes : qui alimente quelle donnée, qui valide les imports automatiques.",
    ])
    add_body(doc,
        "La passerelle applicative sert de point d'entrée unique pour toutes les intégrations "
        "avec les SI tiers. Elle applique les règles d'authentification, de limitation de débit "
        "et de journalisation sur chaque appel entrant ou sortant. Les connecteurs vers les "
        "SI existants du MAPMDREF sont développés de façon modulaire, ce qui permet d'en "
        "ajouter de nouveaux sans modifier le coeur de la plateforme."
    )

    # --- Sécurité
    add_heading(doc, "Sécurité", 3)

    add_body(doc,
        "La sécurité est intégrée au processus de développement dès la première ligne de code, "
        "pas ajoutée lors des tests finaux. Chaque fonctionnalité est développée avec le "
        "principe de moindre privilège : un composant ne reçoit que les droits strictement "
        "nécessaires à son fonctionnement. Le compte de base de données du service de "
        "reporting n'a accès qu'aux vues agrégées du schéma de reporting, jamais aux tables "
        "brutes de saisie."
    )
    add_body(doc,
        "Les pratiques de sécurité appliquées tout au long du développement couvrent :"
    )
    add_bullet(doc, [
        "Validation des entrées à deux niveaux : côté interface avant envoi, côté serveur avant tout traitement, sans jamais faire confiance aux données venant du client.",
        "Requêtes paramétrées vers la base de données, sans concaténation de chaînes, couvrant la catégorie A01 (injections) des dix risques les plus critiques identifiés par l'OWASP.",
        "Audit des dépendances tierces à chaque intégration de code : toute dépendance présentant une vulnérabilité connue de score de gravité élevé bloque automatiquement l'intégration.",
        "Gestion des secrets par variables d'environnement injectées au démarrage, jamais stockés dans le code source ni dans les images de conteneurs.",
        "Tests de sécurité en environnement de préproduction avant toute mise en production, couvrant les dix catégories de risque de l'OWASP et les exigences du standard ASVS.",
    ])
    add_body(doc,
        "La phase de test en version bêta est conduite sur un échantillon représentatif "
        "d'indicateurs et avec la participation de quelques Directions centrales et "
        "régionales volontaires. Elle vérifie l'adéquation des fonctionnalités avec "
        "les besoins opérationnels dans des conditions proches du réel, identifie les "
        "anomalies techniques ou fonctionnelles restantes et apporte les ajustements "
        "nécessaires avant déploiement généralisé. Le rapport de test bêta, incluant "
        "les résultats, les anomalies identifiées, les corrections appliquées et les "
        "recommandations pour la mise en production, est remis au maître d'ouvrage. "
        "Sa validation formelle conditionne le passage à la Mission III."
    )

    mod_headers = ["Mois", "Module", "Fonctionnalités principales développées"]
    mod_rows = [
        ["M4", "Environnement et infrastructure",
         "Pipeline CI/CD, conteneurisation, gestion des identités, six rôles, environnements développement/staging/production"],
        ["M4-M5", "Gestion des référentiels",
         "Indicateurs (fiches complètes), programmes, filières (21), régions (12), niveaux de collecte, règles de validation par indicateur"],
        ["M5-M6", "Saisie intelligente et circuit de validation",
         "Formulaires multi-niveaux avec validation de plage et cohérence, brouillon automatique, circuit paramétrable DPA-DRA-DSS, feedback bidirectionnel"],
        ["M6-M7", "Tableaux de bord et reporting",
         "Vues par fondement/programme/filière/région, comparatif réalisé/cible 2030, axe temporel, exports PDF et Excel, interface bilingue"],
        ["M7", "Alertes, notifications et échéanciers",
         "Alertes J-5/J0/dépassement, escalade paramétrable, tableau de bord des retards par programme et par région"],
        ["M7-M8", "Archivage, traçabilité et versioning",
         "Versioning avec motif obligatoire, audit log immuable, archivage sécurisé des pièces justificatives, sauvegardes automatiques"],
        ["M7-M8", "Interopérabilité",
         "Passerelle applicative, PID, connecteurs SI MAPMDREF, import fichiers Excel historiques 2020-2026"],
        ["M8", "Tests bêta et ajustements",
         "Tests de charge, tests de sécurité OWASP, tests fonctionnels avec utilisateurs DSS et DRA, corrections et rapport final"],
        ["M4-M8", "Gestion des utilisateurs et des droits",
         "Six rôles, isolation régionale des données par territoire, délégation de droits, WCAG 2.1 AA, interface bilingue"],
    ]
    add_table(doc, mod_headers, mod_rows,
              caption="Tableau 7. Planification des modules de développement, Mission II",
              col_widths=[1.3, 4.2, 9.0])

    # ===================================================================
    # MISSION III
    # ===================================================================
    add_heading(doc,
        "Mission III : Déploiement de la plateforme de suivi-évaluation de la stratégie Génération Green",
        2)

    add_body(doc,
        "La Mission III met la plateforme entre les mains de ses utilisateurs réels. "
        "Elle dure trois mois et couvre deux activités distinctes : le déploiement "
        "général de la plateforme en environnement de production, et le support technique "
        "post-déploiement pendant la période de garantie de douze mois. Ces deux activités "
        "sont indissociables : un bon déploiement réduit les incidents de support, "
        "et un bon support consolide la confiance des utilisateurs dans la plateforme."
    )

    # --- 1- Déploiement
    add_heading(doc, "Déploiement de la plateforme", 3)

    add_body(doc,
        "Le déploiement en production est précédé d'une phase de préparation rigoureuse. "
        "La leçon des projets informatiques publics qui échouent au déploiement est "
        "presque toujours la même : des fonctionnalités développées sans confrontation "
        "suffisante avec les conditions réelles d'utilisation, des utilisateurs formés "
        "trop tard, et des données historiques non migrées qui obligent à maintenir "
        "deux systèmes en parallèle pendant des mois. Cette mission est conçue pour "
        "éviter ces trois écueils."
    )

    add_body(doc,
        "Avant tout déploiement en production, une checklist go/no-go formelle est "
        "validée conjointement par l'équipe projet et le maître d'ouvrage. Aucun "
        "déploiement n'est déclenché sans signature écrite sur chacun des points suivants :"
    )
    add_bullet(doc, [
        "Tests de charge validés : la plateforme supporte le nombre d'utilisateurs simultanés et le volume de données prévisibles pour les 3 premières années d'utilisation, avec des temps de réponse conformes aux seuils définis avec le maître d'ouvrage.",
        "Conformité sécurité vérifiée en préproduction : tests OWASP Top 10 et ASVS réalisés, aucune vulnérabilité critique ou élevée non corrigée.",
        "Plan de sauvegarde et de reprise opérationnel : sauvegardes automatiques testées, procédure de restauration vérifiée sur une copie de la base de production.",
        "Administrateurs DSS formés : au moins deux agents DSS maîtrisent l'administration complète de la plateforme avant le déploiement.",
        "Plan de migration des données historiques validé et testé sur un échantillon représentatif.",
        "Plan de rollback documenté et testé en environnement de préproduction.",
    ])

    add_body(doc,
        "La stratégie de déploiement est Blue-Green. Deux environnements de production "
        "identiques sont maintenus. La nouvelle version est déployée sur le second "
        "environnement pendant que le premier reste actif. Le basculement du trafic "
        "vers le nouvel environnement s'effectue en quelques secondes, sans interruption "
        "de service pour les utilisateurs. En cas d'anomalie critique constatée "
        "après basculement, la procédure de rollback ramène l'ensemble du trafic "
        "vers l'environnement précédent en moins de cinq minutes, sans perte de données. "
        "Cette procédure est documentée, testée en préproduction et remise au maître "
        "d'ouvrage sous forme d'un manuel d'urgence opérationnel."
    )

    add_body(doc,
        "La migration des données historiques couvre la période 2020-2026, soit "
        "six années de données de suivi accumulées dans les fichiers Excel de la DSS "
        "et les bases provinciales. Le protocole de migration se déroule en trois étapes. "
        "La première est l'extraction et la transformation : les fichiers sources sont "
        "lus par le module d'import, les données sont normalisées selon le référentiel "
        "de la plateforme et un rapport d'anomalies liste les valeurs hors plage, "
        "les indicateurs non trouvés dans le référentiel et les doublons détectés. "
        "La deuxième est la validation : le rapport d'anomalies est soumis au maître "
        "d'ouvrage pour arbitrage avant tout chargement. La troisième est le chargement "
        "itératif, programme par programme et région par région, avec vérification "
        "des totaux après chaque lot. À l'issue de la migration, les données historiques "
        "sont accessibles dans les tableaux de bord d'évolution dès le premier jour d'utilisation."
    )

    # Tableau formation
    form_headers = ["Profil", "Contenu de la formation", "Durée", "Format"]
    form_rows = [
        ["Administrateurs système DSS (2 à 3 agents)",
         "Configuration de la plateforme. Gestion des référentiels (indicateurs, programmes, filières, régions). Paramétrage des circuits de validation. Gestion des comptes utilisateurs et des droits d'accès. Procédures de sauvegarde et de restauration. Mise à jour des composants.",
         "3 jours", "Présentiel, exercices pratiques sur environnement de formation"],
        ["Validateurs centraux DSS (5 à 10 agents)",
         "Circuit de validation complet. Gestion des anomalies de saisie. Formulation de demandes de clarification. Consultation de l'historique des versions. Génération des rapports et exports. Tableau de bord des échanges en cours.",
         "2 jours", "Présentiel, mise en situation avec données réelles"],
        ["Directeurs et agents de validation DRA (par région)",
         "Validation régionale. Consultation des données provinciales. Alertes et échéanciers. Tableau de bord régional. Communication avec le niveau central via la plateforme.",
         "1 jour", "Présentiel en région ou visioconférence selon la région"],
        ["Agents de saisie DPA (par province)",
         "Saisie des formulaires par programme et par filière. Règles de validation automatique et interprétation des avertissements. Sauvegarde en brouillon et reprise. Correction des rejets de validation. Import de fichiers CSV.",
         "1 jour", "Présentiel en province, support par guide utilisateur illustré"],
        ["Décideurs DSS et directions (profil lecture)",
         "Navigation dans les tableaux de bord. Interprétation des indicateurs et des graphiques d'évolution. Comparatif réalisé/cible 2030. Génération et export des rapports de pilotage.",
         "2 heures", "Démonstration guidée, remise d'un guide de référence rapide"],
    ]
    add_table(doc, form_headers, form_rows,
              caption="Tableau 8. Programme de formation par profil utilisateur, Mission III",
              col_widths=[3.5, 7.5, 1.5, 2.0])

    add_body(doc,
        "Les tests de recette fonctionnelle sont conduits par les équipes DSS assistées "
        "de l'équipe projet, sur la base du plan de recette préparé en fin de Mission II. "
        "Chaque scénario de test est exécuté par un agent DSS dans des conditions réelles "
        "de saisie. L'équipe projet observe et consigne les anomalies. Chaque anomalie "
        "identifiée est tracée, classée par criticité, corrigée et revalidée avant "
        "l'établissement du procès-verbal de réception provisoire. Ce PV, signé par le "
        "maître d'ouvrage, déclenche officiellement la période de garantie de douze mois."
    )

    # --- 2- Support technique
    add_heading(doc, "Support technique", 3)

    add_body(doc,
        "L'accompagnement post-déploiement vise deux objectifs distincts. Le premier "
        "est de maintenir la plateforme en état de fonctionnement optimal face à "
        "l'évolution des volumes de données, du nombre d'utilisateurs actifs et des "
        "règles de gestion des programmes. Le second est de permettre à l'équipe DSS "
        "de monter progressivement en autonomie sur l'administration et l'exploitation "
        "de la plateforme, conformément à l'obligation de transfert de compétences "
        "prévue à l'article 8 du CPS."
    )
    add_body(doc,
        "Le support technique comprend deux catégories d'interventions. Les interventions "
        "correctives traitent les anomalies constatées en production : diagnostic, correction, "
        "test de non-régression et déploiement du correctif. Les interventions préventives "
        "comprennent les contrôles périodiques du système, les mises à jour des composants "
        "logiciels pour couvrir les nouvelles vulnérabilités publiées, et les audits de "
        "performance et de capacité pour anticiper les besoins avant qu'ils ne deviennent "
        "des incidents."
    )

    sla_headers = ["Niveau", "Définition", "Prise en charge", "Résolution"]
    sla_rows = [
        ["Criticité 1 : Bloquante",
         "Plateforme inaccessible ou module essentiel (saisie, validation) hors service pour l'ensemble des utilisateurs. Aucune procédure de contournement.",
         "2 heures ouvrées",
         "8 heures ouvrées"],
        ["Criticité 2 : Dégradée",
         "Fonctionnalité importante défaillante pour un sous-ensemble d'utilisateurs. Une procédure de contournement temporaire est possible.",
         "4 heures ouvrées",
         "3 jours ouvrés"],
        ["Criticité 3 : Mineure",
         "Anomalie d'affichage, incohérence de données non bloquante pour le suivi, ou demande d'évolution mineure du référentiel.",
         "2 jours ouvrés",
         "10 jours ouvrés"],
    ]
    add_table(doc, sla_headers, sla_rows,
              caption="Tableau 9. Niveaux de service (SLA), période de garantie de 12 mois",
              col_widths=[3.2, 7.3, 2.5, 1.5])

    add_body(doc,
        "Durant toute la période de garantie, toute anomalie est corrigée sans coût "
        "supplémentaire pour le maître d'ouvrage. Un rapport mensuel de suivi du support "
        "technique est transmis à la DSS. Ce rapport comprend : le récapitulatif des "
        "tickets ouverts et fermés dans le mois, le détail des corrections appliquées, "
        "les mises à jour de composants réalisées, l'état des indicateurs de performance "
        "système (disponibilité, temps de réponse, espace disque) et les recommandations "
        "pour l'optimisation ou l'extension de la plateforme."
    )
    add_body(doc,
        "Le programme de transfert de compétences est formalisé dans un plan remis au "
        "maître d'ouvrage dès le démarrage de la Mission III. Il couvre quatre domaines : "
        "l'administration système (configuration, sauvegardes, mises à jour courantes), "
        "la gestion des référentiels (ajout ou modification d'indicateurs, de filières "
        "ou de programmes), le paramétrage des circuits de validation (modification des "
        "règles sans intervention du titulaire) et l'exploitation des tableaux de bord "
        "(création de nouvelles vues, modification des filtres et des axes d'analyse). "
        "À l'issue de la période de garantie, les équipes DSS désignées sont autonomes "
        "sur ces quatre domaines, sans dépendance permanente au titulaire."
    )

    # ------------------------------------------------------------------
    add_heading(doc, "Pratiques d'ingénierie et philosophie Agile/DevOps", 2)

    add_heading(doc, "Développement Agile et cycles bimensuels", 3)
    add_body(doc,
        "La Mission II est conduite en cycles bimensuels. Chaque cycle comprend quatre moments: "
        "la revue du backlog avec les équipes DSS, le développement des fonctionnalités prioritaires, "
        "la revue de code interne par l'Architecte SI, et la démonstration au maître d'ouvrage. "
        "Cette cadence donne au maître d'ouvrage une visibilité continue sur l'avancement, "
        "pas seulement à la livraison de la version bêta."
    )
    add_body(doc,
        "Le backlog est priorisé par valeur métier. Les modules qui débloquent les autres sont "
        "développés en premier: la gestion des référentiels avant la saisie, la saisie avant la "
        "validation, la validation avant les tableaux de bord. Les demandes d'évolution identifiées "
        "en cours de cycle sont tracées dans le backlog et traitées dans le cycle suivant, jamais "
        "en cours de cycle. Cette discipline protège la stabilité des livraisons et la confiance "
        "du maître d'ouvrage dans les délais."
    )
    add_body(doc,
        "Les cérémonies sont courtes et orientées décision. Le stand-up quotidien de l'équipe "
        "dure quinze minutes. La revue de sprint avec le maître d'ouvrage dure deux heures "
        "maximum et produit une liste explicite de validations et de points ouverts. "
        "La rétrospective en fin de cycle identifie un seul point d'amélioration, appliqué "
        "dans le cycle suivant. Cette discipline agile évite les réunions longues qui "
        "consomment le temps de développement sans produire de décisions."
    )

    add_heading(doc, "Qualité logicielle: SOLID, KISS, DRY et design patterns", 3)
    add_body(doc,
        "Chaque microservice applique le principe de responsabilité unique (S de SOLID): "
        "l'API principale gère la logique métier, le service notification gère uniquement la "
        "diffusion des alertes, le service audit consigne uniquement les actions. Ce découplage "
        "rend chaque service testable de façon isolée et remplaçable sans impact sur les autres "
        "composants."
    )
    add_body(doc,
        "Le circuit de validation paramétrable est implémenté par le pattern Strategy: la règle "
        "de validation d'un programme (nombre de niveaux requis, délai maximum, profils autorisés) "
        "est un objet configurable, pas du code conditionnel dupliqué. Ajouter un programme avec "
        "des règles différentes ne nécessite pas de toucher au code existant. C'est l'application "
        "directe du principe Open/Closed (O de SOLID): ouvert à l'extension, fermé à la "
        "modification."
    )
    add_body(doc,
        "Le principe KISS gouverne les choix d'implémentation: la solution la plus simple qui "
        "répond au besoin est retenue, pas la plus sophistiquée. Le principe DRY élimine la "
        "duplication: les règles de calcul d'un indicateur sont définies une fois dans le "
        "référentiel et lues par tous les modules qui en ont besoin, jamais recopiées. "
        "Le clean code se traduit par un nommage qui reflète le vocabulaire métier: "
        "`indicateur`, `filiere`, `niveauValidation`, jamais `data`, `item` ou `flag`."
    )
    add_body(doc,
        "Le Repository pattern sépare l'accès aux données de la logique métier dans chaque "
        "service: un service ne sait pas si ses données viennent de la base de données ou d'un "
        "cache. Ce découplage facilite les tests unitaires (on injecte un dépôt de test "
        "sans base de données réelle) et simplifie les évolutions futures de schéma. "
        "Le pattern CQRS sépare les opérations de lecture (tableaux de bord, exports) des "
        "opérations d'écriture (saisie, validation), ce qui permet d'optimiser chaque chemin "
        "indépendamment."
    )

    add_heading(doc, "Pipeline CI/CD et intégration continue", 3)
    add_body(doc,
        "Chaque push sur la branche de développement déclenche automatiquement: la compilation, "
        "les tests unitaires, les tests d'intégration contre une base de données de test, et une "
        "analyse des dépendances contre les CVE connues (OWASP Dependency Check). Un push qui "
        "échoue un test ne peut pas être intégré dans la branche principale. Cette règle protège "
        "la stabilité de la version bêta contre les régressions."
    )
    add_body(doc,
        "Trois environnements sont maintenus en parallèle. L'environnement de développement est "
        "local à chaque développeur. L'environnement de staging est déployé automatiquement à "
        "chaque merge sur la branche principale: c'est l'environnement des démonstrations "
        "bimensuelles au maître d'ouvrage. L'environnement de production est déployé manuellement "
        "après validation formelle de la version bêta. Les environnements staging et production "
        "partagent la même configuration de déploiement, ce qui garantit que ce qui fonctionne "
        "en staging fonctionne en production."
    )
    add_body(doc,
        "Toute modification du code passe par la revue de l'Architecte SI avant intégration. "
        "La revue vérifie trois points: la conformité au SFD (la fonctionnalité répond-elle "
        "exactement à la spécification?), les implications de sécurité (les entrées sont-elles "
        "validées, les accès contrôlés au niveau approprié?), et la lisibilité (un développeur "
        "absent peut-il comprendre ce code en cinq minutes?). Les modules du circuit de validation "
        "et de la gestion des droits ont une couverture de tests automatisés de 80% minimum."
    )

    add_heading(doc, "Sécurité par les pratiques de développement", 3)
    add_body(doc,
        "La sécurité est intégrée au processus de développement, pas ajoutée en fin de Mission II. "
        "Chaque fonctionnalité est développée avec le principe de moindre privilège: un composant "
        "ne reçoit que les droits dont il a besoin pour fonctionner. Le compte applicatif du "
        "service de reporting ne peut lire que les vues agrégées; il ne peut pas écrire dans "
        "les tables brutes ni accéder aux données de gestion des utilisateurs."
    )
    add_body(doc,
        "Les entrées utilisateur sont validées à deux niveaux. L'interface web valide le format "
        "avant envoi. Le service principal re-valide côté serveur avant tout traitement, sans "
        "jamais faire confiance aux données venant du client. Les requêtes vers la base de données "
        "sont construites avec des paramètres liés, jamais par concaténation de chaînes. Ces règles "
        "couvrent les catégories A1 (injections) et A3 (exposition de données) de l'OWASP Top 10 "
        "par construction, pas par test après coup."
    )
    add_body(doc,
        "Les dépendances tierces sont auditées à chaque build via OWASP Dependency Check. "
        "Une dépendance avec une CVE de score CVSS supérieur à 7 bloque le pipeline. "
        "Les secrets (mots de passe, clés d'authentification, credentials de base de données) "
        "ne sont jamais stockés dans le code source ni dans les images de conteneurs: ils sont "
        "injectés via les variables d'environnement au démarrage."
    )

    # ------------------------------------------------------------------
    add_heading(doc, "Communication et collaboration interdisciplinaire", 2)
    add_body(doc,
        "La difficulté principale de ce projet n'est pas technique: c'est la traduction fidèle "
        "des indicateurs agronomiques en spécifications logicielles. Un expert S&E qui décrit "
        "un indicateur de taux d'adoption de l'irrigation localisée et un développeur qui doit "
        "le stocker et l'afficher parlent deux référentiels distincts. Sans mécanisme explicite "
        "pour combler cet écart, les ambiguïtés se transforment en corrections coûteuses "
        "pendant la Mission II."
    )
    add_body(doc,
        "Les ateliers de la Mission I réunissent systématiquement un expert S&E et le développeur "
        "responsable du module concerné. Le développeur reformule la règle de calcul sous forme "
        "de pseudocode, que l'expert valide ou corrige immédiatement. Ce mécanisme élimine les "
        "malentendus avant l'écriture du code, pas après. Le SFD résultant est co-signé par les "
        "deux profils, ce qui engage les deux parties sur la même compréhension."
    )
    add_body(doc,
        "Les démonstrations bimensuelles pendant la Mission II suivent un format adapté aux "
        "profils présents. Les décideurs DSS voient les indicateurs et les tableaux de bord. "
        "Les techniciens voient le circuit de validation et les interfaces de saisie. "
        "Les experts S&E valident que les règles de calcul s'affichent correctement. "
        "Ce format mixte garantit que personne ne valide une fonctionnalité sans la comprendre."
    )
    add_body(doc,
        "L'expert Genre et Inclusion intervient en Mission I pour valider que le référentiel "
        "d'indicateurs intègre les désagrégations par genre requises pour le suivi des programmes "
        "d'inclusion agricole. Elle intervient en Mission III pour adapter les supports de "
        "formation aux agents de saisie en milieu rural. Cette intervention n'est pas symbolique: "
        "la collecte d'indicateurs genrés exige que les agents comprennent pourquoi ces données "
        "sont demandées et comment les renseigner sans ambiguïté."
    )


def section_architecture(doc, arch_path, org_path, usecase_path, classes_path, sequence_path):
    add_heading(doc, "Architecture technique proposée", 1, numbered="4")

    add_body(doc,
        "La plateforme est conçue autour de quatre microservices aux responsabilités séparées et "
        "d'un gateway centralisé. L'API principale gère la logique métier: indicateurs, référentiels, "
        "circuit de validation. Le service audit enregistre chaque action dans un journal immuable. "
        "Le service notification diffuse les alertes aux utilisateurs concernés. Le service rapport "
        "génère les exports à la demande. Ce découplage présente un avantage concret pour le "
        "maître d'ouvrage: chaque service peut être mis à jour ou redémarré indépendamment, sans "
        "interruption des autres composants."
    )

    add_figure(doc, arch_path,
               "Figure 4. Architecture microservices de la plateforme SGG",
               width_cm=15.0)

    add_body(doc,
        "Les trois vues UML ci-après complètent le diagramme d'architecture globale: la vue "
        "fonctionnelle (cas d'utilisation), la vue structurelle (classes des entités métier), "
        "et la vue comportementale (séquence du circuit de validation). Ces diagrammes ont été "
        "produits en Mission I comme livrables du SFD et constituent la référence contractuelle "
        "pour le développement en Mission II."
    )

    # ------------------------------------------------------------------
    add_heading(doc, "Périmètre fonctionnel - Diagramme de cas d'utilisation", 2)

    add_body(doc,
        "La plateforme expose onze cas d'utilisation répartis entre cinq profils d'acteurs. "
        "Ces profils correspondent exactement aux six rôles Keycloak définis dans le realm "
        "sgg-platform, le rôle LECTEUR partageant les cas d'utilisation du DÉCIDEUR. "
        "Chaque association acteur-cas d'utilisation correspond à une règle de contrôle d'accès "
        "vérifiée à deux niveaux: par le Spring Cloud Gateway au niveau applicatif, "
        "et par le Row-Level Security de PostgreSQL au niveau base de données."
    )
    add_figure(doc, usecase_path,
               "Figure 5. Diagramme de cas d'utilisation - Plateforme SGG",
               width_cm=15.5)
    add_body(doc,
        "Le SAISISSEUR couvre les fonctions de collecte: saisie manuelle, import CSV et soumission "
        "pour validation. Il n'a pas accès aux tableaux de bord ni aux données d'autres régions. "
        "Les validateurs RÉGIONAL et CENTRAL partagent le cas d'utilisation de validation, mais "
        "leur portée sur les données est distincte par Row-Level Security. L'ADMIN est le seul "
        "profil habilité à modifier le référentiel des indicateurs, à paramétrer le circuit de "
        "validation, et à consulter le journal d'audit: ces trois opérations sont critiques "
        "pour l'intégrité du dispositif de suivi."
    )

    # ------------------------------------------------------------------
    add_heading(doc, "Structure des données - Diagramme de classes", 2)

    add_body(doc,
        "Le diagramme de classes représente les huit entités métier principales de la plateforme "
        "et leurs relations. La conception s'appuie sur le principe de séparation des données de "
        "référence (Indicateur, Programme, Filière, Région) et des données de collecte "
        "(SaisieDonnée, ValidationCircuit). Cette séparation permet de modifier le référentiel "
        "sans affecter les saisies existantes, et de versionner les indicateurs indépendamment "
        "des données collectées."
    )
    add_figure(doc, classes_path,
               "Figure 6. Diagramme de classes - Entités métier principales",
               width_cm=15.0)
    add_body(doc,
        "SaisieDonnée est l'entité centrale du modèle. Elle est liée à un Indicateur du "
        "référentiel, à une Région (ce lien active le Row-Level Security), et porte un "
        "ValidationCircuit qui trace le niveau de validation atteint et le statut courant "
        "(BROUILLON, EN_ATTENTE, VALIDÉ, REJETÉ). Toute transition de statut génère "
        "automatiquement une EntréeAudit dans MongoDB via l'événement RabbitMQ correspondant. "
        "L'entité Alerte est produite par le service notification lorsqu'une saisie dépasse "
        "son délai de validation sans traitement."
    )
    add_body(doc,
        "Les attributs `ancienneValeur` et `nouvelleValeur` de EntréeAudit sont stockés en JSON "
        "pour conserver la flexibilité face aux évolutions du référentiel. Ce choix permet "
        "d'archiver l'état complet d'une saisie à chaque transition sans schéma d'audit fixe. "
        "Le champ `horodatage` est un UTC timestamp indexé: les requêtes d'audit par plage de "
        "dates sont exécutées sans scan complet de la collection."
    )

    # ------------------------------------------------------------------
    add_heading(doc, "Circuit de validation - Diagramme de séquence", 2)

    add_body(doc,
        "Le diagramme de séquence détaille le flux complet d'une saisie d'indicateur, de la "
        "création initiale par l'agent provincial jusqu'à l'enregistrement dans le journal d'audit. "
        "Ce flux implique sept composants et onze messages. Sa lisibilité est essentielle lors "
        "des tests de recette: chaque étape du diagramme correspond à un cas de test vérifiable "
        "par les équipes DSS."
    )
    add_figure(doc, sequence_path,
               "Figure 7. Diagramme de séquence - Circuit de validation d'une saisie d'indicateur",
               width_cm=16.0)
    add_body(doc,
        "Le flux se décompose en deux phases. La phase de collecte (messages 1 à 2) crée la "
        "saisie en base avec le statut BROUILLON: le SAISISSEUR peut modifier ses données "
        "sans déclencher le circuit. La phase de soumission (messages 3 à 11) commence par "
        "une mise à jour du statut en EN_ATTENTE, suivie de la publication d'un événement "
        "data.submitted sur RabbitMQ. Ce message est consommé par deux services en parallèle: "
        "le service notification qui alerte le validateur concerné, et le service audit qui "
        "enregistre la soumission de façon immuable."
    )
    add_body(doc,
        "Le découplage asynchrone via RabbitMQ est le choix architectural central de ce flux. "
        "Si le service notification est temporairement indisponible, la saisie reste validable: "
        "les messages en attente sont traités dès la reprise du service, sans perte et sans "
        "intervention manuelle. La validation ou le rejet par le VALIDATEUR (message 8) suit "
        "le même schéma: mise à jour du statut, publication de data.validated ou "
        "data.rejected, consommation par le service audit pour clôturer l'entrée du journal."
    )

    # ------------------------------------------------------------------
    add_heading(doc, "Justification des choix techniques", 2)

    add_heading(doc, "Communication synchrone et asynchrone", 3)
    add_body(doc,
        "Les opérations synchrones (lecture et modification de données métier) transitent par le "
        "Spring Cloud Gateway, qui valide le token Keycloak avant chaque requête. Les opérations "
        "asynchrones (audit, notifications, alertes d'échéancier) transitent par RabbitMQ. Ce "
        "découplage garantit l'exécution différée des opérations secondaires même en cas de "
        "surcharge temporaire, sans bloquer la réponse à l'utilisateur."
    )

    add_heading(doc, "Sécurité et isolation des données", 3)
    add_body(doc,
        "Keycloak gère l'authentification et l'autorisation pour l'ensemble de la plateforme selon "
        "six rôles couvrant les profils décrits dans le CPS: ADMIN, VALIDATEUR_CENTRAL, "
        "VALIDATEUR_REGIONAL, SAISISSEUR, LECTEUR et DECIDEUR. Le Row-Level Security de PostgreSQL "
        "ajoute une protection au niveau de la base de données: un agent de saisie de la région "
        "de Souss-Massa ne peut pas accéder aux données de la région de l'Oriental, même avec un "
        "token valide. Cette garantie est vérifiable lors des tests de recette."
    )

    add_heading(doc, "Traçabilité et conformité DGSSI", 3)
    add_body(doc,
        "Le service audit consigne chaque action dans MongoDB en mode append-only: aucune donnée "
        "du journal ne peut être modifiée après écriture. Chaque entrée porte l'horodatage, "
        "l'identifiant de l'utilisateur, son rôle Keycloak, et la donnée avant et après "
        "modification. Cette conception répond aux exigences de traçabilité du CPS et satisfait "
        "aux recommandations DGSSI en matière d'audit des accès aux systèmes d'information publics."
    )

    add_heading(doc, "Propriété totale du code par le maître d'ouvrage", 3)
    add_body(doc,
        "L'architecture utilise exclusivement des technologies open source: Kotlin avec Spring Boot, "
        "PostgreSQL, MongoDB, React, Keycloak, RabbitMQ, et Garage pour le stockage objet. "
        "Aucune dépendance envers un éditeur propriétaire, aucune licence récurrente. "
        "Conformément à l'article correspondant du CPS, le maître d'ouvrage reçoit l'intégralité "
        "du code source à la réception provisoire et peut le maintenir sans recours au titulaire."
    )

    add_heading(doc, "Infrastructure, déploiement et exploitation", 2)

    add_heading(doc, "Conteneurisation Docker et gestion des environnements", 3)
    add_body(doc,
        "La plateforme est packagée en sept images Docker. Chaque service déclare ses dépendances "
        "via les health checks Docker: l'API principale ne démarre que si PostgreSQL répond sur "
        "son port de santé, le service notification ne démarre que si RabbitMQ est disponible. "
        "Cette configuration rend le démarrage déterministe et les pannes explicites: un service "
        "en état 'unhealthy' indique précisément quel composant vérifier, sans diagnostic manuel."
    )
    add_body(doc,
        "Docker Compose orchestre l'ensemble sur un seul serveur. Ce choix répond aux contraintes "
        "réelles du déploiement: le maître d'ouvrage dispose d'un serveur de production, pas d'un "
        "cluster. La complexité d'orchestration doit correspondre aux capacités de maintenance "
        "de la DSS. Un fichier Docker Compose de deux cents lignes est maintenable par un "
        "administrateur de niveau intermédiaire; un cluster Kubernetes exige des compétences "
        "spécialisées que peu d'administrations disposent en interne."
    )
    add_body(doc,
        "Les volumes Docker sont montés sur des chemins dédiés du serveur hôte, séparés des "
        "images applicatives. Une mise à jour de l'application ne touche pas aux données. "
        "Une migration échouée peut être annulée en quelques minutes en revenant à l'image "
        "précédente. Les migrations de schéma PostgreSQL sont gérées par Flyway: chaque "
        "modification de schéma est un script SQL versionné et idempotent, appliqué "
        "automatiquement au démarrage du service."
    )

    add_heading(doc, "Nginx: reverse proxy et sécurité applicative", 3)
    add_body(doc,
        "Nginx remplit trois fonctions dans l'architecture. En tant que reverse proxy, il route "
        "les requêtes vers le bon service selon le chemin de l'URL. En tant que terminaison TLS, "
        "il chiffre toutes les communications entre les navigateurs et la plateforme, satisfaisant "
        "l'exigence de chiffrement en transit de la DGSSI. En tant que protecteur applicatif, "
        "il applique un rate limiting par IP qui bloque les tentatives d'attaque par force brute "
        "sur l'authentification Keycloak."
    )
    add_body(doc,
        "La configuration Nginx inclut les en-têtes de sécurité HTTP recommandés par OWASP: "
        "X-Frame-Options (protection contre le clickjacking), X-Content-Type-Options (blocage "
        "du MIME sniffing), et Content-Security-Policy (restriction des sources JavaScript "
        "autorisées). Ces en-têtes couvrent directement les catégories A5 (mauvaise configuration "
        "de sécurité) et A6 (composants vulnérables) de l'OWASP Top 10. La configuration "
        "complète est livrée dans un fichier versionné avec le code source et documentée "
        "pour l'administrateur DSS."
    )

    add_heading(doc, "Gestion des données et sauvegardes automatisées", 3)
    add_body(doc,
        "Les sauvegardes de PostgreSQL sont déclenchées chaque nuit par un script cron qui "
        "exécute pg_dump avec compression. Les sauvegardes de MongoDB utilisent mongodump. "
        "Les fichiers produits sont chiffrés avant transfert vers le stockage objet de la DSS. "
        "La rétention automatique couvre trente jours glissants. Un script de restauration "
        "documenté et testé sur l'environnement de staging permet la reprise après sinistre "
        "en moins de quatre heures."
    )
    add_body(doc,
        "Les migrations de schéma PostgreSQL sont versionnées avec Flyway: chaque modification "
        "est un script SQL numéroté séquentiellement. Flyway applique les migrations dans l'ordre "
        "au démarrage du service, sans intervention manuelle. Staging et production ont toujours "
        "le même schéma, ce qui élimine la principale cause d'incidents lors des déploiements "
        "en production."
    )

    add_heading(doc, "Automatisation opérationnelle", 3)
    add_body(doc,
        "L'équipe livre avec le code source un ensemble de scripts bash couvrant: le déploiement "
        "initial de l'infrastructure, les mises à jour des images Docker, le déclenchement manuel "
        "des sauvegardes, la rotation des secrets Keycloak, et la vérification de l'état de "
        "santé de tous les composants. Chaque script expose un mode dry-run qui affiche les "
        "actions qu'il prendrait sans les exécuter, permettant à l'administrateur DSS de "
        "comprendre l'impact avant de lancer une opération."
    )
    add_body(doc,
        "La documentation opérationnelle livrée avec le code source explique chaque script, "
        "ses paramètres, et les cas d'erreur fréquents. Un administrateur système de niveau "
        "intermédiaire peut effectuer une mise à jour, une restauration, ou une rotation de "
        "secrets sans connaissance approfondie des composants sous-jacents. Cette autonomie "
        "du maître d'ouvrage après la réception provisoire est une condition explicite du CPS."
    )

    add_heading(doc, "Tableau des composants et justifications", 2)

    headers_tech = ["Composant", "Technologie retenue", "Justification principale"]
    rows_tech = [
        ["API Principale", "Kotlin / Spring Boot",
         "Null safety compilateur, data classes pour DTOs, coroutines pour validation asynchrone"],
        ["Service Audit", "Kotlin / Spring Boot + MongoDB",
         "MongoDB en append-only: journal immuable sans risque de modification accidentelle"],
        ["Service Notification", "Kotlin / Spring Boot",
         "Cohérence de l'écosystème JVM, réutilisation des bibliothèques de sécurité Spring"],
        ["Service Rapport", "Python / FastAPI",
         "Écosystème Python pour la génération de graphiques et exports (pandas, openpyxl, ReportLab)"],
        ["API Gateway", "Spring Cloud Gateway",
         "Intégration native avec Keycloak, filtrage par rôle, rate limiting"],
        ["IAM", "Keycloak",
         "RBAC éprouvé pour les SI publics, compatible DGSSI, open source"],
        ["Base de données principale", "PostgreSQL + TimescaleDB",
         "Row-Level Security natif, TimescaleDB pour les séries temporelles d'indicateurs"],
        ["Messagerie", "RabbitMQ",
         "Adapté au volume de la plateforme, plus léger que Kafka pour ce périmètre"],
        ["Stockage documents", "Garage (API S3)",
         "Open source Apache 2.0, S3-compatible, remplace MinIO après son archivage en avril 2026"],
        ["Frontend", "React + Vite",
         "Plateforme entièrement authentifiée: pas besoin de SSR. Vite avec toolchain Rust."],
    ]
    add_table(doc, headers_tech, rows_tech,
              caption="Tableau 4. Composants techniques et justifications",
              col_widths=[3.5, 4.0, 7.0])


def section_equipe(doc, org_path):
    add_heading(doc, "Équipe proposée", 1, numbered="5")

    add_body(doc,
        "L'équipe proposée est composée de six profils complémentaires couvrant la totalité "
        "des compétences requises par le CPS : direction de projet, suivi-évaluation de "
        "politiques agricoles, intelligence économique et analytique, architecture des "
        "systèmes d'information, et développement logiciel full-stack. Chaque expert est "
        "positionné sur une ou plusieurs missions avec un périmètre d'intervention clairement "
        "délimité. La coordination transversale est assurée par le chef de projet sur les "
        "onze mois contractuels."
    )

    add_heading(doc, "Organisation et gouvernance de la mission", 2)

    add_body(doc,
        "La gouvernance repose sur trois niveaux d'intervention articulés entre eux. "
        "Le comité de pilotage, co-présidé par le directeur de la DSS et le chef de projet, "
        "se réunit à la fin de chaque mission pour valider les livrables et autoriser le "
        "démarrage de la phase suivante. C'est la seule instance habilitée à émettre un "
        "accord officiel de passage à la mission suivante. Le comité technique, composé de "
        "l'expert suivi-évaluation, de l'expert BI et de l'architecte SI, se réunit toutes "
        "les deux semaines pour arbitrer les questions fonctionnelles avant qu'elles ne "
        "bloquent le développement."
    )
    add_body(doc,
        "Un rapport mensuel écrit est transmis au MO avec l'avancement par livrable, "
        "la consommation de charge par expert, et les risques actifs. Un espace de suivi "
        "partagé est mis à disposition du référent DSS en lecture permanente."
    )

    add_figure(doc, org_path,
               "Figure 8. Organigramme de l'équipe - Mission AOOI N°01/2026/DSS",
               width_cm=14.0)

    add_heading(doc, "Profils détaillés des experts", 2)

    add_heading(doc, "Chef de projet : Brahim EL ORF", 3)
    add_body(doc,
        "Brahim EL ORF assure la direction de la mission sur les onze mois contractuels. "
        "Son parcours couvre la conduite de projets informatiques pour des administrations "
        "publiques marocaines, avec une maîtrise des procédures de marchés publics et "
        "des exigences de reporting propres aux maîtres d'ouvrage institutionnels. "
        "Il est l'interlocuteur principal du maître d'ouvrage pour toutes les décisions "
        "engageant le calendrier, le périmètre ou les livrables de la mission."
    )
    add_body(doc,
        "Son rôle couvre quatre responsabilités distinctes : la planification et le suivi "
        "de l'avancement par rapport aux jalons contractuels, la gestion des risques et "
        "des aléas, la coordination entre les pôles thématique et technique, et la "
        "production des livrables de gouvernance remis au MO à chaque fin de mission."
    )
    add_bullet_square(doc, [
        "Pilotage de l'ensemble des livrables sur les trois missions (Mois 1 à 11)",
        "Animation du comité de pilotage et du comité technique bimensuel",
        "Interlocuteur unique du maître d'ouvrage pour les décisions engageant le contrat",
    ])

    add_heading(doc, "Expert en suivi-évaluation : Hassan KAMIL", 3)
    add_body(doc,
        "Hassan KAMIL porte la dimension méthodologique de la Mission I. "
        "Son expérience couvre le suivi-évaluation des politiques agricoles marocaines "
        "avec une maîtrise du cadre GAR et une connaissance directe des pratiques de "
        "collecte des DPA et DRA. Sa compréhension des contraintes opérationnelles "
        "des directions provinciales est déterminante pour la conception du circuit "
        "de validation : il sait quels indicateurs sont réellement renseignables à "
        "l'échelle provinciale et dans quels délais."
    )
    add_body(doc,
        "Il est responsable de la sélection et de la validation des indicateurs selon "
        "la méthode SMART+, de la rédaction des fiches indicateurs standardisées, "
        "et de la cohérence entre le référentiel fonctionnel et les pratiques de terrain "
        "des soixante-quinze directions provinciales concernées."
    )
    add_bullet_square(doc, [
        "Responsable du référentiel des indicateurs SGG (150 à 200 indicateurs SMART+)",
        "Co-rédacteur des Spécifications Fonctionnelles Générales (SFG)",
        "Validation du circuit de collecte provincial/régional/central",
    ])

    add_heading(doc, "Expert BI / Data Analytics : Hassan EL BAHI", 3)
    add_body(doc,
        "Hassan EL BAHI intervient en Mission I pour la conception des tableaux de bord "
        "analytiques et en Mission II pour valider leur implémentation. Sa valeur dans "
        "cette mission tient à son positionnement en amont : en intégrant la logique "
        "analytique dès la phase de spécification, il évite les aller-retours habituels "
        "entre ce que les spécifications décrivent et ce que la réalité technique permet "
        "d'afficher. Un tableau de bord conçu sans référence aux données disponibles "
        "produit des écrans vides ou des indicateurs non renseignés."
    )
    add_body(doc,
        "Il est responsable de la conception des axes d'analyse par fondement, par "
        "programme et par territoire, de la logique de visualisation pour les trois "
        "niveaux décisionnels (DPA, DRA, DSS centrale), et de la validation des vues "
        "PostgreSQL qui alimentent le module de reporting."
    )
    add_bullet_square(doc, [
        "Conception des tableaux de bord pour les trois niveaux hiérarchiques (DPA/DRA/DSS)",
        "Modélisation du schéma de reporting sur vues PostgreSQL read-only",
        "Validation des indicateurs visuels avec les équipes DSS en Mission I",
    ])

    add_heading(doc, "Architecte des systèmes d'information : Abdelouahab AZIZ", 3)
    add_body(doc,
        "Abdelouahab AZIZ est le référent technique de la mission. Il conçoit "
        "l'architecture en Mission I et supervise le développement en Mission II, "
        "avec une présence couvrant les cinq mois de développement actif. Cette "
        "continuité lui permet de valider chaque composant avant intégration et "
        "de maintenir la cohérence entre l'architecture documentée et ce qui est "
        "effectivement construit."
    )
    add_body(doc,
        "Il est responsable du choix et de la justification de chaque composant "
        "technique, de la note d'architecture et des diagrammes UML remis en fin "
        "de Mission I, de la supervision du pipeline CI/CD, et de la stratégie "
        "de déploiement Blue-Green en Mission III. Sa responsabilité s'étend à "
        "la conformité sécurité de l'ensemble de la plateforme vis-à-vis des "
        "exigences DGSSI."
    )
    add_bullet_square(doc, [
        "Architecture microservices : 4 services indépendants, API Gateway, Row-Level Security",
        "Responsable de la note d'architecture et des diagrammes UML (livrables Mission I)",
        "Supervision du pipeline CI/CD avec OWASP Dependency Check et tests ASVS en staging",
        "Stratégie Blue-Green et procédure de rollback documentée pour la mise en production",
    ])

    add_heading(doc, "Développeurs Full-stack : Mohammed OUAZZANE et Noureddine AMENZOU", 3)
    add_body(doc,
        "Mohammed OUAZZANE et Noureddine AMENZOU interviennent à partir du quatrième mois "
        "et couvrent l'ensemble des cinq mois de développement actif en Mission II. "
        "Leurs compétences couvrent les deux couches de la plateforme : le backend "
        "Kotlin avec Spring Boot pour la logique métier et les APIs, et le frontend "
        "React pour les interfaces utilisateur. Le travail est réparti par module "
        "fonctionnel selon un plan d'affectation défini en début de Mission II par "
        "l'architecte SI."
    )
    add_body(doc,
        "Chaque développeur est responsable de ses modules de bout en bout : "
        "développement, tests unitaires et d'intégration, documentation technique, "
        "et participation à la recette. Cette organisation évite les zones grises "
        "sur la propriété du code et facilite la maintenance après livraison."
    )
    add_bullet_square(doc, [
        "Stack technique : Kotlin/Spring Boot (backend), React/Vite (frontend), PostgreSQL, RabbitMQ",
        "Développement par cycles bimensuels avec démonstrations au MO à chaque fin de cycle",
        "Participation à la recette technique et au transfert de compétences en Mission III",
    ])

    add_heading(doc, "Tableau synthétique de l'équipe", 2)

    headers_eq = ["Expert", "Rôle", "Missions", "Jours"]
    rows_eq = [
        ["Brahim EL ORF",
         "Chef de projet",
         "I + II + III",
         "110j"],
        ["Hassan KAMIL",
         "Expert en suivi-évaluation",
         "I (principal) + III (formation)",
         "90j"],
        ["Hassan EL BAHI",
         "Expert BI / Data Analytics",
         "I (conception) + II (validation)",
         "75j"],
        ["Abdelouahab AZIZ",
         "Architecte des systèmes d'information",
         "I + II + III",
         "128j"],
        ["Mohammed OUAZZANE",
         "Développeur Full-stack 1",
         "II + III",
         "128j"],
        ["Noureddine AMENZOU",
         "Développeur Full-stack 2",
         "II + III",
         "128j"],
    ]
    add_table(doc, headers_eq, rows_eq,
              caption="Tableau 5. Équipe proposée - synthèse des interventions",
              col_widths=[4.0, 4.5, 4.5, 1.5])

    add_heading(doc, "Conformité aux exigences du CPS", 2)

    add_body(doc,
        "Le tableau ci-après met en regard chaque profil exigé par le CPS avec l'expert "
        "proposé. Cette lecture croisée permet au comité d'évaluation de vérifier la "
        "conformité point par point."
    )

    headers_conf = ["Profil exigé par le CPS", "Expert proposé", "Justification"]
    rows_conf = [
        ["Chef de projet",
         "Brahim EL ORF",
         "Direction de projets SI pour administrations publiques marocaines. "
         "Pilotage de l'ensemble des livrables sur les 11 mois contractuels."],
        ["Expert en suivi-évaluation",
         "Hassan KAMIL",
         "Maîtrise du cadre GAR et des pratiques de collecte DPA/DRA. "
         "Responsable du référentiel des indicateurs SMART+ (150 à 200 indicateurs)."],
        ["Expert BI / Data Analytics",
         "Hassan EL BAHI",
         "Conception des tableaux de bord pour les trois niveaux décisionnels. "
         "Modélisation du schéma de reporting et validation des vues PostgreSQL."],
        ["Architecte des systèmes d'information",
         "Abdelouahab AZIZ",
         "Architecture microservices, note d'architecture et diagrammes UML. "
         "Supervision CI/CD, conformité DGSSI et stratégie de déploiement Blue-Green."],
        ["Développeur Full-stack 1",
         "Mohammed OUAZZANE",
         "Développement backend Kotlin/Spring Boot et frontend React/Vite. "
         "Intervention sur les cinq mois de Mission II et recette en Mission III."],
        ["Développeur Full-stack 2",
         "Noureddine AMENZOU",
         "Développement backend Kotlin/Spring Boot et frontend React/Vite. "
         "Intervention sur les cinq mois de Mission II et recette en Mission III."],
    ]
    add_table(doc, headers_conf, rows_conf,
              caption="Tableau 6. Correspondance profils CPS / experts proposés",
              col_widths=[4.0, 3.5, 7.0])


def section_planning(doc, gantt_path):
    add_heading(doc, "Planning d'exécution", 1, numbered="6")

    add_body(doc,
        "Le planning couvre les onze mois d'exécution contractuelle répartis en trois missions "
        "enchaînées. Les missions ne se chevauchent pas : chacune démarre formellement après "
        "la validation écrite du livrable final de la mission précédente par le comité de "
        "pilotage DSS. Ce séquencement garantit que le développement en Mission II s'appuie "
        "sur un référentiel fonctionnel gelé, et que la mise en production en Mission III "
        "intervient sur une plateforme pleinement recettée."
    )

    add_heading(doc, "Déroulement par mission", 2)

    add_heading(doc, "Mission I - Mois 1 à 3", 3)
    add_body(doc,
        "La Mission I ouvre dès la première semaine avec la réunion de lancement officielle "
        "et la constitution du comité de pilotage. L'atelier de démarrage, prévu en semaine "
        "deux et trois, rassemble les représentants des DPA, des DRA et de la DSS centrale "
        "pour partager la vision du dispositif de suivi et recueillir les premières contraintes "
        "terrain. Les semaines suivantes sont consacrées aux entretiens structurés : soixante-quinze "
        "Directions Provinciales, douze Directions Régionales et les unités centrales de la DSS "
        "sont interrogées sur leurs pratiques actuelles de collecte, leurs sources de données "
        "et leurs attentes vis-à-vis de la plateforme."
    )
    add_body(doc,
        "À partir du deuxième mois, la modélisation du référentiel des indicateurs s'appuie "
        "sur les données collectées et sur le cadre GAR. Chaque indicateur est traité selon "
        "la méthode SMART+, validé par les experts S&E et documenté dans une fiche standardisée "
        "avant d'être intégré au SFG. La conception de l'architecture technique démarre en "
        "parallèle dès la neuvième semaine, aboutissant à la note d'architecture et aux "
        "diagrammes UML remis en fin de Mission I avec le SFD."
    )
    add_bullet_square(doc, [
        "Livrable de jalon M3 : SFG + SFD + note d'architecture + 4 diagrammes UML + fiches indicateurs",
        "Condition de déclenchement Mission II : validation écrite du SFD par le comité de pilotage DSS",
    ])

    add_heading(doc, "Mission II - Mois 4 à 8", 3)
    add_body(doc,
        "Le mois quatre est consacré à la mise en place de l'environnement de développement : "
        "pipeline CI/CD, registre d'images Docker, environnement de staging. Les modules sont "
        "développés dans un ordre dicté par les dépendances fonctionnelles. Le module "
        "référentiel et les formulaires de saisie arrivent en premier, car le circuit de "
        "validation en dépend. Les tableaux de bord et le module de reporting suivent une "
        "fois que les données de saisie peuvent circuler dans le système."
    )
    add_body(doc,
        "Les deux derniers mois de la Mission II sont partagés entre la finalisation des "
        "modules d'archivage et de traçabilité, et les campagnes de tests. Les tests "
        "d'intégration vérifient les flux bout-en-bout ; les tests de sécurité OWASP Top 10 "
        "et ASVS sont conduits sur le staging avant toute livraison. Le rapport de tests "
        "constitue un livrable contractuel remis avec la version bêta en fin de Mois 8."
    )

    add_heading(doc, "Mission III - Mois 9 à 11", 3)
    add_body(doc,
        "La Mission III ne démarre pas par le déploiement mais par une checklist go/no-go "
        "formelle et la migration des données historiques 2020-2026. Ce travail préalable "
        "garantit que la plateforme n'est pas vide le jour de la mise en production. "
        "Le déploiement en production suit la stratégie Blue-Green : l'ancienne version "
        "reste disponible le temps de vérifier que la nouvelle est stable, et le rollback "
        "s'exécute en moins de cinq minutes si nécessaire."
    )
    add_body(doc,
        "Les mois dix et onze combinent l'accompagnement post-déploiement, les formations "
        "par profil utilisateur (saisisseurs DPA, validateurs DRA, décideurs, administrateurs) "
        "et les corrections des anomalies relevées lors de la recette. La réception provisoire, "
        "signée en fin de Mois 11, déclenche la période de garantie de douze mois."
    )

    add_figure(doc, gantt_path,
               "Figure 9. Diagramme de Gantt - Planning d'exécution sur 11 mois",
               width_cm=15.5)

    add_heading(doc, "Jalons contractuels et conditions d'engagement", 2)

    headers_jal = ["Jalon", "Échéance", "Livrables associés", "Condition de démarrage suivant"]
    rows_jal = [
        ["Réunion de lancement",
         "Semaine 1",
         "PV de lancement signé",
         "Notification officielle du marché"],
        ["Atelier de démarrage",
         "Semaine 2-3",
         "Compte-rendu atelier, liste participants",
         "GO du comité de pilotage"],
        ["Livraison SFG",
         "Fin Mois 2",
         "SFG validé par DSS",
         "Approbation écrite du MO sous 10 jours"],
        ["Livraison SFD + Note architecture",
         "Fin Mois 3",
         "SFD, diagrammes UML, note architecture",
         "Validation = gel du référentiel, démarrage Mission II"],
        ["Livraison version bêta",
         "Fin Mois 8",
         "Plateforme bêta, rapport tests sécurité, PID",
         "Recette bêta = démarrage Mission III"],
        ["Réception provisoire",
         "Fin Mois 11",
         "Code source, documentation, formations réalisées",
         "Départ période de garantie 12 mois"],
    ]
    add_table(doc, headers_jal, rows_jal,
              caption="Tableau 7. Jalons contractuels et conditions d'enchaînement",
              col_widths=[3.5, 2.0, 4.5, 4.5])

    add_heading(doc, "Points de vigilance sur le chemin critique", 2)

    add_body(doc,
        "Le chemin critique passe par deux transitions : la validation du SFD en fin de Mois 3 "
        "et la recette de la version bêta en fin de Mois 8. Un retard non maîtrisé sur l'un "
        "de ces points se propage directement sur le jalon suivant et sur la date de réception "
        "provisoire. Trois mesures spécifiques encadrent ces transitions."
    )
    add_bullet_dot(doc, [
        "Délai de validation contractualisé à 10 jours pour chaque livrable majeur, avec approbation tacite au 15e jour sans retour écrit du MO",
        "Gel formel du référentiel à la signature du SFD : toute évolution postérieure est traitée en avenant, pas absorbée silencieusement",
        "Checklist go/no-go quatre semaines avant le déploiement en Mission III, avec marge de 5 jours ouvrés intégrée à chaque phase",
    ])


def section_chronogramme(doc):
    add_heading(doc, "Chronogramme d'affectation du personnel", 1, numbered="7")

    add_body(doc,
        "Le tableau ci-après détaille le nombre de jours d'intervention de chaque expert "
        "sur les onze mois de la mission. La charge totale s'élève à 934 jours, répartis "
        "de manière asymétrique selon les phases : la Mission I mobilise principalement "
        "les experts S&E et le chef de projet, la Mission II transfère la charge vers le "
        "pôle technique, et la Mission III répartit l'effort sur l'ensemble de l'équipe "
        "pour la formation et l'accompagnement au démarrage."
    )

    add_heading(doc, "Logique d'affectation par phase", 2)

    add_body(doc,
        "La répartition des charges n'est pas uniforme : elle suit la nature de chaque phase. "
        "En Mission I, les experts S&E et l'architecte SI portent l'essentiel de la charge "
        "(respectivement 51j, 36j et 56j sur trois mois) car la phase de conception est "
        "intensive en ateliers, entretiens et rédaction de spécifications. Les développeurs "
        "n'interviennent pas encore : leur entrée avant la fin du SFD serait sans objet."
    )
    add_body(doc,
        "La Mission II inverse cette répartition. Les développeurs full-stack (98j chacun) "
        "et l'expert BI (84j) portent la production. Les experts S&E restent présents pour "
        "répondre aux questions fonctionnelles qui surgissent inévitablement en cours de "
        "développement, mais leur charge diminue. En Mission III, l'effort se redistribue "
        "sur l'ensemble de l'équipe pour la formation, la recette et l'accompagnement. "
        "Le chef de projet maintient une présence constante à 10 jours par mois sur toute "
        "la durée, ce qui garantit une disponibilité prévisible pour le maître d'ouvrage "
        "quelle que soit la phase."
    )

    months = ["M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "Total"]
    headers_chrono = ["Expert"] + months

    rows_chrono = [
        ["Ahmed BEN HAMMOU\n(Chef de projet)",
         10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 110],
        ["Hassan KAMIL\n(Expert S&E)",
         18, 18, 15, 10, 10, 10,  8,  8,  8,  8,  8, 121],
        ["S.-E. BOUYOUSFI\n(Expert S&E/BI)",
         12, 12, 12, 12, 12, 12, 12, 12, 10, 10,  8, 124],
        ["Expert BI\n(Data Analytics)",
          5,  5,  5, 15, 18, 18, 18, 15, 10, 10,  5, 124],
        ["Architecte SI",
         20, 18, 18, 18, 12, 10, 10, 10,  8,  8,  8, 140],
        ["Développeur Full-stack 1",
          0,  0,  0, 18, 20, 20, 20, 20, 10, 10, 10, 128],
        ["Développeur Full-stack 2",
          0,  0,  0, 18, 20, 20, 20, 20, 10, 10, 10, 128],
        ["Latifa CHBANAT\n(Expert Genre)",
          5,  5,  5,  3,  3,  3,  3,  3,  5,  5,  5,  45],
    ]

    # Convert to string rows for add_table
    str_rows = [[str(v) for v in row] for row in rows_chrono]

    # Custom table for chronogramme
    table = doc.add_table(rows=1 + len(str_rows), cols=len(headers_chrono))
    table.style = 'Table Grid'

    hdr = table.rows[0]
    for i, h in enumerate(headers_chrono):
        cell = hdr.cells[i]
        shade_cell(cell, PRIMARY_HEX)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = 'Calibri'
        r.font.size = Pt(8)
        r.font.bold = True
        r.font.color.rgb = WHITE_RGB

    for ri, row_data in enumerate(str_rows):
        bg = LIGHT_BG_HEX if ri % 2 == 0 else WHITE_HEX
        row = table.rows[ri + 1]
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            # Last column (Total) and last row always highlighted
            if ci == len(row_data) - 1:
                shade_cell(cell, "D5E8D4")
            else:
                shade_cell(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci > 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.name = 'Calibri'
            r.font.size = Pt(8)
            if ci == 0 or ci == len(row_data) - 1:
                r.font.bold = True

    # Column widths
    table.rows[0].cells[0].width = Cm(3.5)
    for row in table.rows:
        for ci in range(1, len(headers_chrono) - 1):
            row.cells[ci].width = Cm(1.0)
        row.cells[-1].width = Cm(1.3)

    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cr = cp.add_run("Tableau 8. Chronogramme d'affectation du personnel (jours/mois)")
    cr.font.name = 'Calibri'
    cr.font.size = Pt(9)
    cr.font.italic = True
    cr.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_heading(doc, "Charge totale par expert et par mission", 2)

    add_body(doc,
        "Le tableau ci-après agrège les jours par mission pour chaque expert. "
        "Cette vue synthétique permet au comité de pilotage de vérifier l'alignement "
        "entre les rôles déclarés dans la section Équipe et la charge réellement affectée "
        "à chaque phase."
    )

    headers_par_mission = ["Expert", "Mission I (M1-M3)", "Mission II (M4-M8)", "Mission III (M9-M11)", "Total"]
    rows_par_mission = [
        ["Ahmed BEN HAMMOU", "30j", "50j", "30j", "110j"],
        ["Hassan KAMIL",     "51j", "46j", "24j", "121j"],
        ["S.-E. BOUYOUSFI",  "36j", "60j", "28j", "124j"],
        ["Expert BI",        "15j", "84j", "25j", "124j"],
        ["Architecte SI",    "56j", "60j", "24j", "140j"],
        ["Développeur FS 1", "0j",  "98j", "30j", "128j"],
        ["Développeur FS 2", "0j",  "98j", "30j", "128j"],
        ["Latifa CHBANAT",   "15j", "15j", "15j",  "45j"],
        ["TOTAL",           "203j","511j","206j","920j"],
    ]
    add_table(doc, headers_par_mission, rows_par_mission,
              caption="Tableau 9. Charge totale par expert et par mission",
              col_widths=[4.0, 3.0, 3.5, 3.5, 2.0])


def section_valeur_ajoutee(doc, cicd_path):
    add_heading(doc, "Engagements qualité et spécificités techniques", 1, numbered="8")

    add_body(doc,
        "Cette section présente les six points sur lesquels l'offre d'ABI CONSULTING dépasse "
        "les exigences minimales du CPS. Chaque point est ancré dans une exigence concrète "
        "du marché et étayé par une référence de mission ou un mécanisme technique vérifiable "
        "lors des tests de recette."
    )

    add_heading(doc, "Architecture open source et propriété totale du MO", 2)
    add_body(doc,
        "Tous les composants retenus sont sous licences open source permissives. A la "
        "réception provisoire, le maître d'ouvrage reçoit l'intégralité du code source, "
        "la documentation technique, et les scripts de déploiement commentés. Aucune "
        "redevance de licence ne sera due après la garantie, et le MO peut migrer la "
        "plateforme vers un autre prestataire sans dépendance contractuelle envers le "
        "titulaire. Chaque script d'administration expose un mode dry-run qui affiche "
        "les actions prévues sans les exécuter, permettant à l'équipe DSS de comprendre "
        "l'opération avant de la lancer. Ce niveau d'autonomie opérationnelle est une "
        "condition explicite de l'Article 8 du CPS sur le transfert de compétences."
    )

    add_heading(doc, "Pipeline CI/CD avec contrôle de sécurité intégré", 2)
    add_body(doc,
        "Chaque commit déclenche automatiquement le pipeline de livraison continue. "
        "Les tests unitaires s'exécutent en premier ; si un test échoue, le build s'arrête. "
        "L'analyse OWASP Dependency Check intervient avant toute construction d'image : "
        "une bibliothèque présentant une CVE critique bloque le pipeline jusqu'à correction. "
        "Le déploiement sur l'environnement de staging n'a lieu qu'après que ces deux "
        "barrières sont franchies. Les tests d'intégration et les tests de sécurité ASVS "
        "sont conduits sur staging avant toute livraison au MO."
    )
    add_figure(doc, cicd_path,
               "Figure 10. Pipeline CI/CD - de la revue de code au déploiement production",
               width_cm=15.5)

    add_heading(doc, "Audit log immuable et traçabilité", 2)
    add_body(doc,
        "Le service audit enregistre chaque opération dans une collection MongoDB en mode "
        "append-only. Aucune entrée ne peut être modifiée ou supprimée, même par un "
        "administrateur système. Chaque entrée porte l'horodatage UTC, l'identifiant de "
        "l'utilisateur, son rôle, et les valeurs avant et après modification. Ce journal "
        "est exportable en JSON et CSV pour inspection lors d'un audit DGSSI ou d'un "
        "contrôle de la Cour des comptes. La non-répudiation est garantie : un agent ne "
        "peut pas contester une action enregistrée dans le journal horodaté."
    )
    add_bullet_square(doc, [
        "Versioning des données : chaque version conservée (N, N+1), motif de modification obligatoire",
        "Rétention configurable : 7 ans par défaut, paramétrable par l'administrateur DSS",
    ])

    add_heading(doc, "Tests de sécurité en environnement de préproduction", 2)
    add_body(doc,
        "La campagne de tests de sécurité est conduite sur l'environnement de staging, "
        "qui réplique la configuration de production à l'identique. Elle couvre les dix "
        "catégories de l'OWASP Top 10, l'OWASP ASVS niveau L2 pour les fonctions "
        "d'authentification et de validation, et des tests d'injection ciblés sur tous "
        "les formulaires et endpoints. Le résultat est un rapport remis au maître d'ouvrage "
        "avant le déploiement en production : aucune vulnérabilité de niveau critique ou "
        "élevé ne reste ouverte. Ce rapport constitue un livrable contractuel, pas un "
        "document de dernière minute."
    )

    add_heading(doc, "Expérience directement comparable : SABIL et DigiTPME", 2)
    add_body(doc,
        "ABI CONSULTING a conçu et déployé deux systèmes de suivi-évaluation numériques "
        "pour des institutions marocaines avant ce marché. SABIL, réalisé pour Expertise "
        "France en 2023, couvre un périmètre multi-niveaux directement comparable à "
        "la présente mission : collecte provincial/régional/central, circuit de validation "
        "paramétrable, tableaux de bord différenciés par rôle, huit régions couvertes. "
        "DigiTPME, réalisé pour GIZ-Maroc en 2022, a nécessité le déploiement dans "
        "douze régions pour 3 500 entreprises. Les deux missions ont confronté l'équipe "
        "aux mêmes difficultés que cette mission posera : référentiel évolutif, migration "
        "de fichiers Excel vers une base structurée, formation d'utilisateurs aux profils "
        "hétérogènes. Les décisions d'architecture proposées sont éprouvées, pas théoriques."
    )

    add_heading(doc, "Isolation régionale par Row-Level Security", 2)
    add_body(doc,
        "La plateforme couvre douze régions avec des agents de saisie et de validation "
        "distincts par territoire. Le Row-Level Security de PostgreSQL garantit que "
        "l'isolation est assurée au niveau de la base de données, en plus du contrôle "
        "applicatif par Keycloak. Ces deux barrières sont indépendantes : même une erreur "
        "de code dans l'API ne peut pas exposer les données d'une région à un utilisateur "
        "d'une autre. Un scénario de test de recette dédié vérifie explicitement ce "
        "cloisonnement lors de la recette fonctionnelle en fin de Mission III."
    )

    add_heading(doc, "Migration des données historiques et alimentation initiale", 2)
    add_body(doc,
        "Déployer une plateforme de suivi sans données historiques prive le maître d'ouvrage "
        "de toute perspective temporelle le jour de la mise en production. Le protocole de "
        "migration est élaboré dès la Mission I en recensant les sources existantes : "
        "rapports annuels, fichiers Excel des DPA, bases partielles de la DSS. Le module "
        "d'import CSV, développé en Mission II, charge ces données avec validation de format, "
        "détection des doublons et rapport d'import traçable. Les données 2020-2026 sont "
        "migrées avant la mise en production, ce qui permet aux décideurs de consulter "
        "six années de tendances dès le premier jour."
    )

    add_heading(doc, "Tableau récapitulatif des engagements au-delà du CPS", 2)
    add_body(doc,
        "Le tableau ci-après résume les six engagements différenciants avec leur traduction "
        "concrète pour le maître d'ouvrage."
    )
    headers_va = ["Engagement", "Ce que cela signifie pour le MO", "Livrable associé"]
    rows_va = [
        ["Open source + propriété totale",
         "Aucune dépendance après garantie. Migration future libre.",
         "Code source + documentation à la réception provisoire"],
        ["Pipeline CI/CD sécurisé",
         "Toute version livrée est testée, analyse OWASP passée.",
         "Rapport build CI joint à chaque livraison"],
        ["Audit log immuable",
         "Chaque action traçable, exportable pour audit DGSSI.",
         "Journal d'audit accessible par l'administrateur DSS"],
        ["Tests sécu en préproduction",
         "Aucune CVE critique ouverte au déploiement. Rapport livré.",
         "Rapport OWASP joint à la version bêta (fin M8)"],
        ["Isolation RLS par région",
         "Cloisonnement garanti au niveau BDD, pas seulement applicatif.",
         "Scénario de test de recette dédié"],
        ["Migration données historiques",
         "Plateforme opérationnelle dès J1 avec historique 2020-2026.",
         "Rapport d'import fourni lors de la réception provisoire"],
    ]
    add_table(doc, headers_va, rows_va,
              caption="Tableau 10. Engagements différenciants et livrables associés",
              col_widths=[4.0, 5.5, 5.0])


def section_gestion_risques(doc, risk_path):
    add_heading(doc, "Gestion des risques", 1, numbered="9")

    add_body(doc,
        "Six risques spécifiques à cette mission ont été identifiés par croisement entre "
        "l'analyse du CPS et le retour d'expérience des missions SABIL et DigiTPME. "
        "Chaque risque est accompagné d'une mesure de mitigation intégrée dans le planning "
        "de la mission concernée, et d'un indicateur de suivi permettant d'en vérifier "
        "l'évolution lors des comités de pilotage."
    )

    add_heading(doc, "Registre des risques", 2)

    headers_r = ["Risque", "Prob.", "Impact", "Mesure de mitigation", "Indicateur de suivi"]
    rows_r = [
        ["R1 - Disponibilité limitée des agents DSS pour les ateliers de validation",
         "Moyen", "Fort",
         "Planifier les ateliers quatre semaines à l'avance. Sessions de deux heures maximum. "
         "Session de rattrapage systématiquement prévue dans la convocation.",
         "Taux de participation aux ateliers (cible : 80% minimum)"],
        ["R2 - Instabilité du référentiel pendant la phase de développement",
         "Moyen", "Fort",
         "Geler le référentiel par signature du SFD à la fin du Mois 3. "
         "Toute évolution postérieure traitée comme avenant formel au marché.",
         "Nombre de demandes de modification du référentiel après gel"],
        ["R3 - Absence de données historiques disponibles à la mise en production",
         "Elevé", "Moyen",
         "Recenser les sources existantes dès la Mission I. "
         "Développer le module d'import CSV en Mission II. Migration avant Mois 9.",
         "Taux de couverture historique atteint avant mise en production"],
        ["R4 - Délais dans la validation des livrables par le comité de pilotage",
         "Moyen", "Fort",
         "Délai de validation de dix jours contractualisé. Relance formelle au 7e jour. "
         "Approbation tacite après quinze jours sans retour écrit.",
         "Délai moyen de validation par livrable (cible : 10 jours)"],
        ["R5 - Accès non autorisé aux données sensibles (indicateurs non publiés)",
         "Faible", "Fort",
         "Row-Level Security par région. Audit log immuable. Tests OWASP avant chaque déploiement. "
         "Chiffrement TLS sur tous les flux internes et externes.",
         "Nombre de tentatives d'accès non autorisé détectées par le journal d'audit"],
        ["R6 - Interopérabilité difficile avec des SI existants non documentés",
         "Elevé", "Moyen",
         "Cartographier les SI existants dès la Semaine 1. "
         "Développer une API d'import REST + CSV pour les sources sans connecteur natif.",
         "Nombre de sources intégrées / nombre de sources identifiées en Mission I"],
    ]
    add_table(doc, headers_r, rows_r,
              caption="Tableau 11. Registre des risques identifiés",
              col_widths=[4.0, 1.0, 1.2, 5.0, 3.8])

    add_heading(doc, "Matrice de criticité", 2)

    add_body(doc,
        "La matrice ci-après positionne les six risques selon leur probabilité d'occurrence "
        "et leur impact potentiel sur la mission. Les risques R1, R2 et R4, classés en zone "
        "modérée à élevée, font l'objet d'une surveillance mensuelle au comité de pilotage. "
        "R5 est positionné en zone critique malgré sa faible probabilité : son impact sur "
        "la confiance du maître d'ouvrage justifie les mesures de sécurité multi-niveaux."
    )

    add_figure(doc, risk_path,
               "Figure 11. Matrice de criticité des risques identifiés",
               width_cm=13.0)

    add_heading(doc, "Mesures transverses", 2)

    add_body(doc,
        "Quatre mesures s'appliquent à l'ensemble de la mission indépendamment des risques "
        "spécifiques. Le registre des risques est revu à chaque comité de pilotage : les "
        "niveaux de probabilité sont mis à jour selon l'état réel du projet, pas figés "
        "à la valeur initiale. Le rapport mensuel transmis au MO documente les risques "
        "actifs et les décisions demandées, créant une traçabilité qui protège les deux "
        "parties. Une marge de cinq jours ouvrés par mission est intégrée au planning et "
        "activée uniquement sur décision formelle du chef de projet. Enfin, un plan de "
        "continuité documente les procédures de reprise pour les trois composants critiques "
        "(API principale, base de données, IAM) avec des délais de reprise testés."
    )
    add_bullet_dot(doc, [
        "Revue du registre à chaque comité de pilotage, avec mise à jour des niveaux de probabilité",
        "Rapport mensuel MO : risques actifs, décisions demandées, état des mitigations en cours",
        "Plan de continuité documenté et testé pour les 3 composants critiques : API, BDD, IAM",
    ])

    add_heading(doc, "Niveaux de service pendant la période de garantie", 2)

    add_body(doc,
        "La période de garantie de douze mois est encadrée par un dispositif structuré, "
        "et non par une obligation générique de corriger les anomalies. Trois niveaux de "
        "criticité définissent les engagements de prise en charge et de résolution, avec "
        "un rapport mensuel transmis au MO sur l'état des anomalies traitées. Toute "
        "anomalie résolue donne lieu à une note technique transmise à l'équipe DSS, "
        "contribuant au transfert de compétences prévu par l'Article 8 du CPS."
    )
    add_bullet_square(doc, [
        "Criticité 1 - Bloquante : prise en charge sous 4 heures, résolution sous 24 heures ouvrées",
        "Criticité 2 - Dégradée : prise en charge sous 24 heures, résolution sous 5 jours ouvrés",
        "Criticité 3 - Mineure : intégrée au lot de maintenance mensuel",
    ])


def generate_cicd_pipeline():
    """Pipeline CI/CD - Graphviz horizontal."""
    g = Digraph('cicd', format='png')
    g.attr(rankdir='LR', dpi='250', bgcolor='white',
           nodesep='0.3', ranksep='0.5', fontname='Calibri')

    stages = [
        ('COMMIT', 'Commit\ncode source', '#37474F', 'white'),
        ('LINT',   'Lint +\nTests unitaires', '#1565C0', 'white'),
        ('BUILD',  'Build\nimage Docker', '#1B4F35', 'white'),
        ('OWASP',  'Analyse\nOWASP Dep.', '#6A1B9A', 'white'),
        ('STAGING','Déploiement\nStaging', '#E65100', 'white'),
        ('INTEG',  'Tests\nintégration', '#1B4F35', 'white'),
        ('SECU',   'Tests sécu\nOWASP ASVS', '#880E4F', 'white'),
        ('PROD',   'Déploiement\nProduction', '#1565C0', 'white'),
    ]

    for nid, label, color, fc in stages:
        g.node(nid, label=label,
               shape='rectangle', style='filled,rounded',
               fillcolor=color, fontcolor=fc,
               fontsize='10', fontname='Calibri',
               width='1.4', height='0.7')

    for i in range(len(stages) - 1):
        g.edge(stages[i][0], stages[i + 1][0],
               color='#455A64', penwidth='1.5')

    g.edge('LINT', 'COMMIT', label='  Echec',
           color='#C62828', penwidth='1.0', style='dashed',
           fontsize='8', fontcolor='#C62828', constraint='false')

    path = os.path.join(tempfile.gettempdir(), 'cicd_pipeline.png')
    g.render(path.replace('.png', ''), format='png', cleanup=True)
    return path


def generate_risk_matrix():
    """Matrice de criticité des risques (matplotlib 3x3)."""
    fig, ax = plt.subplots(figsize=(14 / 2.54, 11 / 2.54))

    zones = [
        ((0, 0), '#C8E6C9'), ((1, 0), '#C8E6C9'), ((2, 0), '#FFF9C4'),
        ((0, 1), '#C8E6C9'), ((1, 1), '#FFF9C4'), ((2, 1), '#FFCCBC'),
        ((0, 2), '#FFF9C4'), ((1, 2), '#FFCCBC'), ((2, 2), '#FFCDD2'),
    ]
    for (xi, yi), color in zones:
        ax.add_patch(mpatches.Rectangle((xi, yi), 1, 1,
                     facecolor=color, edgecolor='#9E9E9E', linewidth=0.8))

    risks = [
        ("R1 - Disponibilité\nagents DSS",       1.0,  2.15),
        ("R2 - Référentiel\ninstable",            1.35, 1.85),
        ("R3 - Données\nhistoriques absentes",    1.85, 1.15),
        ("R4 - Délais\nvalidation livrables",     0.65, 1.85),
        ("R5 - Accès\nnon autorisé",              0.15, 1.65),
        ("R6 - Interop.\nSI existants",           2.15, 0.85),
    ]

    for label, xp, yp in risks:
        ax.scatter(xp, yp, s=55, color='#37474F', zorder=5)
        ax.text(xp + 0.06, yp, label, fontsize=6.5, va='center',
                color='#212121', fontname='Calibri')

    ax.set_xlim(0, 3)
    ax.set_ylim(0, 3)
    ax.set_xticks([0.5, 1.5, 2.5])
    ax.set_xticklabels(['Faible', 'Moyen', 'Elevé'], fontsize=9, fontname='Calibri')
    ax.set_yticks([0.5, 1.5, 2.5])
    ax.set_yticklabels(['Faible', 'Moyen', 'Fort'], fontsize=9, fontname='Calibri')
    ax.set_xlabel('Probabilité', fontsize=9, fontname='Calibri', labelpad=6)
    ax.set_ylabel('Impact', fontsize=9, fontname='Calibri', labelpad=6)

    from matplotlib.patches import Patch
    legend = [Patch(facecolor='#C8E6C9', label='Risque faible'),
              Patch(facecolor='#FFF9C4', label='Risque modéré'),
              Patch(facecolor='#FFCDD2', label='Risque élevé')]
    ax.legend(handles=legend, fontsize=7.5, loc='lower right',
              framealpha=0.9, prop={'family': 'Calibri', 'size': 7.5})

    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'risk_matrix.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    return path


# ---------------------------------------------------------------
# Assemblage du document
# ---------------------------------------------------------------

def build_document():
    print("Génération des figures...")
    gantt_path = generate_gantt()
    print(f"  Gantt: {gantt_path}")
    arch_path = generate_architecture_diagram()
    print(f"  Architecture: {arch_path}")
    org_path = generate_organigramme()
    print(f"  Organigramme: {org_path}")
    usecase_path = generate_usecase_diagram()
    print(f"  Cas d'utilisation: {usecase_path}")
    classes_path = generate_class_diagram()
    print(f"  Classes: {classes_path}")
    sequence_path = generate_sequence_diagram()
    print(f"  Séquence: {sequence_path}")
    sgg_path = generate_sgg_structure()
    print(f"  Structure SGG: {sgg_path}")
    circuit_path = generate_circuit_collecte()
    print(f"  Circuit collecte: {circuit_path}")
    gar_path = generate_gar_chain()
    print(f"  Chaîne GAR: {gar_path}")
    cicd_path = generate_cicd_pipeline()
    print(f"  Pipeline CI/CD: {cicd_path}")
    risk_path = generate_risk_matrix()
    print(f"  Matrice risques: {risk_path}")

    print("Construction du document...")
    doc = Document()
    set_margins(doc)
    set_document_font_theme(doc)
    add_page_numbers(doc)

    page_de_garde(doc)
    section_sommaire(doc)
    section_comprehension_contexte(doc, sgg_path, circuit_path)
    doc.add_page_break()
    section_comprehension_mission(doc, gar_path)
    doc.add_page_break()
    section_approche_methodologique(doc)
    doc.add_page_break()
    section_architecture(doc, arch_path, org_path, usecase_path, classes_path, sequence_path)
    doc.add_page_break()
    section_equipe(doc, org_path)
    doc.add_page_break()
    section_planning(doc, gantt_path)
    doc.add_page_break()
    section_chronogramme(doc)
    doc.add_page_break()
    section_valeur_ajoutee(doc, cicd_path)
    doc.add_page_break()
    section_gestion_risques(doc, risk_path)

    return doc


def inject_template_cover(output_path, template_path):
    """Injecte la page de garde du template DSS dans le document généré."""
    import uuid

    with zipfile.ZipFile(template_path, 'r') as zt:
        template_doc_xml = zt.read('word/document.xml').decode('utf-8')
        template_rels_xml = zt.read('word/_rels/document.xml.rels').decode('utf-8')
        cover_image_bytes = zt.read('word/media/image1.png')

    # Extraire le bloc SDT (page de garde) du template
    sdt_start = template_doc_xml.find('<w:sdt>')
    if sdt_start == -1:
        print("Avertissement: bloc de couverture introuvable dans le template")
        return
    depth = 0
    i = sdt_start
    while i < len(template_doc_xml):
        if template_doc_xml[i:i+7] == '<w:sdt>':
            depth += 1
            i += 7
        elif template_doc_xml[i:i+8] == '<w:sdt ':
            depth += 1
            i += 8
        elif template_doc_xml[i:i+8] == '</w:sdt>':
            depth -= 1
            if depth == 0:
                sdt_end = i + 8
                break
            i += 8
        else:
            i += 1
    cover_xml = template_doc_xml[sdt_start:sdt_end]

    # Trouver rId8 dans le template -> media/image1.png
    template_rids = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', template_rels_xml))
    # On renomme pour eviter collision dans notre docx
    new_rid = "rIdCover1"
    cover_xml = cover_xml.replace('r:embed="rId8"', f'r:embed="{new_rid}"')

    # Modifier le docx généré
    tmp_path = output_path + ".tmp.docx"
    with zipfile.ZipFile(output_path, 'r') as zin, zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        for name in names:
            data = zin.read(name)

            if name == 'word/document.xml':
                doc_xml = data.decode('utf-8')
                # Injerer le bloc couverture + saut de page au début du body
                body_start = doc_xml.find('<w:body>') + len('<w:body>')
                page_break = (
                    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                    '<w:r><w:br w:type="page"/></w:r></w:p>'
                )
                doc_xml = doc_xml[:body_start] + cover_xml + page_break + doc_xml[body_start:]
                zout.writestr(name, doc_xml.encode('utf-8'))

            elif name == 'word/settings.xml':
                settings_xml = data.decode('utf-8')
                if 'w:updateFields' not in settings_xml:
                    settings_xml = settings_xml.replace(
                        '</w:settings>',
                        '<w:updateFields w:val="1"/></w:settings>'
                    )
                zout.writestr(name, settings_xml.encode('utf-8'))

            elif name == 'word/_rels/document.xml.rels':
                rels_xml = data.decode('utf-8')
                new_rel = (
                    f'<Relationship Id="{new_rid}" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                    'Target="media/cover_image1.png"/>'
                )
                rels_xml = rels_xml.replace('</Relationships>', new_rel + '</Relationships>')
                zout.writestr(name, rels_xml.encode('utf-8'))

            else:
                zout.writestr(name, data)

        # Ajouter l'image de couverture
        zout.writestr('word/media/cover_image1.png', cover_image_bytes)

    shutil.move(tmp_path, output_path)
    print("Page de garde template injectée.")


if __name__ == "__main__":
    doc = build_document()
    output_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "offre_technique_plateforme_sgg_20260606.docx"
    )
    doc.save(output_path)
    template_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "note methdologique DSS.docx"
    )
    if os.path.exists(template_path):
        inject_template_cover(output_path, template_path)
    print(f"\nDocument sauvegardé: {output_path}")
