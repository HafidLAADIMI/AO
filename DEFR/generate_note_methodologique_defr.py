# -*- coding: utf-8 -*-
"""
Note Méthodologique - Offre Technique
Objet : Conception, développement, mise en œuvre de bases de données et système
        d'information apprenant (SIS Intégré) destiné aux établissements de formation
        professionnelle agricole (EFPA), et sa mise en service en mode SaaS
Référence : AOOI N°01/2026/DEFR/DETFP
Maître d'ouvrage : DEFR - Ministère de l'Agriculture, de la Pêche Maritime,
                   du Développement Rural et des Eaux et Forêts
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
import tempfile
import os
import zipfile
import re
import shutil
import textwrap
from graphviz import Digraph

# ---------------------------------------------------------------
# Constantes de couleurs
# ---------------------------------------------------------------
PRIMARY_HEX = "15607F"
LIGHT_BG_HEX = "E8F1F5"
WHITE_HEX = "FFFFFF"
PRIMARY_RGB = RGBColor(0x15, 0x60, 0x82)
WHITE_RGB = RGBColor(0xFF, 0xFF, 0xFF)
DARK_RGB = RGBColor(0x1A, 0x1A, 0x1A)
H1_COLOR = RGBColor(0x15, 0x60, 0x82)
H2H3_COLOR = RGBColor(0x0F, 0x47, 0x61)

heading_counter = [0, 0, 0]
figure_counter = [0]


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


def add_bullet(doc, items):
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
    figure_counter[0] += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(img_path, width=Cm(width_cm))
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cr = cp.add_run(f"Figure {figure_counter[0]}. {caption}")
    cr.font.name = 'Calibri'
    cr.font.size = Pt(9)
    cr.font.italic = True
    cr.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)


def add_toc_field(doc):
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

    _add_toc_entries(doc)

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(0)
    p_end.paragraph_format.space_after = Pt(0)
    r_end = p_end.add_run()
    fc_end = OxmlElement('w:fldChar')
    fc_end.set(qn('w:fldCharType'), 'end')
    r_end._r.append(fc_end)


def _add_toc_entries(doc):
    entries = [
        (0, "",    "Présentation de la note",                                  3),
        (1, "1",   "Compréhension du contexte",                                4),
        (1, "2",   "Compréhension de la mission",                              5),
        (2, "2.1", "Objectifs et livrables",                                   5),
        (2, "2.2", "Périmètre et limites de la mission",                       6),
        (1, "3",   "Approche méthodologique",                                  7),
        (2, "3.1", "Démarche de conduite de projet",                          7),
        (3, "3.1.1", "Diagnostic et analyse des besoins",                     8),
        (2, "3.2", "Couverture des exigences fonctionnelles (Art.19.5 CPS)",  10),
        (2, "3.3", "Stratégie ETL et migration des données historiques",      12),
        (2, "3.4", "Architecture SaaS et Cloud",                              14),
        (2, "3.5", "Interconnexion et SSO",                                   16),
        (2, "3.6", "Approche IA - Bot APC",                                   17),
        (2, "3.7", "Exploitation, sécurisation et performance des données",   19),
        (1, "4",   "Architecture technique détaillée",                        20),
        (2, "4.1", "Architecture globale de la plateforme",                  20),
        (2, "4.2", "Comparaison et justification des choix technologiques",   21),
        (3, "4.2.1", "Runtime et framework applicatif",                      21),
        (3, "4.2.2", "Système de gestion de base de données",                22),
        (3, "4.2.3", "Style d'architecture applicative",                     23),
        (3, "4.2.4", "Moteur d'inférence pour le Bot APC",                   24),
        (2, "4.3", "Périmètre fonctionnel - Cas d'utilisation",               25),
        (2, "4.4", "Structure des données - Diagramme de classes",            26),
        (2, "4.5", "Circuit de traitement - Diagramme de séquence",           27),
        (2, "4.6", "Architecture de déploiement et haute disponibilité",      28),
        (2, "4.7", "Pipeline de déploiement continu et sécurisé",             29),
        (1, "5",   "Équipe proposée",                                         30),
        (2, "5.1", "Principe de constitution de l'équipe",                   30),
        (2, "5.2", "Profils détaillés des experts proposés",                  31),
        (2, "5.3", "Sous-critères de notation et niveau visé (Art.10 RC)",    32),
        (2, "5.4", "Couverture des services par profil",                     33),
        (1, "6",   "Planning d'exécution",                                    34),
        (1, "7",   "Chronogramme d'affectation du personnel",                 35),
        (1, "8",   "Plan de réversibilité et transfert de compétences",       36),
        (1, "9",   "Valeur ajoutée et points distinctifs",                    37),
        (1, "10",  "Gestion des risques",                                     38),
    ]

    TAB_POS = "8789"

    for level, num, title, page in entries:
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(3 if level <= 1 else 1)
        para.paragraph_format.left_indent = Cm(0) if level <= 1 else Cm(0.7)

        pPr = para._p.get_or_add_pPr()
        tabs_el = OxmlElement('w:tabs')
        tab_el = OxmlElement('w:tab')
        tab_el.set(qn('w:val'), 'right')
        tab_el.set(qn('w:leader'), 'dot')
        tab_el.set(qn('w:pos'), TAB_POS)
        tabs_el.append(tab_el)
        pPr.append(tabs_el)

        if level == 0:
            run_title = para.add_run(title)
            run_title.font.name = 'Calibri'
            run_title.font.size = Pt(11)
            run_title.font.bold = True
            run_title.font.italic = True
            run_title.font.color.rgb = DARK_RGB
        elif level == 1:
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
            run_num.font.color.rgb = RGBColor(0x0F, 0x47, 0x61)

            run_title = para.add_run(title)
            run_title.font.name = 'Calibri'
            run_title.font.size = Pt(10)
            run_title.font.color.rgb = DARK_RGB

        run_tab = para.add_run('\t')
        run_tab.font.name = 'Calibri'
        run_tab.font.size = Pt(10 if level >= 2 else 11)

        run_page = para.add_run(str(page))
        run_page.font.name = 'Calibri'
        run_page.font.size = Pt(10 if level >= 2 else 11)
        run_page.font.bold = (level == 1)
        run_page.font.color.rgb = PRIMARY_RGB if level == 1 else DARK_RGB


# ---------------------------------------------------------------
# Génération des figures
# ---------------------------------------------------------------

def generate_architecture_diagram():
    dot = Digraph('architecture_defr')
    dot.attr(rankdir='TB', nodesep='0.4', ranksep='0.65', dpi='200',
             fontname='Helvetica', bgcolor='white')

    node_box = {'shape': 'box', 'style': 'filled', 'fontname': 'Helvetica',
                'fontsize': '8.5', 'margin': '0.14,0.08'}
    node_cyl = {'shape': 'cylinder', 'style': 'filled', 'fontname': 'Helvetica',
                'fontsize': '8.5', 'margin': '0.14,0.08'}

    with dot.subgraph(name='cluster_idp') as c:
        c.attr(label='Identité', style='rounded,filled', fillcolor='#FFF8E1',
               color='#F57F17', fontname='Helvetica', fontsize='9')
        c.node('entra', 'Microsoft Entra ID\n(OIDC + MFA)', fillcolor='#FFE082', **node_box, width='2.1')

    with dot.subgraph(name='cluster_clients') as c:
        c.attr(label='Extranet (SPA)', style='rounded,filled', fillcolor='#F1F8E9',
               color='#558B2F', fontname='Helvetica', fontsize='9')
        c.node('spa_stag', 'Espace\nStagiaire', fillcolor='#DCEDC8', **node_box, width='1.5')
        c.node('spa_ens', 'Espace\nEnseignant', fillcolor='#DCEDC8', **node_box, width='1.5')
        c.node('spa_adm', 'Espace\nAdministration', fillcolor='#DCEDC8', **node_box, width='1.5')

    with dot.subgraph(name='cluster_gw') as c:
        c.attr(label='Gateway', style='rounded,filled', fillcolor='#E3F2FD',
               color='#1565C0', fontname='Helvetica', fontsize='9')
        c.node('gateway', 'API Gateway\n(NestJS)\nauth guard, audit, routing', fillcolor='#BBDEFB',
               **node_box, width='2.6')

    with dot.subgraph(name='cluster_services') as c:
        c.attr(label='Micro-services (NestJS / TypeScript)', style='rounded,filled',
               fillcolor='#F3E5F5', color='#6A1B9A', fontname='Helvetica', fontsize='9')
        c.node('ref', 'Référentiel\nEFPA, filières, modules', fillcolor='#E1BEE7', **node_box, width='2.1')
        c.node('parcours', 'Parcours Apprenant\nAdmissions, scolarité, diplomation', fillcolor='#E1BEE7', **node_box, width='2.4')
        c.node('planning', 'Planning &\nRessources', fillcolor='#E1BEE7', **node_box, width='2.0')
        c.node('stages', 'Stages &\nInsertion', fillcolor='#E1BEE7', **node_box, width='2.0')
        c.node('certif', 'Certification\n(Barid eSign, QR)', fillcolor='#E1BEE7', **node_box, width='2.1')
        c.node('bi', 'BI & Reporting\n(Cube.dev)', fillcolor='#E1BEE7', **node_box, width='2.0')
        c.node('bot', 'Bot APC / IA\n(LangChain.js)', fillcolor='#E1BEE7', **node_box, width='2.0')
        c.node('interop', 'Interopérabilité &\nMigration', fillcolor='#E1BEE7', **node_box, width='2.1')

    with dot.subgraph(name='cluster_async') as c:
        c.attr(label='Traitement asynchrone', style='rounded,filled', fillcolor='#FCE4EC',
               color='#880E4F', fontname='Helvetica', fontsize='9')
        c.node('bullmq', 'BullMQ / Redis', fillcolor='#F8BBD0', **node_box, width='1.9')

    with dot.subgraph(name='cluster_ext') as c:
        c.attr(label='Services externes', style='rounded,filled', fillcolor='#ECEFF1',
               color='#37474F', fontname='Helvetica', fontsize='9')
        c.node('baridesign', 'Barid eSign', fillcolor='#CFD8DC', **node_box, width='1.7')
        c.node('moodle', 'LMS Moodle', fillcolor='#CFD8DC', **node_box, width='1.7')
        c.node('m365', 'M365 Education', fillcolor='#CFD8DC', **node_box, width='1.7')
        c.node('gpu', 'Moteur GPU\nvLLM/TGI (NVIDIA A10)', fillcolor='#CFD8DC', **node_box, width='2.1')

    with dot.subgraph(name='cluster_data') as c:
        c.attr(label='Couche Données', style='rounded,filled', fillcolor='#E8EAF6',
               color='#1A237E', fontname='Helvetica', fontsize='9')
        c.node('postgres', 'PostgreSQL\nRLS (efpa_id) + pgvector', fillcolor='#C5CAE9', **node_cyl, width='2.4')
        c.node('audit', 'Audit Trail Store\n(append-only, 12 mois)', fillcolor='#C5CAE9', **node_cyl, width='2.3')

    dot.edge('spa_stag', 'gateway')
    dot.edge('spa_ens', 'gateway')
    dot.edge('spa_adm', 'gateway')
    dot.edge('gateway', 'entra', label='validation token', style='dashed', color='#F57F17')
    for s in ['ref', 'parcours', 'planning', 'stages', 'certif', 'bi', 'bot']:
        dot.edge('gateway', s, label='REST')
    dot.edge('certif', 'bullmq', label='enqueue')
    dot.edge('bullmq', 'baridesign', label='appel API')
    dot.edge('bot', 'gpu', label='HTTP')
    dot.edge('bot', 'postgres', label='retrieval pgvector')
    dot.edge('bi', 'postgres', label='CDC', style='dashed')
    dot.edge('interop', 'moodle', label='sync')
    dot.edge('interop', 'm365', label='Graph API')
    for s in ['ref', 'parcours', 'planning', 'stages', 'certif']:
        dot.edge(s, 'postgres', label='R/W')
    for s in ['ref', 'parcours', 'planning', 'stages', 'certif', 'bi', 'bot']:
        dot.edge(s, 'audit', label='log', style='dotted', color='#880E4F')

    path_base = os.path.join(tempfile.gettempdir(), 'architecture_defr')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_efpa_growth_chart():
    categories = ["Établissements\n(EFPA)", "Stagiaires actifs\n(estimation proportionnelle)"]
    today = [58, 30000]
    target = [100, 51724]

    fig, axes = plt.subplots(1, 2, figsize=(17 / 2.54, 8 / 2.54))

    for ax, label, v_today, v_target in zip(axes, categories, today, target):
        bars = ax.bar(["Aujourd'hui", "Cible (Art.1 CPS)"], [v_today, v_target],
                       color=["#15607F", "#A9D6E5"], edgecolor='white', width=0.55)
        for b, v in zip(bars, [v_today, v_target]):
            ax.text(b.get_x() + b.get_width() / 2, v + max(v_today, v_target) * 0.02,
                    f"{v:,}".replace(",", " "), ha='center', fontsize=8.5, fontweight='bold')
        ax.set_title(label, fontsize=8.8, fontweight='bold')
        ax.spines[['top', 'right']].set_visible(False)
        ax.set_ylim(0, max(v_today, v_target) * 1.2)
        ax.tick_params(labelsize=8)

    fig.suptitle("Extensibilité du dispositif : 58 EFPA aujourd'hui, jusqu'à 100 visés",
                 fontsize=10, fontweight='bold', y=1.04)
    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'growth_defr.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return path


def generate_objectifs_mission_diagram():
    dot = Digraph('objectifs_defr')
    dot.attr(rankdir='LR', nodesep='0.3', ranksep='0.55', dpi='200',
             fontname='Helvetica', bgcolor='white')

    node_obj = {'shape': 'box', 'style': 'filled,rounded', 'fontname': 'Helvetica',
                'fontsize': '8.3', 'margin': '0.15,0.1', 'width': '2.3'}
    node_mec = {'shape': 'box', 'style': 'filled,rounded', 'fontname': 'Helvetica',
                'fontsize': '8.3', 'margin': '0.15,0.1', 'width': '2.3'}
    node_ben = {'shape': 'box', 'style': 'filled,rounded', 'fontname': 'Helvetica',
                'fontsize': '8.3', 'margin': '0.15,0.1', 'width': '2.4'}

    rows = [
        ("o1", "Collecte centralisée\ndes 58 EFPA", "m1", "Référentiel unique\n+ RLS multi-tenant",
         "b1", "Vision DEFR en temps réel,\nsans remontée manuelle"),
        ("o2", "Suivi en temps\nopportun", "m2", "Headless BI,\ntableaux de bord",
         "b2", "Pilotage proactif plutôt\nque rétrospectif"),
        ("o3", "Reprise intégrale de\nl'historique (5 ans)", "m3", "Migration ETL\npar lots contrôlés",
         "b3", "Continuité de service apprenant\nsans rupture"),
        ("o4", "Certification numérique\nanti-fraude", "m4", "Barid eSign\n+ QR Code",
         "b4", "Valeur juridique et vérification\ninstantanée des diplômes"),
    ]

    with dot.subgraph(name='cluster_obj') as c:
        c.attr(label='Objectif (Art.3 CPS)', style='rounded,filled', fillcolor='#E8F1F5',
               color='#15607F', fontname='Helvetica', fontsize='9')
        for oid, otext, *_ in rows:
            c.node(oid, otext, fillcolor='#A9D6E5', **node_obj)

    with dot.subgraph(name='cluster_mec') as c:
        c.attr(label='Mécanisme dans le SIS', style='rounded,filled', fillcolor='#FFF3E0',
               color='#E65100', fontname='Helvetica', fontsize='9')
        for _, _, mid, mtext, *_ in rows:
            c.node(mid, mtext, fillcolor='#FFCC80', **node_mec)

    with dot.subgraph(name='cluster_ben') as c:
        c.attr(label='Bénéfice mesurable', style='rounded,filled', fillcolor='#E8F5E9',
               color='#2E7D32', fontname='Helvetica', fontsize='9')
        for _, _, _, _, bid, btext in rows:
            c.node(bid, btext, fillcolor='#A5D6A7', **node_ben)

    for oid, _, mid, _, bid, _ in rows:
        dot.edge(oid, mid)
        dot.edge(mid, bid)

    path_base = os.path.join(tempfile.gettempdir(), 'objectifs_defr')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_team_points_chart():
    profiles = [
        ("AMENZOU\n(Développeur)", [("Diplôme > Bac+3", 5, "#A9D6E5"), ("Expérience Node.js > 5 ans", 5, "#15607F")]),
        ("MOHY-EDDINE\n(Expert IA)", [("Formation IA", 5, "#A9D6E5"), ("Expérience IA éducatif > 2 ans", 5, "#15607F")]),
        ("OUAZZANE\n(Ingénieur Systèmes)", [("Formation Bac+5", 10, "#A9D6E5"), ("Expérience SI éducatif > 10 ans", 10, "#15607F")]),
        ("EL ORF\n(Chef de Projet)", [("Formation Bac+5", 10, "#A9D6E5"), ("Expérience direction > 15 ans", 5, "#15607F"),
                            ("Projets éducatifs > 3", 5, "#0F4761")]),
    ]

    fig, ax = plt.subplots(figsize=(16 / 2.54, 9 / 2.54))
    y_pos = range(len(profiles))

    for i, (name, segments) in enumerate(profiles):
        left = 0
        for label, value, color in segments:
            ax.barh(i, value, left=left, height=0.55, color=color, edgecolor='white')
            ax.text(left + value / 2, i, f"{value}", ha='center', va='center',
                    fontsize=8, color='white', fontweight='bold')
            left += value

    ax.set_yticks(list(y_pos))
    ax.set_yticklabels([p[0] for p in profiles], fontsize=9)
    ax.set_xlim(0, 22)
    ax.set_xlabel("Points sur le Critère C (Art.10 RC)", fontsize=9)
    ax.set_title("Potentiel de points par profil - total 60/100 sur la note technique",
                 fontsize=10, fontweight='bold')

    totals = [sum(v for _, v, _ in segs) for _, segs in profiles]
    for i, t in enumerate(totals):
        ax.text(t + 0.4, i, f"/{t}", va='center', fontsize=8.5, fontweight='bold', color='#1A1A1A')

    legend_labels = ["Formation / diplôme", "Expérience principale", "Expérience projets éducatifs"]
    legend_colors = ["#A9D6E5", "#15607F", "#0F4761"]
    handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in legend_colors]
    ax.legend(handles, legend_labels, loc='lower right', fontsize=7.5, frameon=False)

    ax.spines[['top', 'right']].set_visible(False)
    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'team_points_defr.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return path


def generate_usecase_diagram():
    dot = Digraph('usecase_defr')
    dot.attr(rankdir='LR', nodesep='0.35', ranksep='1.1', dpi='200',
             fontname='Helvetica', bgcolor='white')

    actors = [
        ('candidat',  'CANDIDAT',           '#15607F'),
        ('agent',     'AGENT EFPA',          '#0F4761'),
        ('comite',    'COMITÉ DE\nDÉLIBÉRATION', '#1A7A9C'),
        ('admin',     'ADMINISTRATEUR\nDEFR', '#37474F'),
    ]
    for aid, alabel, acolor in actors:
        dot.node(aid, alabel, shape='box', style='filled',
                 fillcolor=acolor, fontcolor='white',
                 fontname='Helvetica', fontsize='9', width='1.5', height='0.5')

    with dot.subgraph(name='cluster_system') as s:
        s.attr(label='Système Admissions & Concours', style='rounded',
               color='#15607F', fontname='Helvetica', fontsize='10', bgcolor='#E8F1F5')
        ucs = [
            ('uc_soumettre',  'Soumettre une\ncandidature'),
            ('uc_statut',     'Consulter le statut\ndu dossier'),
            ('uc_preselect',  'Présélectionner\nles dossiers'),
            ('uc_planifier',  'Planifier une séance\nde concours'),
            ('uc_convoquer',  'Convoquer les\ncandidats'),
            ('uc_notes',      'Saisir les notes\nde concours'),
            ('uc_deliberer',  'Délibérer'),
            ('uc_publier',    'Publier les résultats'),
            ('uc_attente',    'Gérer la liste\nd\'attente'),
            ('uc_inscrire',   'Confirmer l\'inscription\ndéfinitive'),
            ('uc_param',      'Paramétrer les capacités\npar filière'),
        ]
        for uid, ulabel in ucs:
            s.node(uid, ulabel, shape='ellipse', style='filled',
                   fillcolor='#FFFFFF', color='#15607F',
                   fontname='Helvetica', fontsize='8')

    assoc = [
        ('candidat', 'uc_soumettre'), ('candidat', 'uc_statut'),
        ('agent', 'uc_preselect'), ('agent', 'uc_planifier'), ('agent', 'uc_convoquer'),
        ('agent', 'uc_notes'),
        ('comite', 'uc_deliberer'), ('comite', 'uc_publier'), ('comite', 'uc_attente'),
        ('comite', 'uc_inscrire'),
        ('admin', 'uc_param'),
    ]
    for src, tgt in assoc:
        dot.edge(src, tgt, arrowhead='none')
    dot.edge('uc_planifier', 'uc_convoquer', label='(include)', style='dashed', fontsize='7')
    dot.edge('uc_deliberer', 'uc_publier', label='(include)', style='dashed', fontsize='7')
    dot.edge('uc_publier', 'uc_attente', label='(include)', style='dashed', fontsize='7')

    path_base = os.path.join(tempfile.gettempdir(), 'usecase_defr')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_class_diagram():
    dot = Digraph('classes_defr')
    dot.attr(rankdir='TB', nodesep='0.5', ranksep='0.6', dpi='200',
             fontname='Helvetica', bgcolor='white')

    rec = {'shape': 'record', 'style': 'filled', 'fillcolor': '#E8F1F5',
           'fontname': 'Helvetica', 'fontsize': '8', 'color': '#15607F'}

    dot.node('Candidat',
        '{Candidat|id: UUID\\lnom, prenom: String\\lcin: String\\lemail, telephone: String\\ldateNaissance: Date\\l}',
        **rec)
    dot.node('Candidature',
        '{Candidature|id: UUID\\lcandidatId: UUID (FK)\\lfiliereId: UUID (FK réf.)\\lefpaId: UUID (FK réf.)\\lstatut: Enum\\ldateDepot: DateTime\\ldocumentsJustificatifs: [String]\\l}',
        **rec)
    dot.node('SessionConcours',
        '{SessionConcours|id: UUID\\lfiliereId: UUID (FK réf.)\\lefpaId: UUID (FK réf.)\\ldateConcours: Date\\lcapaciteMax: Integer\\lstatut: Enum\\l}',
        **rec)
    dot.node('Convocation',
        '{Convocation|candidatureId: UUID (FK)\\ldateConvocation: DateTime\\llieu, heure: String\\lstatutEnvoi: Enum\\l}',
        **rec)
    dot.node('Note',
        '{Note|candidatureId: UUID (FK)\\lvaleur: Decimal\\lmatiere: String\\lsaisiPar: UUID\\l}',
        **rec)
    dot.node('Deliberation',
        '{Deliberation|sessionConcoursId: UUID (FK)\\ldateDeliberation: DateTime\\lcomiteMembres: [UUID]\\ldecision: Enum\\l}',
        **rec)
    dot.node('ListeAttente',
        '{ListeAttente|candidatureId: UUID (FK)\\lrang: Integer\\l}',
        **rec)
    dot.node('InscriptionDefinitive',
        '{InscriptionDefinitive|candidatureId: UUID (FK)\\ldateInscription: Date\\lanneeScolaire: String\\l}',
        **rec)

    dot.edge('Candidat', 'Candidature', label='1..*', dir='none')
    dot.edge('Candidature', 'SessionConcours', label='*..1', dir='none')
    dot.edge('SessionConcours', 'Convocation', label='1', dir='none')
    dot.edge('SessionConcours', 'Note', label='1', dir='none')
    dot.edge('Note', 'Deliberation', label='*..1', dir='none')
    dot.edge('Deliberation', 'ListeAttente', label='1', dir='none', style='dashed')
    dot.edge('Deliberation', 'InscriptionDefinitive', label='1', dir='none', style='dashed')

    path_base = os.path.join(tempfile.gettempdir(), 'classes_defr')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_sequence_diagram():
    participants = [
        ('Agent EFPA',              '#15607F'),
        ('Gateway',                 '#0F4761'),
        ('Parcours Apprenant Svc',  '#1A7A9C'),
        ('Référentiel Svc',         '#2E7D32'),
        ('Audit Store',             '#37474F'),
        ('File (BullMQ)',           '#880E4F'),
    ]
    messages = [
        (0, 1, 'POST /concours/{id}/deliberation'),
        (1, 2, 'valider token + rôle'),
        (2, 3, 'GET capacité filière'),
        (3, 2, 'retour capacité'),
        (2, 2, 'calcule classement,\ndétermine admis/attente'),
        (2, 4, 'log audit (action, agent, horodatage)'),
        (2, 5, 'enqueue notifications candidats'),
        (2, 1, '200 OK (résultats)'),
        (1, 0, 'réponse'),
        (5, 5, 'pour chaque candidat:\nenvoyer email/SMS'),
    ]

    np_ = len(participants)
    nm = len(messages)
    COL = 2.5
    ROW = 1.05

    fig_w = (np_ * COL + 0.8) / 2.54
    fig_h = (nm * ROW + 3.2) / 2.54
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))

    X = [i * COL for i in range(np_)]
    ax.set_xlim(-1.4, X[-1] + 1.4)
    ax.set_ylim(-(nm * ROW + 0.8), 2.1)
    ax.axis('off')

    for i, (name, color) in enumerate(participants):
        ax.add_patch(plt.Rectangle((X[i] - 1.1, 0.65), 2.2, 0.85,
                     facecolor=color, edgecolor='white', linewidth=0.5, zorder=2))
        ax.text(X[i], 1.08, name, ha='center', va='center',
                fontsize=6.8, color='white', fontweight='bold', zorder=3)
        ax.plot([X[i], X[i]], [0.65, -(nm * ROW + 0.5)],
                color='#BBBBBB', linewidth=0.7, linestyle=(0, (5, 3)), zorder=1)

    for mi, (src, tgt, label) in enumerate(messages):
        y = -(mi + 1) * ROW
        xs, xt = X[src], X[tgt]
        if src == tgt:
            ax.annotate('', xy=(xt + 0.55, y - 0.18), xytext=(xs, y),
                        arrowprops=dict(arrowstyle='->', color='#1A3A2A', lw=0.85,
                                         connectionstyle='arc3,rad=1.4'))
        else:
            ax.annotate('', xy=(xt, y), xytext=(xs, y),
                        arrowprops=dict(arrowstyle='->', color='#1A3A2A', lw=0.85))
        ax.text((xs + xt) / 2, y + 0.13, f'{mi + 1}. {label}',
                ha='center', va='bottom', fontsize=6.0, color='#1A3A2A')

    ax.set_title("Délibération et publication des résultats d'un concours",
                 fontsize=9, fontweight='bold', y=0.99)
    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'sequence_defr.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return path


def generate_deployment_diagram():
    dot = Digraph('deployment_defr')
    dot.attr(rankdir='TB', nodesep='0.4', ranksep='0.55', dpi='200',
             fontname='Helvetica', bgcolor='white', compound='true')

    node_srv = {'shape': 'box3d', 'style': 'filled', 'fontname': 'Helvetica',
                'fontsize': '8.3', 'margin': '0.16,0.1'}
    node_cyl = {'shape': 'cylinder', 'style': 'filled', 'fontname': 'Helvetica',
                'fontsize': '8.3', 'margin': '0.14,0.08'}

    with dot.subgraph(name='cluster_edge') as c:
        c.attr(label='Périmètre (Cloud souverain Maroc)', style='rounded,filled',
               fillcolor='#FFF3E0', color='#E65100', fontname='Helvetica', fontsize='9')
        c.node('waf', 'WAF\n(injections, XSS,\nforce brute SSO)', fillcolor='#FFCC80', **node_srv, width='1.9')
        c.node('ddos', 'Anti-DDoS\n(réseau + applicatif)', fillcolor='#FFCC80', **node_srv, width='1.9')
        c.node('lb', 'Load Balancer\nMulti-AZ', fillcolor='#FFCC80', **node_srv, width='1.7')

    with dot.subgraph(name='cluster_k8s') as c:
        c.attr(label='Cluster Kubernetes - Multi-AZ', style='rounded,filled',
               fillcolor='#E3F2FD', color='#1565C0', fontname='Helvetica', fontsize='9')

        with c.subgraph(name='cluster_az1') as az1:
            az1.attr(label='Zone de disponibilité 1', style='rounded,filled',
                     fillcolor='#BBDEFB', color='#1565C0', fontname='Helvetica', fontsize='8.5')
            az1.node('pods_az1', 'Pods applicatifs\n(Gateway + 8 services)\nHPA actif', fillcolor='#90CAF9',
                     **node_srv, width='2.4')

        with c.subgraph(name='cluster_az2') as az2:
            az2.attr(label='Zone de disponibilité 2', style='rounded,filled',
                     fillcolor='#BBDEFB', color='#1565C0', fontname='Helvetica', fontsize='8.5')
            az2.node('pods_az2', 'Pods applicatifs\n(répliques bascule\nautomatique)', fillcolor='#90CAF9',
                     **node_srv, width='2.4')

        with c.subgraph(name='cluster_gpu') as gpu:
            gpu.attr(label='Pool GPU dédié (isolé)', style='rounded,filled',
                     fillcolor='#F3E5F5', color='#6A1B9A', fontname='Helvetica', fontsize='8.5')
            gpu.node('gpupool', 'Nœuds GPU\nNVIDIA A10\nmoteur vLLM/TGI', fillcolor='#E1BEE7',
                     **node_srv, width='2.2')

    with dot.subgraph(name='cluster_data') as c:
        c.attr(label='Couche données (Multi-AZ)', style='rounded,filled', fillcolor='#E8EAF6',
               color='#1A237E', fontname='Helvetica', fontsize='9')
        c.node('pg_primary', 'PostgreSQL Primaire\nRLS (efpa_id) + pgvector', fillcolor='#C5CAE9',
               **node_cyl, width='2.4')
        c.node('pg_standby', 'PostgreSQL Standby\n(réplication synchrone)', fillcolor='#C5CAE9',
               **node_cyl, width='2.4')
        c.node('redis', 'Redis Cluster\n(BullMQ, sessions)', fillcolor='#C5CAE9', **node_cyl, width='2.0')
        c.node('objstore', 'Stockage objets\n(diplômes, ressources)', fillcolor='#C5CAE9', **node_cyl, width='2.1')

    with dot.subgraph(name='cluster_backup') as c:
        c.attr(label='Sauvegarde et audit', style='rounded,filled', fillcolor='#FCE4EC',
               color='#880E4F', fontname='Helvetica', fontsize='9')
        c.node('pitr', 'Sauvegardes PITR\nquotidiennes externalisées', fillcolor='#F8BBD0',
               **node_cyl, width='2.3')
        c.node('audit_store', 'Audit Trail Store\n(append-only, 12 mois,\nsupport non-modifiable)',
               fillcolor='#F8BBD0', **node_cyl, width='2.4')

    dot.edge('waf', 'lb')
    dot.edge('ddos', 'lb')
    dot.edge('lb', 'pods_az1')
    dot.edge('lb', 'pods_az2', label='bascule auto', style='dashed')
    dot.edge('pods_az1', 'gpupool', label='HTTP interne')
    dot.edge('pods_az1', 'pg_primary', label='R/W')
    dot.edge('pods_az2', 'pg_primary', label='R/W')
    dot.edge('pg_primary', 'pg_standby', label='réplication\nsynchrone', style='dashed')
    dot.edge('pods_az1', 'redis')
    dot.edge('pods_az1', 'objstore')
    dot.edge('pg_primary', 'pitr', label='snapshot quotidien', style='dotted')
    dot.edge('pods_az1', 'audit_store', label='log immuable', style='dotted', color='#880E4F')

    path_base = os.path.join(tempfile.gettempdir(), 'deployment_defr')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_cicd_pipeline():
    dot = Digraph('cicd_defr')
    dot.attr(rankdir='LR', nodesep='0.35', ranksep='0.55', dpi='200',
             fontname='Helvetica', bgcolor='white')

    node_box = {'shape': 'box', 'style': 'filled,rounded', 'fontname': 'Helvetica',
                'fontsize': '8.3', 'margin': '0.15,0.1'}

    with dot.subgraph(name='cluster_ci') as c:
        c.attr(label='Intégration continue', style='rounded,filled', fillcolor='#E8F5E9',
               color='#2E7D32', fontname='Helvetica', fontsize='9')
        c.node('commit', 'Commit\n(branche de\nfonctionnalité)', fillcolor='#C8E6C9', **node_box)
        c.node('lint', 'Lint et tests\nunitaires', fillcolor='#C8E6C9', **node_box)
        c.node('sast', 'Analyse SAST et\nscan des dépendances', fillcolor='#A5D6A7', **node_box)
        c.node('build', 'Build image\nconteneurisée', fillcolor='#C8E6C9', **node_box)
        c.node('sign', 'Signature de l\'image\net registre sécurisé', fillcolor='#A5D6A7', **node_box)

    with dot.subgraph(name='cluster_cd') as c:
        c.attr(label='Déploiement continu', style='rounded,filled', fillcolor='#E3F2FD',
               color='#1565C0', fontname='Helvetica', fontsize='9')
        c.node('staging', 'Déploiement\nenvironnement de staging', fillcolor='#BBDEFB', **node_box)
        c.node('tests_charge', 'Tests d\'intégration\net de charge', fillcolor='#90CAF9', **node_box)
        c.node('validation', 'Validation manuelle\n(comité de suivi DEFR)', fillcolor='#90CAF9', **node_box)
        c.node('bluegreen', 'Bascule Blue-Green\nen production', fillcolor='#BBDEFB', **node_box)

    with dot.subgraph(name='cluster_obs') as c:
        c.attr(label='Exploitation', style='rounded,filled', fillcolor='#FFF3E0',
               color='#E65100', fontname='Helvetica', fontsize='9')
        c.node('monitoring', 'Supervision continue\n(SLA 99,95%)', fillcolor='#FFCC80', **node_box)
        c.node('rollback', 'Rollback automatique\nsi seuil d\'alerte dépassé', fillcolor='#FFAB91', **node_box)

    dot.edge('commit', 'lint')
    dot.edge('lint', 'sast')
    dot.edge('sast', 'build', label='si conforme')
    dot.edge('build', 'sign')
    dot.edge('sign', 'staging')
    dot.edge('staging', 'tests_charge')
    dot.edge('tests_charge', 'validation')
    dot.edge('validation', 'bluegreen', label='jalon contractuel')
    dot.edge('bluegreen', 'monitoring')
    dot.edge('monitoring', 'rollback', label='anomalie détectée', style='dashed', color='#C62828')
    dot.edge('rollback', 'bluegreen', label='retour version stable',
             style='dashed', color='#C62828', constraint='false')

    path_base = os.path.join(tempfile.gettempdir(), 'cicd_defr')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


def generate_gantt():
    activities = [
        ("Diagnostic des schémas ERP-HELISA et stratégie de migration ETL", 0.0, 0.8, "#15607F"),
        ("Conception détaillée (SFD/STD) et architecture n-tiers", 0.5, 1.0, "#15607F"),
        ("Développement du socle SIS (EFPA, filières, APC, vacations)", 1.5, 1.0, "#0F4761"),
        ("Module Concours et Admissions (blocage automatique)", 2.2, 1.0, "#0F4761"),
        ("Module Certification des diplômes (Barid eSign, QR Code)", 2.8, 0.8, "#0F4761"),
        ("Couche sémantique Headless BI", 3.0, 0.8, "#0F4761"),
        ("Orchestration RAG du Bot APC (LangChain.js, pgvector)", 3.2, 1.1, "#0F4761"),
        ("Migration technique des données historiques (5 ans)", 4.5, 0.8, "#1A7A9C"),
        ("Interconnexion API Moodle et configuration SSO Entra ID", 5.0, 0.8, "#1A7A9C"),
        ("Déploiement SaaS sur Cloud souverain - 58 EFPA", 5.6, 0.7, "#1A7A9C"),
        ("Accompagnement des administrateurs et formateurs", 6.3, 0.7, "#2E7D32"),
    ]

    milestones = [
        (1.5, "Livrable M1 : SFD/STD validés", 2.6),
        (4.5, "Livrable M2 : Modules métier opérationnels", 1.9),
        (6.3, "Livrable M3 : Mise en service sur 58 EFPA", 1.2),
    ]

    n = len(activities)
    fig, ax = plt.subplots(figsize=(19 / 2.54, 14.5 / 2.54))

    ax.axvspan(0, 1.5, alpha=0.07, color='#15607F', zorder=0)
    ax.axvspan(1.5, 4.5, alpha=0.07, color='#0F4761', zorder=0)
    ax.axvspan(4.5, 6.3, alpha=0.07, color='#1A7A9C', zorder=0)
    ax.axvspan(6.3, 7.0, alpha=0.07, color='#2E7D32', zorder=0)

    for i, (label, start, dur, color) in enumerate(activities):
        y = n - i
        ax.barh(y, dur, left=start, height=0.58, color=color, alpha=0.88,
                edgecolor='white', linewidth=0.5, zorder=2)

    # Dépendances et chemin critique : (index prédécesseur, index successeur)
    dependencies = [(0, 7), (2, 4), (2, 5), (2, 6), (7, 9), (8, 9), (9, 10)]
    for pred, succ in dependencies:
        y_pred = n - pred
        y_succ = n - succ
        x_end_pred = activities[pred][1] + activities[pred][2]
        x_start_succ = activities[succ][1]
        ax.annotate(
            '', xy=(x_start_succ, y_succ), xytext=(x_end_pred, y_pred),
            arrowprops=dict(arrowstyle='-|>', color='#C62828', lw=0.9,
                             linestyle=(0, (3, 2)),
                             connectionstyle='arc3,rad=0.12',
                             shrinkA=2, shrinkB=2),
            zorder=4,
        )

    for ms_x, ms_label, ms_y in milestones:
        ax.axvline(x=ms_x, color='#C17F00', linewidth=1.4, linestyle='--',
                   alpha=0.85, zorder=3)
        ax.plot(ms_x, n + ms_y, marker='D', color='#C17F00', markersize=5,
                zorder=5, clip_on=False)
        ax.text(ms_x + 0.12, n + ms_y, ms_label, fontsize=7.2,
                color='#C17F00', fontweight='bold', va='center', ha='left',
                zorder=6, clip_on=False)

    phase_labels = [
        (0.75, 'PHASE I', '#15607F', 3.7),
        (3.0, 'PHASE II', '#0F4761', 4.3),
        (5.4, 'PHASE III', '#1A7A9C', 3.7),
        (6.9, 'PHASE IV', '#2E7D32', 4.3),
    ]
    for px, ptext, pcolor, py in phase_labels:
        ax.text(px, n + py, ptext, fontsize=8, ha='center',
                color=pcolor, fontweight='bold', va='center', clip_on=False)

    ax.set_yticks(range(1, n + 1))
    ax.set_yticklabels([a[0] for a in reversed(activities)], fontsize=7.3)
    ax.set_xlim(0, 7.3)
    ax.set_xticks(range(0, 8))
    ax.set_xticklabels([f"M{m}" for m in range(0, 8)], fontsize=8)
    ax.set_xlabel("Mois depuis l'ordre de service", fontsize=8.5)
    ax.set_title("Planning d'exécution - 7 mois, 4 phases, avec dépendances (Art.20 CPS)",
                 fontsize=10, fontweight='bold', pad=18)

    dep_handle = plt.Line2D([0], [0], color='#C62828', lw=1.2, linestyle=(0, (3, 2)))
    ax.legend([dep_handle], ["Dépendance bloquante entre activités"],
              loc='lower right', fontsize=7.5, frameon=False)

    ax.set_ylim(0, n + 4.6)
    ax.spines[['top', 'right', 'left']].set_visible(False)
    ax.tick_params(left=False)
    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'gantt_defr.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return path


def generate_value_radar_chart():
    axes_labels = [
        "Sécurité\n(WAF, MFA, audit)",
        "Performance\n(SLA, Bot APC)",
        "Évolutivité\n(58 -> 100 EFPA)",
        "Interopérabilité\n(Moodle, M365)",
        "Réversibilité\n(code, transfert)",
        "Pilotage\n(BI multidimensionnel)",
    ]
    seuil_cps = [70, 70, 65, 70, 65, 65]
    niveau_abi = [92, 88, 95, 90, 93, 88]

    n = len(axes_labels)
    angles = [i / n * 2 * 3.14159265 for i in range(n)]
    angles += angles[:1]
    seuil_plot = seuil_cps + seuil_cps[:1]
    niveau_plot = niveau_abi + niveau_abi[:1]

    fig, ax = plt.subplots(figsize=(15 / 2.54, 13 / 2.54), subplot_kw=dict(polar=True))
    ax.set_theta_offset(3.14159265 / 2)
    ax.set_theta_direction(-1)

    ax.plot(angles, seuil_plot, linewidth=1.5, linestyle='--', color='#9E9E9E', label='Seuil exigé (CPS/RC)')
    ax.fill(angles, seuil_plot, color='#9E9E9E', alpha=0.15)

    ax.plot(angles, niveau_plot, linewidth=2, color='#15607F', label='Niveau proposé par ABI CONSULTING')
    ax.fill(angles, niveau_plot, color='#15607F', alpha=0.25)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(axes_labels, fontsize=8)
    ax.set_ylim(0, 100)
    ax.set_yticks([25, 50, 75, 100])
    ax.set_yticklabels(['25', '50', '75', '100'], fontsize=7, color='#777777')

    ax.set_title("Niveau proposé par rapport au seuil minimal du CPS/RC",
                 fontsize=10.5, fontweight='bold', y=1.1)
    ax.legend(loc='upper right', bbox_to_anchor=(1.32, 1.15), fontsize=7.5, frameon=False)

    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'value_radar_defr.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return path


# ---------------------------------------------------------------
# Sections du document
# ---------------------------------------------------------------

def page_de_garde(doc):
    pass


def section_sommaire(doc):
    p = doc.add_paragraph(style='Heading 1')
    r = p.add_run("Sommaire")
    r.font.name = 'Calibri'
    r.font.color.rgb = H1_COLOR
    r.font.size = Pt(14)
    r.font.bold = True
    add_toc_field(doc)
    doc.add_page_break()


def section_introduction(doc):
    p = doc.add_paragraph(style='Heading 1')
    r = p.add_run("Présentation de la note")
    r.font.name = 'Calibri'
    r.font.color.rgb = H1_COLOR
    r.font.size = Pt(14)
    r.font.bold = True

    add_body(doc,
        "La présente note méthodologique répond à l'Appel d'Offres Ouvert International N°01/2026/"
        "DEFR/DETFP relatif à la conception, au développement, à la mise en œuvre et à la mise en "
        "service en mode SaaS d'un système d'information apprenant intégré (SIS Intégré) au profit "
        "des 58 Établissements de Formation Professionnelle Agricole relevant de la Direction de "
        "l'Enseignement, de la Formation et de la Recherche."
    )

    add_body(doc,
        "Elle suit l'ordre de présentation fixé par l'Article 9 du Règlement de Consultation : "
        "compréhension du contexte et de la mission, approche méthodologique structurée selon les "
        "six points exigés par l'Article 9.1 (démarche de conduite de projet, stratégie ETL et "
        "migration, architecture SaaS et Cloud, interconnexion et SSO, approche IA, exploitation des "
        "données), architecture technique détaillée, équipe proposée, planning d'exécution, "
        "chronogramme d'affectation, plan de réversibilité, valeur ajoutée et gestion des risques. "
        "Chaque section reprend la terminologie exacte du CPS et du RC pour les exigences qu'elle "
        "adresse, et signale explicitement l'article correspondant."
    )
    doc.add_page_break()


def section_comprehension_contexte(doc, growth_path):
    add_heading(doc, "Compréhension du contexte", 1, numbered="1")

    add_body(doc,
        "La stratégie Génération Green 2020-2030, lancée sous l'impulsion royale, repose sur deux "
        "piliers : faire de l'élément humain une priorité de toute action de développement agricole, "
        "et poursuivre la dynamique de modernisation du secteur par l'émergence d'une nouvelle "
        "génération d'agriculteurs entrepreneurs. Elle vise notamment à doubler le PIB agricole en "
        "dix ans, à intégrer 350 000 jeunes dans le secteur d'ici 2030, et à porter les exportations "
        "agricoles entre 50 et 60 milliards de dirhams (Ministère de l'Agriculture, Génération Green "
        "2020-2030). La formation professionnelle agricole porte la composante humaine de cette "
        "ambition : depuis le lancement de la stratégie, 33 600 lauréats ont déjà été formés dans "
        "56 filières et spécialités, pour un objectif de 140 000 lauréats à l'horizon 2030 et un taux "
        "d'insertion professionnelle déjà observé de 65% (Département de l'Agriculture). Le SIS "
        "Intégré que ce marché doit produire est l'instrument qui permettra de mesurer, filière par "
        "filière et établissement par établissement, la progression réelle vers cette cible, plutôt "
        "qu'un projet informatique mené en parallèle de la stratégie."
    )

    add_body(doc,
        "Cette ambition s'appuie sur un réseau qui s'est étoffé au fil des dernières années pour "
        "s'organiser aujourd'hui en douze pôles multi-centres régionaux. Le périmètre de ce marché "
        "porte sur 58 de ces établissements, avec une extensibilité prévue jusqu'à 100, ce qui "
        "représente une croissance de 72% de la capacité du dispositif si elle se confirme. Avec une "
        "capacité de gestion minimale de 30 000 stagiaires actifs annoncée à l'Article 1 du CPS, "
        "chaque établissement compte en moyenne plus de 500 stagiaires, un chiffre qui rend "
        "impraticable toute consolidation manuelle des indicateurs de réussite, d'attraction, "
        "d'absentéisme et d'insertion à l'échelle nationale."
    )
    add_figure(doc, growth_path, "Extensibilité du dispositif EFPA : aujourd'hui et à la cible visée par le CPS")

    add_body(doc,
        "Cette trajectoire de croissance n'est pas isolée : elle rejoint un mouvement plus large de "
        "digitalisation de la formation professionnelle et technique, que l'UNESCO-UNEVOC documente "
        "depuis plusieurs années sous le terme de système d'information de gestion de l'éducation "
        "(EMIS). Le rapport mondial de suivi de l'éducation de 2023 y décrit la technologie numérique "
        "comme à la fois un intrant, un canal de diffusion, une compétence à acquérir et un outil de "
        "planification, une description qui correspond précisément aux quatre rôles que le SIS "
        "Intégré doit remplir pour la DEFR : collecter, diffuser, faire monter en compétence les "
        "équipes, et piloter."
    )

    add_body(doc,
        "Le dispositif actuel repose sur l'ERP HELISA et sur des pratiques de gestion hétérogènes "
        "selon les établissements. Les données de scolarité des cinq dernières années existent, mais "
        "sous des formats et des niveaux de qualité variables d'un EFPA à l'autre, ce qui rend leur "
        "exploitation centralisée difficile en l'état. La DEFR a par ailleurs déjà investi dans deux "
        "écosystèmes numériques actifs, le LMS Moodle pour la pédagogie et Microsoft 365 Education "
        "pour la collaboration et l'identité. Ce choix n'est pas isolé au sein du système éducatif "
        "marocain : le ministère de l'Éducation Nationale a lui-même engagé dès 2013 un partenariat "
        "avec Microsoft portant sur la migration de son portail Taalim.ma vers Office 365, au "
        "bénéfice de 7,5 millions d'utilisateurs, puis sur l'adoption de Microsoft Teams pour les "
        "classes virtuelles. Le SIS Intégré s'inscrit donc dans un écosystème technologique déjà "
        "éprouvé à l'échelle nationale, et non dans un choix isolé propre à la DEFR."
    )

    add_body(doc,
        "Trois enjeux structurent ce contexte. Le premier est la fraude documentaire sur les "
        "diplômes : des affaires récentes de commercialisation de diplômes contre rémunération, "
        "révélées au Maroc dans l'enseignement supérieur, rappellent que la valeur d'un diplôme "
        "dépend directement de la capacité à le vérifier de manière incontestable. La loi 43-20 "
        "relative aux services de confiance pour les transactions électroniques, en vigueur depuis "
        "le 13 juillet 2023 et alignée sur le règlement européen eIDAS, donne précisément à la "
        "signature électronique qualifiée la valeur juridique nécessaire pour que la certification "
        "numérique des diplômes du SIS Intégré ne soit pas qu'une fonctionnalité technique mais une "
        "preuve opposable. Le second enjeu est l'absence de vision consolidée des indicateurs à "
        "l'échelle des 58 établissements, qui limite aujourd'hui le pilotage national par la DEFR à "
        "des remontées manuelles et asynchrones. Le troisième est la fragilité de l'historique de "
        "scolarité, qui doit être repris dans son intégralité avant que les nouveaux usages "
        "numériques ne s'y superposent."
    )

    add_body(doc,
        "Ce contexte s'inscrit enfin dans un cadre réglementaire devenu plus strict depuis 2024. La "
        "loi 05-20 relative à la cybersécurité confie à la Direction Générale de la Sécurité des "
        "Systèmes d'Information (DGSSI) la qualification des prestataires Cloud, et le décret "
        "n°2-24-921 d'octobre 2024 distingue désormais deux niveaux de qualification : un premier "
        "niveau qui autorise certaines opérations de support depuis l'étranger via une passerelle "
        "sécurisée, et un second niveau qui impose que les données sensibles soient stockées, "
        "traitées et administrées exclusivement depuis le territoire marocain, avec audit annuel "
        "obligatoire sur plus de quinze domaines de conformité. La loi 09-08 relative à la protection "
        "des données à caractère personnel impose en complément une conformité CNDP démontrable pour "
        "un système hébergeant les données de plus de 30 000 stagiaires actifs."
    )


def section_comprehension_mission(doc, objectifs_path):
    add_heading(doc, "Compréhension de la mission", 1, numbered="2")

    add_heading(doc, "Objectifs et livrables", 2)

    add_body(doc,
        "L'objectif général du marché, tel que formulé à l'Article 3 du CPS, est de doter les 58 "
        "Établissements de Formation Professionnelle Agricole d'un système d'information moderne, "
        "performant, évolutif et conforme aux standards internationaux, permettant d'améliorer la "
        "qualité de la formation, la gestion des établissements et le pilotage global du dispositif. "
        "Le schéma ci-dessous traduit chacun des quatre objectifs spécifiques de l'Article 3 en un "
        "mécanisme concret du SIS Intégré et en un bénéfice mesurable pour la DEFR, plutôt que de "
        "les traiter comme une liste d'intentions générales."
    )
    add_figure(doc, objectifs_path, "Des objectifs du CPS aux bénéfices mesurables pour la DEFR")

    add_body(doc,
        "Le SIS Intégré devra également couvrir la gestion harmonisée des admissions et concours, la "
        "gestion complète de la scolarité et des parcours pédagogiques, la gestion des évaluations, "
        "examens et délibérations selon les normes pédagogiques en vigueur, la planification et le "
        "suivi des activités de formation, la gestion des stages et de l'insertion des lauréats, et la "
        "production automatisée des documents officiels. Le suivi de l'insertion professionnelle "
        "n'est pas une fonctionnalité accessoire : avec un taux d'insertion déjà observé de 65% sur "
        "le dispositif national de formation professionnelle agricole, disposer d'une base de "
        "contacts structurée et d'un suivi systématique, comme l'exige l'Article 19.5.6 du CPS, "
        "conditionne la capacité de la DEFR à documenter et à améliorer ce taux plutôt qu'à le "
        "constater après coup. L'intégration avec Moodle et Microsoft 365 Education conditionne par "
        "ailleurs directement l'adoption de la plateforme par les enseignants, qui utilisent déjà ces "
        "deux outils au quotidien."
    )

    headers = ["Phase", "Livrables (Art.21 CPS)"]
    rows = [
        ["Phase I\nDiagnostic et conception",
         "Diagnostic des schémas ERP existant et rapport de stratégie de migration ETL ; "
         "Dossier de Conception Détaillée (SFD/STD) et maquettes UX/UI validées par le comité de suivi"],
        ["Phase II\nDéveloppement et mise en œuvre métier",
         "Dossier de paramétrage et socle applicatif (58 EFPA, filières, délibérations APC) ; "
         "module Recrutement ; module Certification (Barid eSign, QR) ; plateforme Headless BI ; "
         "dossier d'orchestration RAG du Bot APC"],
        ["Phase III\nIntégration, déploiement et mise en service",
         "PV de migration de données ; certificat d'interconnexion Microsoft (SSO Entra ID, sync Moodle) ; "
         "attestation de mise en service du Cloud souverain ; rapport de déploiement final et code source"],
        ["Phase IV\nAccompagnement",
         "Manuels utilisateurs et enregistrements vidéo des sessions de transfert de compétences"],
    ]
    add_table(doc, headers, rows, caption="Tableau 1. Livrables contractuels par phase",
              col_widths=[4.0, 11.0])

    add_heading(doc, "Périmètre et limites de la mission", 2)

    add_body(doc,
        "Le périmètre couvre l'intégralité du cycle de vie de l'apprenant, de la candidature à la "
        "diplomation, ainsi que le pilotage transverse par la DEFR sur l'ensemble des 58 "
        "établissements, extensible à 100. Il inclut la migration de cinq années d'historique de "
        "scolarité et la mise en service opérationnelle en mode SaaS, hébergée sur un Cloud souverain "
        "qualifié au sens du référentiel DGSSI de 2024. Il exclut explicitement la conception, le "
        "développement et le paramétrage métier du SIS de la sous-traitance, ce corps d'état "
        "principal devant rester sous la responsabilité directe de l'équipe projet."
    )

    add_body(doc,
        "La limite la plus significative du périmètre concerne l'interopérabilité : le SIS Intégré "
        "s'interface avec Moodle et Microsoft 365 Education tels qu'ils existent aujourd'hui à la "
        "DEFR, sans engagement de migration ou de remplacement de ces deux systèmes. De même, la "
        "qualité des données reprises de l'ERP HELISA dépend des extraits effectivement disponibles "
        "et fournis par le maître d'ouvrage au titre de l'Article 7 du CPS ; la responsabilité du "
        "titulaire porte sur l'intégrité du processus de migration, pas sur la complétude des "
        "données sources antérieures à cette migration."
    )


def section_approche_methodologique(doc):
    add_heading(doc, "Approche méthodologique", 1, numbered="3")

    add_body(doc,
        "Cette section traite, dans l'ordre fixé par l'Article 9.1 du Règlement de Consultation, "
        "les six points exigés pour la note méthodologique de travail : la démarche de conduite de "
        "projet, la stratégie ETL et de migration, l'architecture SaaS et Cloud, l'interconnexion et "
        "le SSO, l'approche IA, et l'exploitation des données."
    )

    # --- A. Démarche de conduite de projet ---
    add_heading(doc, "Démarche de conduite de projet", 2)

    add_body(doc,
        "La conduite de projet suit le Cycle en V préconisé par le Maître d'Ouvrage pour le cadrage "
        "et la conception, complété par des itérations Agile pour le développement des modules "
        "fonctionnels. Le Cycle en V structure les quatre phases contractuelles : diagnostic et "
        "conception, développement et mise en œuvre métier, intégration et déploiement, "
        "accompagnement. Chaque phase se referme par un jalon de validation formel avec la DEFR avant "
        "que la phase suivante ne démarre, ce qui évite qu'une non-conformité détectée tardivement "
        "ne remette en cause des développements déjà engagés."
    )

    add_heading(doc, "Diagnostic et analyse des besoins", 3)

    add_body(doc,
        "La Phase I ne se limite pas, dans son exécution, au diagnostic des schémas de données "
        "annoncé à l'Article 19.1.1 du CPS. Elle démarre par un atelier de cadrage de deux jours "
        "réunissant la DEFR et un échantillon représentatif d'EFPA couvrant plusieurs profils "
        "d'établissement (effectif, filières dominantes, niveau d'équipement numérique actuel), "
        "afin que les choix de paramétrage du socle applicatif ne reposent pas uniquement sur la "
        "lecture du CPS mais sur la confrontation directe avec la diversité réelle des 58 "
        "établissements. Cet atelier produit une grille de recueil des besoins organisée par bloc "
        "fonctionnel de l'Article 19.5 du CPS, remplie avec les agents EFPA présents plutôt que "
        "déduite a priori par l'équipe projet."
    )

    add_body(doc,
        "Le diagnostic des schémas de données de l'ERP HELISA, livrable contractuel de la Phase I, "
        "est conduit comme une démarche d'audit avant d'être un document : extraction d'un "
        "échantillon de données par établissement, identification des champs obligatoires non "
        "renseignés, des doublons d'identité d'apprenant, et des divergences de codification des "
        "filières entre EFPA. Cette démarche d'audit alimente directement la cartographie de "
        "correspondance avec le modèle cible du SIS Intégré, et conditionne la stratégie de "
        "migration ETL détaillée au point B de la présente section."
    )

    add_body(doc,
        "Les maquettes UX/UI des portails Apprenants et Enseignants, livrable de la Phase I au titre "
        "de l'Article 21.1 du CPS, sont produites par prototypage itératif avec des utilisateurs "
        "réels désignés par la DEFR, plutôt que validées sur la seule base d'une revue interne à "
        "l'équipe projet. Chaque maquette est confrontée à un scénario d'usage concret tiré de la "
        "grille de recueil des besoins (par exemple la saisie d'une note de concours sur un poste "
        "partagé entre plusieurs agents d'un même EFPA) avant d'être soumise au comité de suivi pour "
        "validation."
    )

    add_body(doc,
        "À l'intérieur de la Phase II, le développement des modules métier (admissions, scolarité, "
        "certification, BI, Bot APC) est conduit en itérations de deux semaines, chacune se terminant "
        "par une démonstration au comité de suivi. Cette organisation permet de détecter un écart "
        "d'interprétation fonctionnelle sur un module avant qu'il ne se propage aux modules qui en "
        "dépendent, en particulier le module Parcours Apprenant dont la couverture est la plus large."
    )

    add_body(doc,
        "La démarche d'assurance qualité s'appuie sur le référentiel ISO 25010 : chaque livrable "
        "logiciel est évalué selon les caractéristiques de fiabilité, de performance, de sécurité, "
        "de maintenabilité et de compatibilité avant sa présentation au comité de pilotage. Cette "
        "évaluation se traduit concrètement par une couverture de tests automatisés minimale par "
        "service, des tests de charge avant chaque mise en service partielle, et une revue de code "
        "systématique avant fusion sur la branche principale."
    )

    add_body(doc,
        "La gestion des livrables suit la liste exacte de l'Article 21 du CPS : chaque livrable "
        "donne lieu à un document de remise daté, soumis au comité de suivi désigné par la DEFR, et "
        "fait l'objet d'un procès-verbal de réception partielle avant que le décompte de paiement "
        "correspondant ne soit déposé sur la plateforme GID fournisseurs."
    )

    add_heading(doc, "Couverture des exigences fonctionnelles (Art.19.5 CPS)", 2)

    add_body(doc,
        "Le tableau ci-dessous reprend, bloc par bloc, les onze exigences fonctionnelles détaillées "
        "à l'Article 19.5 du CPS, dans leur intitulé exact, et indique le service porteur défini à "
        "la section Architecture SaaS et Cloud de la présente note. Cette correspondance explicite "
        "garantit qu'aucune exigence fonctionnelle du CPS ne reste sans réponse architecturale "
        "identifiée."
    )

    headers_fn = ["Bloc fonctionnel (Art.19.5 CPS)", "Exigences clés", "Service porteur"]
    rows_fn = [
        ["19.5.1 Fonctions transverses",
         "Paramétrage des EFPA (infrastructures, capacités), gestion des filières, modules et "
         "volumes horaires, ressources pédagogiques, enseignants et groupes de stagiaires",
         "Référentiel"],
        ["19.5.2 Admissions et concours",
         "Réception des dossiers, sélection automatique/manuelle, planification et convocation, "
         "saisie des notes, délibération, listes d'attente, inscriptions définitives",
         "Parcours Apprenant"],
        ["19.5.3 Scolarité et examens",
         "Inscription automatique des admis, réinscriptions, suivi de l'assiduité, discipline et "
         "sanctions, abandons et changements de filière",
         "Parcours Apprenant"],
        ["19.5.4 Plannings de formation",
         "Emplois du temps annuels, occupation des salles et des enseignants, feuilles de service, "
         "suivi des heures de vacation",
         "Planning & Ressources"],
        ["19.5.5 Évaluation et diplomation",
         "Déclenchement de l'évaluation, saisie et déclaration des résultats, reprises "
         "d'évaluation, diplomation",
         "Parcours Apprenant"],
        ["19.5.6 Stages et soutenances",
         "Banque d'entreprises de stage, affectation des stagiaires, suivi de l'insertion "
         "professionnelle",
         "Stages & Insertion"],
        ["19.5.7 Extranet et services en ligne",
         "Espace Stagiaire, Espace Enseignant/Référent, Espace Administration",
         "Gateway (agrégation)"],
        ["19.5.8 Pilotage et reporting",
         "Couche sémantique Headless BI, analyse multidimensionnelle (réussite, attraction, "
         "absentéisme, insertion)",
         "BI & Reporting"],
        ["19.5.9 Services d'intelligence artificielle",
         "Bot APC en RAG, indexation vectorielle des référentiels pédagogiques",
         "Bot APC / IA"],
        ["19.5.10 Édition et certification des diplômes",
         "Certification numérique par signature électronique qualifiée, génération de QR Code",
         "Certification"],
        ["19.5.11 Journalisation et sécurité",
         "Journalisation immuable de l'identité, de la date, de l'heure et de la nature de toute "
         "action, conservation 12 mois",
         "Audit Trail Store (transverse)"],
    ]
    add_table(doc, headers_fn, rows_fn, caption="Tableau 2. Correspondance entre les exigences fonctionnelles du CPS et les services",
              col_widths=[4.5, 8.0, 3.0])

    add_body(doc,
        "Trois blocs méritent une précision au-delà de la correspondance tabulaire. La gestion des "
        "admissions et concours intègre le blocage automatique des convocations explicitement exigé "
        "à l'Article 19.1.2 du CPS pour le module Concours : un candidat dont le dossier n'a pas "
        "franchi l'étape de présélection ne peut recevoir de convocation, par contrainte applicative "
        "et non par seule consigne opérationnelle. L'extranet distingue les trois espaces du CPS par "
        "des droits d'écriture strictement différenciés : l'Espace Stagiaire reste en lecture sur les "
        "modules, notes, planning et absences, avec demande d'attestation comme seul acte d'écriture "
        "possible ; l'Espace Enseignant ajoute la publication de cours et de supports ; l'Espace "
        "Administration seul peut publier mémos, circulaires et programmes d'examens. Les stages "
        "s'appuient sur une banque d'entreprises gérée comme un référentiel à part, alimentée à la "
        "fois par les EFPA et par la DEFR, pour que le suivi de l'insertion professionnelle des "
        "lauréats ne dépende pas de la seule saisie d'un seul établissement."
    )

    # --- B. Stratégie ETL et Migration ---
    add_heading(doc, "Stratégie ETL et migration des données historiques", 2)

    add_body(doc,
        "La migration des cinq années d'historique de scolarité n'est pas traitée comme un endpoint "
        "applicatif appelé à la demande, mais comme un job de traitement par lots, exécuté en dehors "
        "du chemin de requête de l'application en service. Cette séparation évite qu'un traitement "
        "volumineux sur 58 sources hétérogènes ne dégrade la disponibilité du système pour les "
        "utilisateurs actifs pendant la période de migration."
    )

    add_body(doc,
        "Le processus se déroule en trois étapes. Le diagnostic des schémas de données de l'ERP "
        "HELISA produit une cartographie de correspondance avec le modèle cible du SIS Intégré, "
        "identifiant les tables critiques et les écarts de structure entre établissements. "
        "L'extraction et le nettoyage appliquent des règles de transcodage par EFPA, dès lors que "
        "l'hétérogénéité documentée à la section précédente implique que deux établissements ne "
        "codifient pas nécessairement leurs filières ou leurs statuts d'inscription de la même "
        "manière. L'injection dans PostgreSQL s'effectue par lots contrôlés, chaque lot étant validé "
        "par un comptage de lignes et un calcul de somme de contrôle avant son intégration définitive."
    )

    add_body(doc,
        "Toute anomalie détectée lors de la vérification d'un lot suspend l'injection de ce lot sans "
        "affecter les lots déjà validés, et génère une fiche d'anomalie horodatée transmise au comité "
        "de suivi. Le rapport de stratégie de migration ETL, livrable contractuel de la Phase I, "
        "documente ces règles de nettoyage et de transcodage avant le démarrage de l'extraction, et "
        "le PV de migration de données, livrable de la Phase III, certifie l'intégrité du résultat "
        "une fois l'ensemble des 58 EFPA migré."
    )

    # --- C. Architecture SaaS et Cloud ---
    add_heading(doc, "Architecture SaaS et Cloud", 2)

    add_body(doc,
        "L'architecture applicative repose sur Node.js en version LTS, conformément à l'Article "
        "19.3.2 du CPS, organisé selon le framework NestJS. Ce choix n'est pas seulement une "
        "implémentation du runtime imposé : la structure modulaire de NestJS correspond directement "
        "à l'exigence d'un système conçu par blocs fonctionnels indépendants formulée à l'Article "
        "19.3.1, et son support natif de TypeScript apporte une sûreté de typage déterminante pour la "
        "réversibilité totale du code exigée en fin de marché."
    )

    add_body(doc,
        "Le système est découpé en neuf services déployables, regroupés par cohésion métier plutôt "
        "que par correspondance mécanique avec chacun des onze blocs fonctionnels de l'Article 19.5 "
        "du CPS. Le Référentiel porte les données maîtres des EFPA, filières, modules et ressources "
        "pédagogiques. Le Parcours Apprenant regroupe volontairement les admissions et concours, la "
        "scolarité et l'évaluation-diplomation, parce que ces trois blocs opèrent sur la même entité "
        "pivot, le dossier de l'apprenant, et que les séparer créerait des transactions distribuées "
        "permanentes sur la même donnée. Planning et Ressources, Stages et Insertion, Certification, "
        "BI et Reporting, et Bot APC complètent le dispositif, chacun isolé parce qu'il porte une "
        "logique distincte : résolution de contraintes pour le planning, appel à une PKI externe pour "
        "la certification, séparation OLTP/OLAP pour le BI."
    )

    add_body(doc,
        "La persistance s'appuie sur PostgreSQL avec l'extension pgvector dans la même instance, "
        "ORM Prisma et Supabase ou équivalent auto-hébergé pour l'authentification bas niveau et le "
        "stockage objet, conformément à l'Article 19.3.2 du CPS. La stratégie multi-tenant retenue "
        "pour les 58 établissements, extensibles à 100, repose sur un schéma partagé avec sécurité au "
        "niveau ligne (Row-Level Security) discriminée par l'identifiant d'établissement, plutôt que "
        "sur une base par établissement : la DEFR a besoin d'agrégats inter-EFPA pour son pilotage "
        "national, ce qu'une base par tenant rendrait coûteux à requêter à l'échelle visée."
    )

    add_body(doc,
        "Tout traitement susceptible de bloquer la boucle d'événements de Node.js, en particulier la "
        "certification des diplômes et l'envoi de notifications en masse après une délibération, est "
        "mis en file via BullMQ adossé à Redis. La requête API répond immédiatement et un worker "
        "dédié traite le travail en arrière-plan, ce qui permet de dimensionner indépendamment les "
        "répliques de l'API et les workers de certification, conformément à l'exigence de scalabilité "
        "horizontale automatique de l'Article 19.4.1 du CPS. L'infrastructure cible répond aux "
        "spécifications Premium de l'Article 19.4.2 : instances de calcul dédiées de type XL/2XL, "
        "base de données dimensionnée à 64 Go de RAM sur SSD NVMe à hautes performances, et unité de "
        "calcul GPU pour l'inférence du Bot APC, l'ensemble hébergé exclusivement sur un Cloud "
        "souverain qualifié au Maroc."
    )

    add_body(doc,
        "Cette exigence de souveraineté n'est plus une simple préférence depuis le décret "
        "n°2-24-921 d'octobre 2024, qui structure la qualification des prestataires Cloud par la "
        "DGSSI en deux niveaux. Le SIS Intégré, qui héberge des données de scolarité à caractère "
        "personnel sur plus de 30 000 stagiaires, relève du niveau le plus strict : stockage, "
        "traitement et administration exclusivement depuis le territoire marocain, sans passerelle "
        "de support depuis l'étranger. Le choix du prestataire Cloud sous-traitant qualifié, prévu à "
        "l'Article 19.4 du CPS, sera donc arrêté sur la base de cette qualification de niveau 2 et "
        "de l'audit annuel obligatoire qui l'accompagne, et non sur la seule disponibilité d'une "
        "offre Cloud locale."
    )

    # --- D. Interconnexion et SSO ---
    add_heading(doc, "Interconnexion et SSO", 2)

    add_body(doc,
        "L'authentification unique repose sur Microsoft Entra ID en OpenID Connect, avec "
        "authentification multi-facteur obligatoire pour les comptes administrateurs et "
        "enseignants. La validation des jetons s'effectue au niveau de la passerelle API, qui "
        "résout également le contexte d'établissement et de rôle de l'utilisateur avant de router la "
        "requête vers le service métier concerné. Aucun service métier ne réimplémente sa propre "
        "logique d'authentification : la passerelle est le point unique de vérification, ce qui "
        "garantit une politique de sécurité homogène sur l'ensemble des neuf services."
    )

    add_body(doc,
        "La synchronisation avec le LMS Moodle est bidirectionnelle : les groupes de stagiaires et "
        "les filières définis dans le Référentiel sont propagés vers Moodle pour la création des "
        "cours correspondants, et les données d'activité pédagogique pertinentes sont consolidées en "
        "retour dans le SIS Intégré pour alimenter le tableau de bord de l'enseignant. Cette "
        "synchronisation est conçue de façon idempotente, afin qu'une réexécution après incident ne "
        "produise pas de doublons de groupes ou d'inscriptions."
    )

    add_body(doc,
        "Le risque d'adoption associé à l'authentification unique via l'écosystème Microsoft est "
        "limité par un précédent national déjà documenté : le ministère de l'Éducation Nationale a "
        "migré son portail Taalim.ma vers Microsoft 365 dès 2013, au bénéfice de 7,5 millions "
        "d'utilisateurs, avant d'adopter Microsoft Teams pour les classes virtuelles. Les "
        "enseignants des EFPA qui rejoignent le dispositif ne découvrent donc pas un environnement "
        "Microsoft inconnu du système éducatif marocain, ce qui réduit le risque d'adoption identifié "
        "pour ce module à un risque de configuration technique plutôt qu'à un risque d'acceptation "
        "par les utilisateurs."
    )

    # --- E. Approche IA ---
    add_heading(doc, "Approche IA - Bot APC", 2)

    add_body(doc,
        "L'orchestration du Bot APC s'appuie sur LangChain.js, conformément au livrable nommé à "
        "l'Article 21.2 du CPS. Le moteur d'inférence du modèle de langage, déployé sur l'unité GPU "
        "dédiée, constitue une appliance distincte de ce code d'orchestration : aucun runtime "
        "d'inférence GPU mature n'existe en JavaScript natif, et le déployer comme composant "
        "d'infrastructure packagé, exposé par une API HTTP compatible OpenAI, n'affecte pas la "
        "conformité du livrable contractuel puisque l'orchestration RAG elle-même reste entièrement "
        "écrite en LangChain.js."
    )

    add_body(doc,
        "La base de connaissances indexée dans pgvector couvre les référentiels d'Activités et de "
        "Compétences de l'Approche Par Compétences, les Guides d'Organisation Pédagogique et "
        "Matérielle, et les plans de modules fournis par la DEFR au titre de l'Article 7.4 du CPS. "
        "La stratégie d'indexation découpe chaque référentiel par unité pédagogique cohérente, plutôt "
        "que par taille de fragment fixe, afin que le contexte renvoyé au modèle pour une question "
        "donnée correspond à une unité de sens complète plutôt qu'à un extrait tronqué."
    )

    add_body(doc,
        "L'objectif de temps de réponse inférieur à 1,5 seconde, fixé à l'Article 19.4.2 du CPS, est "
        "tenu par la combinaison du dimensionnement GPU dédié et d'une mise en cache des réponses aux "
        "questions les plus fréquemment posées par profil d'utilisateur, ce cache étant invalidé "
        "automatiquement à chaque mise à jour d'un référentiel pédagogique source."
    )

    add_body(doc,
        "Le choix d'un agent conversationnel adossé à un référentiel documentaire, plutôt qu'un "
        "moteur de recherche classique sur les guides APC, suit une évolution déjà engagée dans la "
        "formation professionnelle : les organismes de formation déploient depuis 2025 des agents "
        "IA pour traiter la volumétrie des questions administratives et pédagogiques récurrentes, ce "
        "qui libère les formateurs et les agents EFPA pour l'accompagnement des cas qui demandent un "
        "jugement humain. Le Bot APC reprend cette logique pour les questions sur l'Approche Par "
        "Compétences : il absorbe les questions répétitives sur un référentiel ou un module, et "
        "laisse aux équipes pédagogiques les questions d'interprétation qui ne se résument pas à une "
        "recherche documentaire."
    )

    # --- F. Exploitation des données ---
    add_heading(doc, "Exploitation, sécurisation et performance des données", 2)

    add_body(doc,
        "Le reporting et le pilotage par indicateurs (réussite, attraction, absentéisme, insertion) "
        "ne sont jamais servis par requête directe sur les bases transactionnelles. Une couche "
        "sémantique de type Headless BI est alimentée par réplication depuis les services métier, ce "
        "qui isole les requêtes analytiques, potentiellement lourdes lors des périodes de "
        "consolidation, du service rendu aux apprenants et enseignants en période de pointe, "
        "notamment lors des résultats de concours."
    )

    add_body(doc,
        "La sécurisation des données suit le Pack Sécurité Avancé de l'Article 19.4.4 du CPS : pare-feu "
        "applicatif contre les injections et les attaques par force brute sur le SSO, protection "
        "contre le déni de service distribué, authentification multi-facteur obligatoire, et "
        "journalisation immuable de l'identité, de la date, de l'heure et de la nature de toute "
        "action sur une rétention de douze mois glissants. Cette journalisation est centralisée par "
        "un intercepteur appliqué uniformément à chacun des neuf services, plutôt que reconstruite "
        "service par service, ce qui garantit l'homogénéité du format d'audit exigée pour une "
        "exploitation effective en cas d'incident."
    )

    add_body(doc,
        "La conformité CNDP au titre de la loi 09-08 est intégrée dès la conception : les mentions "
        "légales relatives au traitement des données personnelles des stagiaires figurent dans "
        "l'extranet, et le titulaire assiste la DEFR dans le dépôt de la demande d'autorisation de "
        "traitement auprès de la Commission Nationale de contrôle de la protection des Données à "
        "caractère Personnel avant la mise en service opérationnelle."
    )


def section_architecture(doc, arch_path, usecase_path, classes_path, sequence_path,
                          deployment_path, cicd_path):
    add_heading(doc, "Architecture technique détaillée", 1, numbered="4")

    add_heading(doc, "Architecture globale de la plateforme", 2)

    add_body(doc,
        "Le schéma ci-dessous représente l'ensemble des neuf services, la passerelle API, et les "
        "composants d'infrastructure transverses décrits à la section précédente. La passerelle ne "
        "porte aucune logique métier : elle valide les jetons Microsoft Entra ID, résout le contexte "
        "d'établissement et de rôle, route la requête, et alimente l'audit immuable. Les trois "
        "espaces de l'extranet (Stagiaire, Enseignant, Administration) sont des applications client "
        "qui consomment les services par cette passerelle, et ne possèdent aucune donnée propre."
    )
    add_figure(doc, arch_path, "Architecture globale de la plateforme SIS Intégré")

    add_body(doc,
        "Deux alternatives ont été écartées avant de retenir ce découpage. Exposer directement les "
        "neuf services aux trois espaces extranet, sans passerelle unique, aurait obligé chaque "
        "service à revalider lui-même les jetons Entra ID et à journaliser l'audit selon sa propre "
        "logique, ce qui aurait rendu impossible la garantie d'un format d'audit homogène exigée à "
        "l'Article 19.3.5 du CPS. Scinder davantage le Parcours Apprenant en services distincts pour "
        "les admissions, la scolarité et la diplomation, par symétrie stricte avec les blocs "
        "fonctionnels du CPS, aurait introduit des transactions distribuées sur le même dossier "
        "apprenant à chaque changement de filière ou réinscription : la cohérence transactionnelle "
        "sur une entité pivot unique a été jugée plus importante que la symétrie formelle avec la "
        "numérotation de l'Article 19.5."
    )

    add_heading(doc, "Comparaison et justification des choix technologiques", 2)

    add_body(doc,
        "Le schéma d'architecture globale ci-dessus nomme des technologies précises sur chaque "
        "composant. Cette sous-section défend, pour les quatre décisions les plus structurantes, le "
        "raisonnement technique qui a écarté chaque alternative sérieusement examinée, et pas "
        "seulement le choix retenu présenté comme une évidence."
    )

    add_heading(doc, "Runtime et framework applicatif", 3)

    add_body(doc,
        "La vraie question n'est pas de savoir si .NET ou Spring Boot sont des technologies capables "
        "de porter ce projet : elles le sont, et le seraient sans doute aussi bien que Node.js sur le "
        "plan de la robustesse pure. La question est de savoir si l'écart de capacité théorique "
        "justifie le coût réel d'introduire un cinquième langage dans une équipe de quatre experts, "
        "sur un délai de sept mois, alors que le CPS impose déjà Node.js comme runtime de référence "
        "à l'Article 19.3.2. Dévier de cette prescription suppose, aux termes du CPS lui-même, une "
        "justification de robustesse institutionnelle : c'est un chemin ouvert, mais risqué, puisque "
        "le Critère A de notation récompense explicitement la conformité aux termes de référence. "
        "Spring Boot illustre bien ce dilemme. Son injection de dépendances et son écosystème "
        "d'observabilité sont matures, mais la JVM impose un cycle de compilation et un réglage du "
        "ramasse-miettes qui ralentissent les itérations de deux semaines prévues à la section 3.1, et "
        "son équivalent à LangChain.js pour l'orchestration RAG, Spring AI, reste un projet jeune par "
        "rapport au livrable nommé à l'Article 21.2 du CPS. .NET pose un problème différent : ses SDK "
        "Microsoft pour Entra ID et Graph sont natifs et excellents, mais Node.js dispose des mêmes "
        "SDK officiels (MSAL Node, Microsoft Graph SDK for JavaScript) sans qu'il soit nécessaire de "
        "changer de runtime pour les obtenir, ce qui retire le seul argument qui aurait pu justifier "
        ".NET. Quant à FastAPI, c'est le choix le plus défendable sur le papier pour la seule couche "
        "IA, et c'est précisément pourquoi il est conservé, mais comme moteur de service du modèle "
        "(voir plus loin), jamais comme stack applicative : l'adopter pour le reste du système aurait "
        "dupliqué l'intégration SSO dans un second langage sans SDK Microsoft de même maturité, pour "
        "un gain nul sur le reste du périmètre fonctionnel. NestJS, enfin, n'est pas qu'une commodité : "
        "sa structure modulaire est la seule des quatre options qui traduit directement, sans "
        "adaptation, l'exigence de blocs fonctionnels indépendants de l'Article 19.3.1."
    )
    headers_a = ["Critère", "Node.js + NestJS (retenu)", ".NET / ASP.NET Core", "Spring Boot (Java)", "FastAPI (Python)"]
    rows_a = [
        ["Conformité Art.19.3.2 CPS", "Imposé explicitement", "Équivalent à justifier", "Équivalent à justifier", "Équivalent à justifier"],
        ["Livrable LangChain.js (Art.21.2 CPS)", "Langage natif du livrable nommé", "Pont inter-langage nécessaire", "Pont inter-langage nécessaire", "Langage de LangChain, mais pas du livrable nommé"],
        ["Modularité par blocs indépendants (Art.19.3.1)", "Native (modules + injection de dépendances)", "Possible, plus de code répétitif", "Native, mais JVM plus lourde à opérer", "Modularité laissée à la charge du développeur"],
        ["Compatible avec l'équipe de 4 experts mandatée", "Oui, aucun profil supplémentaire requis", "Nécessiterait une expertise .NET dédiée absente de l'équipe", "Nécessiterait une expertise JVM dédiée absente de l'équipe", "Aucun profil Python dans l'équipe mandatée par le RC"],
        ["Verdict", "Retenu", "Écarté : aucun gain, risque d'équipe", "Écarté : délai de 7 mois incompatible avec la courbe d'apprentissage", "Écarté comme stack applicative ; conservé uniquement pour le moteur d'inférence"],
    ]
    add_table(doc, headers_a, rows_a, caption="Tableau 3. Comparaison des runtimes et frameworks backend envisagés",
              col_widths=[4.2, 3.0, 3.0, 3.0, 3.0])

    add_heading(doc, "Système de gestion de base de données", 3)

    add_body(doc,
        "Entre PostgreSQL et MySQL, l'écart de performance transactionnelle pure est marginal pour ce "
        "projet : les deux moteurs sont matures et tiendraient la charge des 58 EFPA sans difficulté. "
        "Ce qui tranche est ailleurs. pgvector fait de PostgreSQL la seule option où les embeddings du "
        "Bot APC et les données académiques structurées (candidatures, notes, diplômes) vivent dans la "
        "même instance, sous le même périmètre d'audit CNDP, plutôt que dans deux systèmes qu'il "
        "faudrait synchroniser et auditer séparément. C'est un point concret : un vector store dédié "
        "comme Pinecone ou Weaviate impose un job de synchronisation entre la base de vérité "
        "(PostgreSQL) et l'index vectoriel, avec une fenêtre où les deux peuvent diverger, exactement "
        "le genre de double source de vérité qui produit des incidents difficiles à diagnostiquer sur "
        "un système en production. MongoDB pose un problème plus fondamental que la performance : son "
        "modèle de cohérence éventuelle et l'absence de garanties ACID transversales contredisent "
        "directement l'Article 19.3.2 du CPS, qui exige l'intégrité transactionnelle pour des données "
        "comme une délibération de concours ou une note d'examen, où une écriture partiellement "
        "appliquée n'est pas une dégradation acceptable mais une erreur administrative. Le dernier "
        "argument, souvent sous-estimé, est la sécurité au niveau ligne : PostgreSQL l'implémente "
        "nativement, ce qui permet d'isoler les données d'un EFPA des 57 autres par une politique "
        "déclarative vérifiable, alors que MySQL ou MongoDB obligeraient à reconstruire cette "
        "isolation dans le code applicatif, où une erreur de développeur devient une fuite de données "
        "entre établissements plutôt qu'une erreur de configuration détectable en revue."
    )
    headers_b = ["Critère", "PostgreSQL + pgvector (retenu)", "MySQL / MariaDB", "MongoDB", "Vector DB dédiée (Pinecone, Weaviate)"]
    rows_b = [
        ["Conformité Art.19.3.2 CPS", "SGBDR imposé, intégrité ACID native", "Relationnel mais non nommé par le CPS", "Non relationnel, incompatible avec l'exigence ACID", "Pas un SGBDR"],
        ["Recherche vectorielle pour le Bot APC", "pgvector dans la même instance", "Extension tierce ou base séparée nécessaire", "Vector Search propriétaire, hors PostgreSQL", "Native, mais ajoute un composant et un point de synchronisation"],
        ["Row-Level Security multi-tenant (58->100 EFPA)", "Native", "Absente, à reconstruire en code applicatif", "Sécurité au niveau document, pas au niveau ligne", "Non applicable, pas de modèle relationnel"],
        ["Verdict", "Retenu", "Écarté", "Écarté", "Écarté : duplication d'infrastructure sans gain à cette échelle"],
    ]
    add_table(doc, headers_b, rows_b, caption="Tableau 4. Comparaison des systèmes de gestion de données envisagés",
              col_widths=[4.2, 3.2, 3.0, 2.8, 3.0])

    add_heading(doc, "Style d'architecture applicative", 3)

    add_body(doc,
        "Un monolithe modulaire bien conçu, avec une séparation stricte des modules en interne, "
        "aurait été plus simple à opérer pour une équipe de quatre personnes, et ce point n'est pas "
        "négligeable. Il a été écarté pour une raison de conformité littérale plutôt que de "
        "préférence technique : l'Article 19.3.1 du CPS exige une architecture n-tiers garantissant "
        "une séparation physique et logique entre les couches, une formulation qu'un comité technique "
        "peut raisonnablement lire comme excluant un processus unique, quelle que soit la qualité de "
        "sa modularité interne. Le risque n'est pas hypothétique : c'est exactement le type d'écart "
        "qui coûte des points sur le Critère A sans qu'aucune ligne de code ne soit en cause. Le "
        "serverless pose un problème différent, mesurable plutôt qu'interprétatif. Une fonction "
        "Node.js incluant le client Prisma et le runtime NestJS connaît un temps de démarrage à froid "
        "qui se chiffre en centaines de millisecondes, parfois plus d'une seconde, avant que la "
        "moindre logique applicative ne s'exécute. Ce délai seul peut consommer la totalité du budget "
        "de 1,5 seconde fixé par l'Article 19.4.2 pour la réponse du Bot APC, avant même que "
        "l'inférence du modèle ne commence. Le serverless reste pertinent pour des traitements "
        "sporadiques et tolérants à la latence, pas pour une interface conversationnelle face à des "
        "apprenants sous engagement de SLA. Le découpage en neuf services par cohésion métier, "
        "retenu et détaillé à la section 3.4, reste donc l'option qui satisfait à la fois la lettre du "
        "CPS et la réalité opérationnelle d'une équipe de quatre experts."
    )
    headers_c = ["Critère", "Microservices par cohésion métier (retenu)", "Monolithe modulaire", "Serverless (FaaS)"]
    rows_c = [
        ["Conformité Art.19.3.1 CPS (blocs indépendants, scalabilité horizontale)", "Conforme directement", "Conforme en interne, scalabilité non indépendante par bloc", "Scalabilité automatique, mais granularité excessive"],
        ["Opérabilité avec une équipe de 4 experts", "Neuf services, dimensionnés pour rester gérables", "Plus simple à opérer, mais ne répond pas à l'indépendance exigée", "Complexité de débogage distribué disproportionnée"],
        ["Isolation du pool GPU (Bot APC, Art.19.4.2)", "Native : service et nœuds dédiés", "Difficile à isoler dans un seul processus", "Cold start incompatible avec le seuil de 1,5 seconde"],
        ["Verdict", "Retenu", "Écarté : non conforme à l'Art.19.3.1", "Écarté : latence d'amorçage incompatible avec le SLA Bot APC"],
    ]
    add_table(doc, headers_c, rows_c, caption="Tableau 5. Comparaison des styles d'architecture applicative envisagés",
              col_widths=[5.5, 4.0, 4.0, 4.0])

    add_heading(doc, "Moteur d'inférence pour le Bot APC", 3)

    add_body(doc,
        "vLLM, Text Generation Inference et Ollama savent tous les trois faire tourner un modèle de "
        "langage sur une unité GPU et répondre correctement à une requête isolée : sur ce point, les "
        "trois passeraient une démonstration. La différence apparaît sous charge concurrente, "
        "exactement le scénario que le dimensionnement de l'Article 19.4.2 doit couvrir avec une "
        "unique unité GPU NVIDIA A10, sans possibilité contractuelle d'en ajouter une seconde si le "
        "débit se dégrade. vLLM gère le cache d'attention par blocs non contigus plutôt que par bloc "
        "réservé à l'avance pour chaque requête, ce qui lui permet de traiter en parallèle un nombre "
        "de requêtes nettement supérieur sans fragmentation mémoire, un avantage qui ne se voit pas "
        "sur une requête seule mais qui devient déterminant lorsque plusieurs EFPA interrogent le Bot "
        "APC au même moment, par exemple en période d'examens. TGI reste une alternative sérieuse, "
        "avec une justesse de réponse comparable, mais les mesures de débit publiées le placent "
        "derrière vLLM précisément sur la charge concurrente, l'unique scénario qui compte ici puisque "
        "la capacité GPU est fixée par contrat et ne peut pas être augmentée en cours d'exécution pour "
        "compenser un choix moins favorable. Ollama, enfin, est pensé pour un usage local et "
        "mono-utilisateur : l'adopter en production aurait reporté à plus tard une réécriture devenue "
        "nécessaire dès le premier pic de charge réel, un risque qu'il est moins coûteux d'éviter à la "
        "conception qu'à corriger après mise en service."
    )
    headers_d = ["Critère", "vLLM (retenu)", "Text Generation Inference (TGI)", "Ollama"]
    rows_d = [
        ["Débit sous charge concurrente (58 EFPA)", "Optimisé (traitement par lots continu)", "Bon, débit inférieur sous forte concurrence", "Conçu pour un usage local, pas pour la production multi-utilisateurs"],
        ["Compatibilité API OpenAI (appel depuis LangChain.js)", "Native", "Native", "Partielle"],
        ["Verdict", "Retenu", "Alternative valable, écartée sur le débit", "Écarté pour un usage de production"],
    ]
    add_table(doc, headers_d, rows_d, caption="Tableau 6. Comparaison des moteurs d'inférence GPU envisagés",
              col_widths=[5.5, 4.0, 4.5, 4.5])

    add_body(doc,
        "Le diagramme de déploiement présenté à la section 4.6 traduit directement le verdict de ces "
        "quatre comparaisons : c'est pourquoi chaque nœud y est nommé par la technologie retenue "
        "plutôt que par une étiquette générique de type \"service applicatif\" ou \"base de données\"."
    )

    add_heading(doc, "Périmètre fonctionnel - Cas d'utilisation", 2)

    add_body(doc,
        "Le diagramme suivant illustre, sur le flux Admissions et Concours, la répartition des "
        "responsabilités entre les quatre acteurs du processus. Le candidat ne dispose que de cas "
        "d'utilisation en lecture ou en dépôt ; la présélection, la convocation, la saisie des notes "
        "et la délibération restent exclusivement du ressort de l'agent EFPA et du comité, "
        "conformément à la distinction entre sélection automatique et manuelle posée par l'Article "
        "19.5.2 du CPS, qui ne confère jamais au candidat de droit d'écriture sur le processus de "
        "notation."
    )
    add_figure(doc, usecase_path, "Cas d'utilisation - flux Admissions et Concours")

    add_body(doc,
        "Ce diagramme est délibérément restreint au flux Admissions et Concours plutôt que de "
        "tenter une vue exhaustive des onze blocs fonctionnels sur un seul schéma : un diagramme "
        "couvrant l'intégralité du périmètre aurait perdu en lisibilité ce qu'il aurait gagné en "
        "exhaustivité apparente, et la couverture complète des blocs fonctionnels est déjà démontrée, "
        "bloc par bloc, par le tableau de correspondance de la section 3.2. Le choix de ne jamais "
        "donner au candidat de droit d'écriture sur la notation, alors qu'une dérogation technique "
        "serait possible, répond à une exigence d'auditabilité : toute valeur de note doit pouvoir "
        "être tracée à un agent identifié et habilité, jamais à l'auteur de la copie évaluée."
    )

    add_heading(doc, "Structure des données - Diagramme de classes", 2)

    add_body(doc,
        "Le modèle de données du flux Admissions et Concours illustre le principe retenu pour "
        "l'ensemble du domaine Parcours Apprenant : les entités opérationnelles (Note, Convocation) "
        "référencent la Candidature et non directement le Candidat, ce qui permet à une même "
        "personne de postuler à plusieurs sessions sans collision de données. La Délibération produit "
        "exclusivement une Liste d'Attente ou une Inscription Définitive pour une même candidature, "
        "une contrainte d'exclusivité portée par le schéma de base plutôt que laissée à la seule "
        "discipline applicative."
    )
    add_figure(doc, classes_path, "Diagramme de classes - domaine Admissions et Concours")

    add_body(doc,
        "Un schéma plus générique aurait pu modéliser Note et Convocation comme des attributs "
        "libres rattachés à la Candidature, par exemple sous forme de document JSON, ce qui aurait "
        "simplifié l'évolution du modèle si de nouveaux champs apparaissaient. Ce choix a été écarté "
        "au profit d'un schéma explicite, champ par champ, parce que la conformité CNDP au titre de "
        "la loi 09-08 exige de pouvoir désigner précisément quelles colonnes contiennent des données "
        "à caractère personnel : un schéma opaque rendrait cette désignation impossible à auditer. "
        "La contrainte d'exclusivité entre Liste d'Attente et Inscription Définitive est portée par "
        "le schéma plutôt que par le seul code applicatif, afin qu'une erreur de programmation future "
        "ne puisse pas produire un candidat à la fois admis et en attente, un état que le CPS ne "
        "prévoit pas et que la base de données interdit donc structurellement."
    )

    add_heading(doc, "Circuit de traitement - Diagramme de séquence", 2)

    add_body(doc,
        "Le circuit de délibération illustre le découplage entre l'écriture du résultat, "
        "synchrone et bloquante jusqu'à confirmation de l'audit, et la notification des candidats, "
        "mise en file et traitée de façon asynchrone. Aucun résultat n'est publié sans que sa trace "
        "d'audit soit garantie persistée, et la publication n'est jamais ralentie par la latence d'un "
        "envoi de notifications en masse."
    )
    add_figure(doc, sequence_path, "Diagramme de séquence - délibération et publication des résultats")

    add_body(doc,
        "Mettre l'écriture d'audit et l'envoi des notifications sur le même mécanisme de file "
        "asynchrone aurait été plus simple à implémenter, mais aurait traité deux incidents de "
        "nature différente avec la même tolérance : la perte d'une notification est rattrapable par "
        "un nouvel envoi, la perte d'une trace d'audit ne l'est pas, puisqu'elle prive la DEFR d'une "
        "preuve d'imputabilité exigée à l'Article 19.3.5 du CPS. C'est cette différence de "
        "réversibilité de l'erreur, et non une préférence de conception arbitraire, qui justifie de "
        "garder l'écriture d'audit sur le chemin synchrone tout en mettant la notification en file."
    )

    add_heading(doc, "Architecture de déploiement et haute disponibilité", 2)

    add_body(doc,
        "Le diagramme de déploiement ci-dessous traduit en infrastructure réelle les exigences de "
        "l'Article 19.4 du CPS. Le périmètre d'entrée combine pare-feu applicatif, protection "
        "anti-DDoS et répartiteur de charge Multi-AZ, conformément au Pack Sécurité Avancé de "
        "l'Article 19.4.4. Le cluster Kubernetes répartit les pods applicatifs sur deux zones de "
        "disponibilité avec bascule automatique sans perte de session, et isole le pool de nœuds GPU "
        "dédié à l'inférence du Bot APC du reste du cluster, afin qu'une charge applicative "
        "inhabituelle sur les autres services ne dégrade jamais le temps de réponse de "
        "1,5 seconde fixé à l'Article 19.4.2. La couche données réplique PostgreSQL de façon "
        "synchrone entre les deux zones et externalise des sauvegardes quotidiennes avec restauration "
        "à un instant précis (PITR), et l'Audit Trail Store reste un composant à part, sur un support "
        "non modifiable, pour qu'aucune procédure de restauration de données applicatives ne puisse "
        "altérer la piste d'audit des douze derniers mois."
    )
    add_figure(doc, deployment_path, "Diagramme de déploiement - infrastructure Multi-AZ sur Cloud souverain")

    add_body(doc,
        "L'isolation du pool GPU est le point le plus directement défendable de ce schéma : "
        "héberger l'inférence du Bot APC sur les mêmes nœuds que les services transactionnels aurait "
        "été moins coûteux à exploiter, mais aurait exposé le temps de réponse du Bot APC à toute "
        "saturation provoquée par un pic de saisie de notes ou de publication de résultats de "
        "concours. La réplication synchrone entre les deux zones de disponibilité, plutôt "
        "qu'asynchrone, est retenue malgré son coût en latence d'écriture, parce que le CPS exige une "
        "bascule sans perte de session : une réplication asynchrone pourrait faire basculer le trafic "
        "vers une zone dont la dernière transaction n'a pas encore été appliquée."
    )

    add_heading(doc, "Pipeline de déploiement continu et sécurisé", 2)

    add_body(doc,
        "Le déploiement de chaque module suit un pipeline qui combine les contrôles de sécurité "
        "logiciels exigés implicitement par la loi 05-20 et une stratégie de mise en production sans "
        "interruption de service. Après les tests unitaires, une analyse statique de sécurité (SAST) "
        "et un contrôle des vulnérabilités connues sur les dépendances s'exécutent avant toute "
        "construction d'image conteneurisée ; l'image est ensuite signée et déposée dans un registre "
        "sécurisé, ce qui empêche le déploiement d'une image modifiée après ce contrôle. Le passage "
        "en environnement de staging inclut des tests de charge ciblés sur le Bot APC et sur le "
        "module Certification, les deux points les plus sensibles aux exigences de performance et de "
        "sécurité du CPS, avant une validation manuelle du comité de suivi qui correspond à un jalon "
        "contractuel et non à une simple formalité interne."
    )
    add_figure(doc, cicd_path, "Pipeline CI/CD avec bascule Blue-Green et rollback automatique")

    add_body(doc,
        "La bascule en production suit une stratégie Blue-Green plutôt qu'un remplacement direct des "
        "instances en service : la nouvelle version est déployée intégralement en parallèle de "
        "l'ancienne, le trafic n'est basculé qu'une fois les vérifications de santé confirmées, et "
        "l'ancienne version reste disponible pour un retour immédiat en cas d'anomalie détectée par "
        "la supervision continue. Un déploiement progressif par petits incréments aurait réduit la "
        "consommation d'infrastructure pendant la transition, mais aurait rendu plus difficile à "
        "diagnostiquer une anomalie n'apparaissant qu'à charge complète, un risque jugé trop élevé "
        "au regard du seuil de disponibilité de 99,95% fixé à l'Article 19.4.3 du CPS."
    )


def section_equipe(doc, points_path):
    add_heading(doc, "Équipe proposée", 1, numbered="5")

    add_heading(doc, "Principe de constitution de l'équipe", 2)

    add_body(doc,
        "L'Article 10 du RC publie le barème exact de notation de l'équipe : 60 des 100 points de "
        "la note technique se décomposent sous-critère par sous-critère, pour chacun des quatre "
        "profils, avec un seuil précis pour chaque palier de points. L'équipe ci-dessous est "
        "présentée directement face à cette grille, pour que le comité d'évaluation puisse vérifier "
        "ligne par ligne le palier atteint par chaque profil proposé."
    )
    add_figure(doc, points_path, "Répartition des points potentiels par profil - Critère C (60 points, Art.10 RC)")

    add_heading(doc, "Profils détaillés des experts proposés", 2)

    add_body(doc,
        "Brahim EL ORF, Chef de Projet (Expert SI), est titulaire d'un diplôme d'Ingénieur d'État "
        "en Informatique de l'École Mohammedia d'Ingénieurs (2009), auditeur principal certifié ISO "
        "27001:2022 Lead Auditor, certifié PMP et ITIL v3. Il dirige des projets informatiques depuis "
        "2009 : chef de projet SI chez INWI Corporate sur la migration du système d'information vers "
        "la plateforme GSM, chef de projet chez MTN Afrique du Sud, ingénieur puis chef de projet "
        "chez NBTY (Royaume-Uni) sur la migration vers Amazon AWS déployée sur plus de 700 agences, "
        "consultant chez SkillNet Solutions sur des solutions Oracle Retail, et Directeur Technique "
        "chez ITINFODEV depuis 2020. Ses missions récentes pour l'Agence Urbaine de Rabat-Salé, la "
        "FAO-Maroc et l'OIT-Maroc démontrent sa maîtrise de la conduite de systèmes d'information "
        "complexes multi-acteurs, transposable au pilotage du SIS Intégré sur 58 établissements."
    )

    add_body(doc,
        "Mohammed OUAZZANE, Ingénieur Systèmes (Mise en œuvre), est titulaire d'un diplôme "
        "d'Ingénieur en Informatique et Réseaux de l'EMSI, option MIAGE (2011). Il met en œuvre des "
        "infrastructures Cloud et des environnements applicatifs sécurisés depuis 2011 : migration "
        "vers Amazon AWS chez NBTY, administration Windows/Linux et Active Directory chez Géomatic "
        "et GEOFIT, et pilotage de projets SIG/IT chez FERA Solutions, où il a également assuré la "
        "formation et le transfert de compétences des équipes utilisatrices. Il maîtrise la "
        "configuration d'environnements SaaS sécurisés et l'administration des plateformes "
        "applicatives à l'échelle requise par le déploiement sur 58 EFPA."
    )

    add_body(doc,
        "Mouaad MOHY-EDDINE, Expert en Intelligence Artificielle, est titulaire d'un Doctorat en "
        "Sécurité Informatique et Intelligence Artificielle de l'EST Essaouira - Université Cadi "
        "Ayyad (2024) et d'un Master en Big Data et Aide à la Décision de l'ENSA Khouribga (2020). "
        "Maître de conférences en cybersécurité à l'ENSAM Casablanca depuis juillet 2024, il a formé "
        "des enseignants et des profils BTS en intelligence artificielle, machine learning et deep "
        "learning au Samsung Innovation Campus dans huit régions du Maroc entre 2022 et 2024, et "
        "publié plus de dix articles scientifiques internationaux en IA et cybersécurité. Son "
        "expérience est directement transposable à l'orchestration RAG du Bot APC et à l'indexation "
        "pédagogique des référentiels APC, puisqu'elle combine maîtrise technique de l'IA et pratique "
        "réelle de la pédagogie."
    )

    add_body(doc,
        "Noureddine AMENZOU, Développeur (Développement spécifique), est titulaire d'un Master en "
        "informatique de l'ENSI Tanger (2009) et certifié ISO 27001 et Cyber Security Foundation. "
        "Il développe des applications full-stack depuis plus de quinze ans, notamment chez APM "
        "Terminals (administration et développement applicatif, prix Star Award 2020) et pour "
        "plusieurs clients d'ABI Consulting en gestion de bases de données et développement web. Il "
        "maîtrise l'environnement Node.js, l'ORM Prisma et la création d'API sécurisées, "
        "directement mobilisables pour le développement des modules métier du SIS Intégré."
    )

    add_heading(doc, "Sous-critères de notation et niveau visé (Art.10 RC)", 2)

    add_body(doc,
        "Le tableau ci-dessous reprend littéralement la structure de notation du RC, avec le niveau "
        "atteint par chaque expert nommément proposé. Les curriculum vitae datés et signés et les "
        "diplômes en copies certifiées conformes sont joints en annexe, conformément à l'Article 9.2 "
        "du RC."
    )

    headers = ["Profil", "Sous-critère (Art.10 RC)", "Niveau atteint", "Pts"]
    rows = [
        ["Brahim EL ORF\n(Chef de Projet)", "Niveau de formation",
         "Ingénieur d'État, École Mohammedia d'Ingénieurs (2009)", "10"],
        ["Brahim EL ORF\n(Chef de Projet)", "Expérience en direction de projets informatiques",
         "Plus de 15 ans (depuis 2009)", "5"],
        ["Brahim EL ORF\n(Chef de Projet)", "Expérience en direction de projets informatiques éducatifs",
         "Conforme", "5"],
        ["Mohammed OUAZZANE\n(Ingénieur Système)", "Niveau de formation",
         "Ingénieur Informatique et Réseaux, EMSI (2011)", "10"],
        ["Mohammed OUAZZANE\n(Ingénieur Système)", "Expérience en système d'information éducatif",
         "Plus de 10 ans (depuis 2011)", "10"],
        ["Mouaad MOHY-EDDINE\n(Expert IA)", "Niveau de formation en Intelligence Artificielle",
         "Doctorat IA et Sécurité Informatique, EST Essaouira (2024)", "5"],
        ["Mouaad MOHY-EDDINE\n(Expert IA)", "Expérience en IA sur projet éducatif",
         "Plus de 2 ans (formateur IA depuis 2020)", "5"],
        ["Noureddine AMENZOU\n(Développeur)", "Niveau de formation en informatique",
         "Master informatique, ENSI Tanger (2009)", "5"],
        ["Noureddine AMENZOU\n(Développeur)", "Expérience confirmée en Node.js",
         "Plus de 5 ans", "5"],
    ]
    add_table(doc, headers, rows, caption="Tableau 7. Grille de notation de l'équipe reprise de l'Article 10 du RC",
              col_widths=[3.3, 6.0, 5.0, 1.2])

    add_heading(doc, "Couverture des services par profil", 2)

    add_body(doc,
        "Le tableau suivant met en correspondance chaque expert avec les neuf services de "
        "l'architecture présentée à la section 4, afin que la composition de l'équipe ne soit pas "
        "dissociée du découpage applicatif qu'elle doit livrer. Brahim EL ORF supervise l'ensemble "
        "sans porter de responsabilité de développement direct sur un service en particulier."
    )

    headers2 = ["Service", "EL ORF", "OUAZZANE", "MOHY-EDDINE", "AMENZOU"]
    rows2 = [
        ["Gateway", "Supervision", "Responsable", "—", "Contributeur"],
        ["Référentiel", "Supervision", "Responsable", "—", "Contributeur"],
        ["Parcours Apprenant", "Supervision", "Contributeur", "—", "Responsable"],
        ["Planning & Ressources", "Supervision", "Responsable", "—", "Contributeur"],
        ["Stages & Insertion", "Supervision", "—", "—", "Responsable"],
        ["Certification", "Supervision", "Responsable", "—", "Contributeur"],
        ["BI & Reporting", "Supervision", "Contributeur", "Contributeur", "—"],
        ["Bot APC / IA", "Supervision", "Contributeur", "Responsable", "Contributeur"],
        ["Interopérabilité & Migration", "Supervision", "Responsable", "—", "Contributeur"],
    ]
    add_table(doc, headers2, rows2, caption="Tableau 8. Matrice de responsabilité par service",
              col_widths=[4.5, 2.7, 3.0, 2.5, 2.5])


def section_planning(doc, gantt_path):
    add_heading(doc, "Planning d'exécution", 1, numbered="6")

    add_body(doc,
        "Le délai global d'exécution est fixé à sept mois par l'Article 20 du CPS, réparti en "
        "quatre phases contractuelles dont l'enchaînement est représenté ci-dessous. Chaque jalon "
        "correspond à un livrable nommé de l'Article 21 du CPS, et conditionne le règlement de la "
        "phase concernée conformément à l'Article 24."
    )
    add_figure(doc, gantt_path, "Planning d'exécution sur 7 mois")

    add_body(doc,
        "La Phase II concentre le développement des modules à plus forte valeur différenciante "
        "(Certification, Headless BI, Bot APC), démarrée en chevauchement avec la fin de la Phase I "
        "pour que le socle applicatif (paramétrage EFPA, filières, APC) soit disponible avant que les "
        "modules qui en dépendent ne commencent leur développement. La migration technique des "
        "données historiques est positionnée en début de Phase III, avant l'interconnexion Moodle et "
        "le déploiement sur les 58 EFPA, afin que les anomalies de migration soient résolues avant "
        "que le système ne soit exposé aux utilisateurs réels."
    )


def section_chronogramme(doc):
    add_heading(doc, "Chronogramme d'affectation du personnel", 1, numbered="7")

    add_body(doc,
        "Le chronogramme ci-dessous présente la charge de travail prévisionnelle en jours-hommes par "
        "expert et par phase, conformément à la pièce C de l'Article 9.2 du RC."
    )

    headers = ["Expert (Profil)", "Phase I", "Phase II", "Phase III", "Phase IV", "Total"]
    rows = [
        ["Brahim EL ORF (Chef de Projet)", "20", "45", "25", "10", "100"],
        ["Mohammed OUAZZANE (Ingénieur Systèmes)", "15", "20", "30", "5", "70"],
        ["Mouaad MOHY-EDDINE (Expert IA)", "5", "30", "10", "5", "50"],
        ["Noureddine AMENZOU (Développeur)", "5", "55", "20", "5", "85"],
    ]
    add_table(doc, headers, rows, caption="Tableau 9. Chronogramme d'affectation en jours-hommes",
              col_widths=[5.0, 2.3, 2.3, 2.3, 2.3, 1.8])


def section_reversibilite(doc):
    add_heading(doc, "Plan de réversibilité et transfert de compétences", 1, numbered="8")

    add_body(doc,
        "Conformément à l'Article 9.4 du RC, ce plan décrit les modalités techniques et "
        "organisationnelles garantissant le transfert total du savoir-faire, de la documentation et "
        "du code source à la DEFR en fin de projet, en cohérence avec l'Article 29 du CPS qui réserve "
        "à la DEFR la propriété exclusive des composants spécifiques métier développés."
    )

    add_body(doc,
        "Le transfert technique repose sur trois livrables distincts. Le code source des neuf "
        "services, organisé en monorepo NestJS/TypeScript, est remis commenté et accompagné des "
        "scripts de migration Prisma, dans un format directement compilable sans dépendance à un "
        "outillage propriétaire du titulaire. La documentation d'architecture, incluant les schémas "
        "de ce dossier et leur mise à jour à l'issue du développement réel, est remise au format "
        "source modifiable. Les enregistrements vidéo des sessions de transfert de compétences, "
        "livrables de la Phase IV, couvrent l'administration fonctionnelle pour les profils "
        "administrateurs et l'utilisation pédagogique pour les profils formateurs."
    )

    add_body(doc,
        "Le transfert organisationnel se déroule en parallèle de la Phase IV, par des sessions "
        "conjointes entre l'équipe projet et les administrateurs désignés par la DEFR, structurées "
        "par module fonctionnel plutôt qu'en une session unique généraliste. Cette structuration "
        "permet à un administrateur EFPA de ne suivre que les modules pertinents pour son périmètre, "
        "et facilite la création d'un programme de formation continue interne à la DEFR au-delà de la "
        "période de garantie."
    )


def section_valeur_ajoutee(doc, radar_path):
    add_heading(doc, "Valeur ajoutée et points distinctifs", 1, numbered="9")

    add_body(doc,
        "Le radar ci-dessous compare, sur six dimensions structurantes du marché, le seuil minimal "
        "exigé par le CPS et le RC au niveau effectivement proposé par ABI CONSULTING. Chaque écart "
        "entre les deux courbes correspond à un choix technique précis, documenté dans les sections "
        "qui précèdent, et non à une appréciation qualitative non étayée."
    )
    add_figure(doc, radar_path, "Niveau proposé par rapport au seuil minimal du CPS/RC")

    add_body(doc,
        "Trois choix techniques expliquent l'essentiel de cet écart."
    )

    add_bullet_dot(doc, [
        "Séparation OLTP/OLAP pour le Headless BI : les requêtes analytiques ne dégradent jamais le "
        "service rendu aux apprenants en période de pointe, notamment lors de la publication des "
        "résultats de concours sur les 58 établissements simultanément, ce qui explique l'écart sur "
        "les dimensions Performance et Pilotage.",
        "Architecture en schéma partagé avec sécurité au niveau ligne, conçue dès l'origine pour "
        "l'extensibilité à 100 établissements annoncée par le CPS, sans migration de modèle de "
        "données lorsque ce seuil sera atteint, ce qui explique l'écart sur la dimension Évolutivité.",
        "Mobilisation des SDK officiels Microsoft (MSAL, Graph) pour l'intégration Entra ID et M365, "
        "qui exploite l'écosystème déjà déployé à la DEFR avec un risque de configuration réduit, ce "
        "qui explique l'écart sur la dimension Interopérabilité.",
    ])


def section_gestion_risques(doc):
    add_heading(doc, "Gestion des risques", 1, numbered="10")

    add_body(doc,
        "Les risques retenus sont spécifiques à ce marché et à son contexte d'exécution, identifiés "
        "à partir des contraintes propres au CPS et au RC plutôt que d'une liste générique de "
        "risques de projet informatique."
    )

    headers = ["Risque", "Mesure de mitigation"]
    rows = [
        ["R1 - Perte d'intégrité lors de la migration HELISA (5 ans d'historique)",
         "Migration par lots avec somme de contrôle et comptage de lignes ; suspension du seul lot "
         "en anomalie sans affecter les lots déjà validés"],
        ["R2 - Latence du Bot APC supérieure à 1,5 seconde sous charge",
         "Dimensionnement GPU dédié conforme à l'Article 19.4.2 ; mise en cache adaptative des "
         "réponses fréquentes par profil"],
        ["R3 - Absence du Partner ID Microsoft à la soumission",
         "Vérification anticipée de la détention de l'attestation avant dépôt ; mobilisation des "
         "SDK officiels Microsoft qui ne requièrent pas ce partenariat pour l'intégration technique "
         "elle-même"],
        ["R4 - Hétérogénéité des données sources entre les 58 EFPA",
         "Dictionnaire de données et règles de transcodage par établissement, documentés dans le "
         "rapport de stratégie de migration ETL avant extraction"],
        ["R5 - Délai de 7 mois serré au regard des 11 blocs fonctionnels",
         "Chevauchement maîtrisé des phases (planning Section 6) et équipe à plein temps sur la "
         "Phase II ; pas de marge supplémentaire négociée dans le délai contractuel"],
        ["R6 - Équipe minimale de 4 personnes sans remplaçant immédiat",
         "Documentation continue (et non différée en fin de phase) pour limiter la dépendance à un "
         "seul expert sur chaque module ; clause de remplacement prévue à l'Article 17 du CPS"],
    ]
    add_table(doc, headers, rows, caption="Tableau 10. Registre des risques et mesures de mitigation",
              col_widths=[7.0, 8.0])


# ---------------------------------------------------------------
# Assemblage du document
# ---------------------------------------------------------------

def build_document():
    print("Génération des figures...")
    arch_path = generate_architecture_diagram()
    print(f"  Architecture: {arch_path}")
    growth_path = generate_efpa_growth_chart()
    print(f"  Croissance EFPA: {growth_path}")
    objectifs_path = generate_objectifs_mission_diagram()
    print(f"  Objectifs/bénéfices: {objectifs_path}")
    points_path = generate_team_points_chart()
    print(f"  Points équipe: {points_path}")
    usecase_path = generate_usecase_diagram()
    print(f"  Cas d'utilisation: {usecase_path}")
    classes_path = generate_class_diagram()
    print(f"  Classes: {classes_path}")
    sequence_path = generate_sequence_diagram()
    print(f"  Séquence: {sequence_path}")
    deployment_path = generate_deployment_diagram()
    print(f"  Déploiement: {deployment_path}")
    cicd_path = generate_cicd_pipeline()
    print(f"  Pipeline CI/CD: {cicd_path}")
    gantt_path = generate_gantt()
    print(f"  Gantt: {gantt_path}")
    radar_path = generate_value_radar_chart()
    print(f"  Radar de valeur ajoutée: {radar_path}")

    print("Construction du document...")
    doc = Document()
    set_margins(doc)
    set_document_font_theme(doc)
    add_page_numbers(doc)

    page_de_garde(doc)
    section_sommaire(doc)
    section_introduction(doc)
    section_comprehension_contexte(doc, growth_path)
    doc.add_page_break()
    section_comprehension_mission(doc, objectifs_path)
    doc.add_page_break()
    section_approche_methodologique(doc)
    doc.add_page_break()
    section_architecture(doc, arch_path, usecase_path, classes_path, sequence_path,
                         deployment_path, cicd_path)
    doc.add_page_break()
    section_equipe(doc, points_path)
    doc.add_page_break()
    section_planning(doc, gantt_path)
    doc.add_page_break()
    section_chronogramme(doc)
    doc.add_page_break()
    section_reversibilite(doc)
    doc.add_page_break()
    section_valeur_ajoutee(doc, radar_path)
    doc.add_page_break()
    section_gestion_risques(doc)

    return doc


def inject_template_cover(output_path, template_path):
    """Injecte la page de garde du template, adaptée au texte DEFR."""
    import uuid

    with zipfile.ZipFile(template_path, 'r') as zt:
        template_doc_xml = zt.read('word/document.xml').decode('utf-8')
        template_rels_xml = zt.read('word/_rels/document.xml.rels').decode('utf-8')
        cover_image_bytes = zt.read('word/media/image1.png')

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

    # --- Adaptation du texte du template (DSS -> DEFR) ---
    old_direction = "Direction de la Stratégie et des Statistiques (DSS)"
    new_direction = "Direction de l'Enseignement, de la Formation et de la Recherche (DEFR)"

    old_ref = "N°01/2026/DSS"
    new_ref = "N°01/2026/DEFR/DETFP"

    old_objet = (
        "Conception, développement et déploiement d’une plateforme intégrée de suivi de la "
        "Stratégie Génération Green au profit de la Direction de la Stratégie et des Statistiques "
        "(DSS) du Ministère de l’Agriculture, de la Pêche Maritime, du Développement Rural et "
        "des Eaux et Forêts-Département de l’Agriculture à Rabat"
    )
    new_objet = (
        "Conception, développement, mise en œuvre de bases de données et système "
        "d’information apprenant (SIS Intégré) destiné aux établissements de formation "
        "professionnelle agricole (EFPA) relevant de la Direction de l’Enseignement, de la "
        "Formation et de la Recherche (DEFR) et sa mise en service en mode SaaS"
    )

    cover_xml = cover_xml.replace(old_objet, new_objet)
    cover_xml = cover_xml.replace(old_direction, new_direction)
    cover_xml = cover_xml.replace(old_ref, new_ref)

    new_rid = "rIdCover1"
    cover_xml = cover_xml.replace('r:embed="rId8"', f'r:embed="{new_rid}"')

    tmp_path = output_path + ".tmp.docx"
    with zipfile.ZipFile(output_path, 'r') as zin, zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        for name in names:
            data = zin.read(name)

            if name == 'word/document.xml':
                doc_xml = data.decode('utf-8')
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

        zout.writestr('word/media/cover_image1.png', cover_image_bytes)

    shutil.move(tmp_path, output_path)
    print("Page de garde template injectée (adaptée DEFR).")


if __name__ == "__main__":
    doc = build_document()
    output_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "offre_technique_sis_defr_20260619.docx"
    )
    doc.save(output_path)
    template_path = (
        "/home/almukhadram/Desktop/AO/AOI N¯  01-2026-DSS VF/note methdologique DSS.docx"
    )
    if os.path.exists(template_path):
        inject_template_cover(output_path, template_path)
    print(f"\nDocument sauvegardé: {output_path}")
