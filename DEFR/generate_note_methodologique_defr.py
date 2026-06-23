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
        (1, "1",   "Compréhension du contexte",                                4),
        (1, "2",   "Compréhension de la mission",                              5),
        (2, "2.1", "Objectifs et livrables",                                   5),
        (2, "2.2", "Périmètre et limites de la mission",                       6),
        (1, "3",   "Approche méthodologique",                                  7),
        (2, "3.1", "Démarche de conduite de projet",                          7),
        (2, "3.2", "Stratégie ETL et migration des données historiques",      9),
        (2, "3.3", "Architecture SaaS et Cloud",                              11),
        (2, "3.4", "Interconnexion et SSO",                                   13),
        (2, "3.5", "Approche IA - Bot APC",                                   14),
        (2, "3.6", "Exploitation, sécurisation et performance des données",   16),
        (1, "4",   "Architecture technique détaillée",                        17),
        (2, "4.1", "Architecture globale de la plateforme",                  17),
        (2, "4.2", "Périmètre fonctionnel - Cas d'utilisation",               18),
        (2, "4.3", "Structure des données - Diagramme de classes",            19),
        (2, "4.4", "Circuit de traitement - Diagramme de séquence",           20),
        (1, "5",   "Équipe proposée",                                         21),
        (2, "5.1", "Organisation de la mission",                              21),
        (2, "5.2", "Profils exigés et conformité au CPS",                     22),
        (2, "5.3", "Tableau synthétique de l'équipe",                         23),
        (1, "6",   "Planning d'exécution",                                    24),
        (1, "7",   "Chronogramme d'affectation du personnel",                 25),
        (1, "8",   "Plan de réversibilité et transfert de compétences",       26),
        (1, "9",   "Valeur ajoutée et points distinctifs",                    27),
        (1, "10",  "Gestion des risques",                                     28),
    ]

    TAB_POS = "8789"

    for level, num, title, page in entries:
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(3 if level == 1 else 1)
        para.paragraph_format.left_indent = Cm(0) if level == 1 else Cm(0.7)

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
            run_num.font.color.rgb = RGBColor(0x0F, 0x47, 0x61)

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


def generate_organigramme():
    dot = Digraph('organigramme_defr')
    dot.attr(rankdir='TB', nodesep='0.6', ranksep='0.55', dpi='200',
             fontname='Helvetica', bgcolor='white')

    common = {'shape': 'box', 'style': 'filled', 'fontname': 'Helvetica',
              'fontsize': '9', 'margin': '0.2,0.12'}

    dot.node('cdp', 'Chef de Projet\n(Expert SI)',
             fillcolor='#15607F', fontcolor='white', width='2.4', **common)

    with dot.subgraph() as s:
        s.attr(rank='same')
        s.node('se', 'Ingénieur Systèmes\n(Mise en œuvre)',
               fillcolor='#0F4761', fontcolor='white', width='2.2', **common)
        s.node('ia', 'Expert en\nIntelligence Artificielle',
               fillcolor='#0F4761', fontcolor='white', width='2.2', **common)

    dot.node('dev', 'Développeur\n(Développement spécifique)',
             fillcolor='#1A7A9C', fontcolor='white', width='2.4', **common)

    dot.edge('cdp', 'se')
    dot.edge('cdp', 'ia')
    dot.edge('se', 'dev')

    path_base = os.path.join(tempfile.gettempdir(), 'organigramme_defr')
    dot.render(path_base, format='png', cleanup=True)
    return path_base + '.png'


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
    fig, ax = plt.subplots(figsize=(18 / 2.54, 13 / 2.54))

    ax.axvspan(0, 1.5, alpha=0.07, color='#15607F', zorder=0)
    ax.axvspan(1.5, 4.5, alpha=0.07, color='#0F4761', zorder=0)
    ax.axvspan(4.5, 6.3, alpha=0.07, color='#1A7A9C', zorder=0)
    ax.axvspan(6.3, 7.0, alpha=0.07, color='#2E7D32', zorder=0)

    for i, (label, start, dur, color) in enumerate(activities):
        y = n - i
        ax.barh(y, dur, left=start, height=0.58, color=color, alpha=0.88,
                edgecolor='white', linewidth=0.5, zorder=2)

    for ms_x, ms_label, ms_y in milestones:
        ax.axvline(x=ms_x, color='#C17F00', linewidth=1.4, linestyle='--',
                   alpha=0.85, zorder=3)
        ax.plot(ms_x, n + ms_y, marker='D', color='#C17F00', markersize=5,
                zorder=5, clip_on=False)
        ax.text(ms_x + 0.12, n + ms_y, ms_label, fontsize=7.2,
                color='#C17F00', fontweight='bold', va='center', ha='left',
                zorder=6, clip_on=False)

    ax.text(0.75, n + 3.3, 'PHASE I\nDiagnostic et\nconception', fontsize=7.5, ha='center',
            color='#15607F', fontweight='bold', va='center')
    ax.text(3.0, n + 3.3, 'PHASE II\nDéveloppement et\nmise en œuvre métier', fontsize=7.5, ha='center',
            color='#0F4761', fontweight='bold', va='center')
    ax.text(5.4, n + 3.3, 'PHASE III\nIntégration et\ndéploiement', fontsize=7.5, ha='center',
            color='#1A7A9C', fontweight='bold', va='center')
    ax.text(6.65, n + 3.3, 'PHASE IV\nAccompa-\ngnement', fontsize=7.5, ha='center',
            color='#2E7D32', fontweight='bold', va='center')

    ax.set_yticks(range(1, n + 1))
    ax.set_yticklabels([a[0] for a in reversed(activities)], fontsize=7.3)
    ax.set_xlim(0, 7.3)
    ax.set_xticks(range(0, 8))
    ax.set_xticklabels([f"M{m}" for m in range(0, 8)], fontsize=8)
    ax.set_xlabel("Mois depuis l'ordre de service", fontsize=8.5)
    ax.set_title("Planning d'exécution - 7 mois, 4 phases (Art.20 CPS)",
                 fontsize=10, fontweight='bold')
    ax.spines[['top', 'right', 'left']].set_visible(False)
    ax.tick_params(left=False)
    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'gantt_defr.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    return path


def generate_risk_matrix():
    risks = [
        ("R1", "Perte d'intégrité lors de la migration\nHELISA (5 ans d'historique)", 2, 3),
        ("R2", "Latence du Bot APC > 1,5s sous charge", 2, 2),
        ("R3", "Absence du Partner ID Microsoft\nà la soumission", 1, 3),
        ("R4", "Hétérogénéité des données sources\nentre les 58 EFPA", 3, 2),
        ("R5", "Dépassement du SLA 99,95%\nen période de pointe", 2, 3),
        ("R6", "Désynchronisation Moodle\n(doublons, groupes)", 2, 2),
    ]

    fig, ax = plt.subplots(figsize=(14 / 2.54, 11 / 2.54))
    ax.set_xlim(0.5, 3.5)
    ax.set_ylim(0.5, 3.5)

    for x in [1, 2, 3]:
        for y in [1, 2, 3]:
            score = x * y
            color = '#C8E6C9' if score <= 2 else ('#FFE082' if score <= 6 else '#FFAB91')
            ax.add_patch(plt.Rectangle((x - 0.5, y - 0.5), 1, 1, facecolor=color,
                                        edgecolor='white', zorder=1))

    for rid, label, prob, impact in risks:
        ax.plot(prob, impact, marker='o', markersize=11, color='#15607F', zorder=3)
        ax.text(prob, impact, rid, color='white', fontsize=7.5, fontweight='bold',
                ha='center', va='center', zorder=4)
        ax.annotate(label, (prob, impact), textcoords="offset points",
                    xytext=(12, 8), fontsize=6.8, color='#1A1A1A', zorder=5)

    ax.set_xticks([1, 2, 3])
    ax.set_xticklabels(['Faible', 'Moyenne', 'Élevée'], fontsize=8.5)
    ax.set_yticks([1, 2, 3])
    ax.set_yticklabels(['Faible', 'Moyen', 'Élevé'], fontsize=8.5)
    ax.set_xlabel('Probabilité', fontsize=9, fontweight='bold')
    ax.set_ylabel('Impact', fontsize=9, fontweight='bold')
    ax.set_title('Matrice de criticité des risques', fontsize=10, fontweight='bold')
    plt.tight_layout()
    path = os.path.join(tempfile.gettempdir(), 'risk_matrix_defr.png')
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


def section_comprehension_contexte(doc):
    add_heading(doc, "Compréhension du contexte", 1, numbered="1")

    add_body(doc,
        "Le Département de l'Agriculture a opté, dans le cadre de la stratégie Génération Green "
        "2020-2030, pour la modernisation et l'informatisation de son système de formation "
        "professionnelle agricole. Cette orientation concerne directement les 58 Établissements de "
        "Formation Professionnelle Agricole (EFPA) relevant de la Direction de l'Enseignement, de "
        "la Formation et de la Recherche (DEFR), avec une extensibilité prévue jusqu'à 100 "
        "établissements. Le SIS Intégré que ce marché doit produire constitue l'outil par lequel "
        "cette modernisation devient opérationnelle, et non un projet informatique isolé du reste "
        "de la stratégie sectorielle."
    )

    add_body(doc,
        "Le dispositif actuel repose sur l'ERP HELISA et sur des pratiques de gestion hétérogènes "
        "selon les établissements. Les données de scolarité des cinq dernières années existent, "
        "mais sous des formats et des niveaux de qualité variables d'un EFPA à l'autre, ce qui rend "
        "leur exploitation centralisée difficile en l'état. La DEFR a par ailleurs déjà investi dans "
        "deux écosystèmes numériques actifs, le LMS Moodle pour la pédagogie et Microsoft 365 "
        "Education pour la collaboration et l'identité, avec lesquels le SIS Intégré doit "
        "s'interfacer dès sa mise en service plutôt que dans une phase ultérieure non financée."
    )

    add_body(doc,
        "Trois enjeux structurent ce contexte. Le premier est la fraude documentaire sur les "
        "diplômes, que la certification numérique avec signature électronique qualifiée doit "
        "neutraliser. Le second est l'absence de vision consolidée des indicateurs de réussite, "
        "d'attraction, d'absentéisme et d'insertion à l'échelle des 58 établissements, qui limite "
        "aujourd'hui le pilotage national par la DEFR à des remontées manuelles et asynchrones. Le "
        "troisième est la fragilité de l'historique de scolarité, qui doit être repris dans son "
        "intégralité avant que les nouveaux usages numériques ne s'y superposent."
    )

    add_body(doc,
        "Ce contexte s'inscrit également dans un cadre réglementaire contraignant que la solution "
        "doit respecter dès sa conception. La loi 05-20 relative à la cybersécurité impose un niveau "
        "de protection renforcé pour un système hébergeant les données de plus de 30 000 stagiaires "
        "actifs. La loi 09-08 relative à la protection des données à caractère personnel impose une "
        "conformité CNDP démontrable, et la loi 43-20 relative aux services de confiance encadre "
        "directement la certification numérique des diplômes. La souveraineté numérique, enfin, "
        "impose un hébergement exclusivement sur un Cloud qualifié situé sur le territoire national."
    )


def section_comprehension_mission(doc):
    add_heading(doc, "Compréhension de la mission", 1, numbered="2")

    add_heading(doc, "Objectifs et livrables", 2)

    add_body(doc,
        "L'objectif général du marché, tel que formulé à l'Article 3 du CPS, est de doter les 58 "
        "Établissements de Formation Professionnelle Agricole d'un système d'information moderne, "
        "performant, évolutif et conforme aux standards internationaux, permettant d'améliorer la "
        "qualité de la formation, la gestion des établissements et le pilotage global du dispositif. "
        "Cet objectif se décompose en exigences précises : collecte centralisée des données des 58 "
        "EFPA, suivi en temps opportun des programmes et activités, reprise intégrale de l'historique "
        "de scolarité sur cinq ans, et certification numérique des diplômes pour lutter contre la "
        "fraude documentaire."
    )

    add_body(doc,
        "Le SIS Intégré devra également couvrir la gestion harmonisée des admissions et concours, la "
        "gestion complète de la scolarité et des parcours pédagogiques, la gestion des évaluations, "
        "examens et délibérations selon les normes pédagogiques en vigueur, la planification et le "
        "suivi des activités de formation, la gestion des stages et de l'insertion des lauréats, et la "
        "production automatisée des documents officiels. L'intégration avec Moodle et Microsoft 365 "
        "Education conditionne directement l'adoption de la plateforme par les enseignants, qui "
        "utilisent déjà ces deux outils au quotidien."
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
        "qualifié. Il exclut explicitement la conception, le développement et le paramétrage métier "
        "du SIS de la sous-traitance, ce corps d'état principal devant rester sous la responsabilité "
        "directe de l'équipe projet."
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


def section_architecture(doc, arch_path, usecase_path, classes_path, sequence_path):
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

    add_heading(doc, "Circuit de traitement - Diagramme de séquence", 2)

    add_body(doc,
        "Le circuit de délibération illustre le découplage entre l'écriture du résultat, "
        "synchrone et bloquante jusqu'à confirmation de l'audit, et la notification des candidats, "
        "mise en file et traitée de façon asynchrone. Aucun résultat n'est publié sans que sa trace "
        "d'audit soit garantie persistée, et la publication n'est jamais ralentie par la latence d'un "
        "envoi de notifications en masse."
    )
    add_figure(doc, sequence_path, "Diagramme de séquence - délibération et publication des résultats")


def section_equipe(doc, org_path):
    add_heading(doc, "Équipe proposée", 1, numbered="5")

    add_heading(doc, "Organisation de la mission", 2)

    add_body(doc,
        "L'équipe mobilisée correspond strictement aux quatre profils minimaux exigés par l'Article "
        "22 du CPS et détaillés à l'Article 9.2 du RC. Le Chef de Projet supervise l'ensemble des "
        "quatre phases et constitue l'interlocuteur unique du comité de suivi de la DEFR. L'Ingénieur "
        "Systèmes et l'Expert en Intelligence Artificielle lui rapportent directement, et le "
        "Développeur intervient sous la supervision technique de l'Ingénieur Systèmes pour le "
        "développement spécifique des modules métier."
    )
    add_figure(doc, org_path, "Organigramme fonctionnel de l'équipe projet")

    add_heading(doc, "Profils exigés et conformité au CPS", 2)

    add_body(doc,
        "Les quatre profils ci-dessous reprennent les qualifications minimales fixées par l'Article "
        "22 du CPS. Les experts nommément proposés, leurs curriculum vitae datés et signés, leurs "
        "diplômes en copies certifiées conformes, ainsi que le chronogramme d'affectation détaillé "
        "par intervenant et par phase, seront intégrés à cette section dès leur identification "
        "définitive, conformément à l'Article 9.2 du RC. Cette intégration intervient avant le dépôt "
        "de l'offre : le RC élimine d'office toute offre technique où l'un de ces éléments serait "
        "absent ou non conforme."
    )

    headers = ["Profil exigé (Art.22 CPS)", "Qualifications minimales", "Expert proposé"]
    rows = [
        ["Chef de Projet\n(Expert SI)",
         "Bac+5 informatique ; plus de 15 ans en direction de projets informatiques ; au moins 3 "
         "projets informatiques éducatifs ; expertise ERP (Microsoft, SAP ou similaire) ; maîtrise "
         "impérative du déploiement et de l'administration de Moodle",
         "À confirmer"],
        ["Ingénieur Systèmes\n(Mise en œuvre)",
         "Bac+5 informatique ou réseaux ; au moins 10 ans dans la mise en œuvre d'infrastructures "
         "Cloud et la configuration d'environnements SaaS sécurisés de type système d'information "
         "éducatif",
         "À confirmer"],
        ["Expert en Intelligence\nArtificielle",
         "Bac+5 Intelligence Artificielle ; au moins 2 ans dans l'implémentation de services "
         "d'orchestration et l'intégration d'interfaces conversationnelles sur des projets de "
         "système d'information éducatif",
         "À confirmer"],
        ["Développeur\n(Développement spécifique)",
         "Bac+3 ou plus ; au moins 5 ans en développement full-stack ; maîtrise de Node.js, de l'ORM "
         "Prisma et de la création d'API sécurisées",
         "À confirmer"],
    ]
    add_table(doc, headers, rows, caption="Tableau 2. Conformité de l'équipe aux exigences du CPS",
              col_widths=[3.5, 9.0, 2.5])

    add_heading(doc, "Tableau synthétique de l'équipe", 2)

    headers2 = ["Expert", "Rôle dans la mission", "Phases d'intervention", "Pièces justificatives"]
    rows2 = [
        ["À confirmer", "Chef de Projet - pilotage global, interface comité de suivi DEFR",
         "Phases I à IV", "CV daté et signé, diplôme certifié conforme"],
        ["À confirmer", "Ingénieur Systèmes - infrastructure Cloud, déploiement SaaS",
         "Phases I, III, IV", "CV daté et signé, diplôme certifié conforme"],
        ["À confirmer", "Expert IA - orchestration RAG, indexation pgvector",
         "Phases II, III", "CV daté et signé, diplôme certifié conforme"],
        ["À confirmer", "Développeur - développement des modules métier",
         "Phases II, III", "CV daté et signé, diplôme certifié conforme"],
    ]
    add_table(doc, headers2, rows2, caption="Tableau 3. Synthèse de l'équipe et de ses pièces justificatives",
              col_widths=[2.5, 6.5, 3.0, 3.0])


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
        "profil et par phase. Il sera complété par les jours-hommes nominatifs de chaque expert dès "
        "réception des CV et de leur disponibilité confirmée, conformément à la pièce C de l'Article "
        "9.2 du RC."
    )

    headers = ["Profil", "Phase I", "Phase II", "Phase III", "Phase IV", "Total"]
    rows = [
        ["Chef de Projet", "20", "45", "25", "10", "100"],
        ["Ingénieur Systèmes", "15", "20", "30", "5", "70"],
        ["Expert IA", "5", "30", "10", "5", "50"],
        ["Développeur", "5", "55", "20", "5", "85"],
    ]
    add_table(doc, headers, rows, caption="Tableau 4. Chronogramme d'affectation en jours-hommes",
              col_widths=[3.5, 2.5, 2.5, 2.5, 2.5, 2.0])


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


def section_valeur_ajoutee(doc):
    add_heading(doc, "Valeur ajoutée et points distinctifs", 1, numbered="9")

    add_body(doc,
        "Au-delà de la stricte conformité aux exigences du CPS, l'offre apporte cinq points de "
        "valeur ajoutée directement ancrés dans les contraintes propres à ce marché."
    )

    add_bullet_dot(doc, [
        "Séparation OLTP/OLAP pour le Headless BI : les requêtes analytiques ne dégradent jamais le "
        "service rendu aux apprenants en période de pointe, notamment lors de la publication des "
        "résultats de concours sur les 58 établissements simultanément.",
        "Mise en cache adaptative des réponses du Bot APC par profil d'utilisateur, invalidée "
        "automatiquement à chaque mise à jour d'un référentiel pédagogique, pour sécuriser l'objectif "
        "de temps de réponse inférieur à 1,5 seconde au-delà des seules conditions de test.",
        "Migration par lots contrôlés avec somme de contrôle et fiche d'anomalie horodatée, qui "
        "permet d'isoler un incident de migration sur un seul lot sans suspendre l'intégration des "
        "lots déjà validés sur les autres EFPA.",
        "Intercepteur d'audit unique appliqué uniformément aux neuf services, qui garantit un format "
        "de journalisation homogène pour l'exploitation réelle du registre en cas d'incident de "
        "sécurité, plutôt qu'une juxtaposition de journaux hétérogènes service par service.",
        "Architecture en schéma partagé avec sécurité au niveau ligne, conçue dès l'origine pour "
        "l'extensibilité à 100 établissements annoncée par le CPS, sans migration de modèle de "
        "données lorsque ce seuil sera atteint.",
    ])


def section_gestion_risques(doc, risk_path):
    add_heading(doc, "Gestion des risques", 1, numbered="10")

    add_body(doc,
        "Les risques retenus sont spécifiques à ce marché et à son contexte d'exécution, et non des "
        "risques génériques de projet informatique."
    )
    add_figure(doc, risk_path, "Matrice de criticité des risques")

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
         "SDK officiels Microsoft (MSAL, Graph) qui ne requièrent pas ce partenariat pour "
         "l'intégration technique elle-même"],
        ["R4 - Hétérogénéité des données sources entre les 58 EFPA",
         "Dictionnaire de données et règles de transcodage par établissement, documentés dans le "
         "rapport de stratégie de migration ETL avant extraction"],
        ["R5 - Dépassement du SLA de disponibilité 99,95% en période de pointe",
         "Architecture en cluster Multi-AZ avec bascule automatique ; scalabilité horizontale des "
         "répliques API découplée des workers de certification"],
        ["R6 - Désynchronisation Moodle (doublons, groupes incohérents)",
         "Synchronisation idempotente avec réconciliation périodique automatisée entre le "
         "Référentiel et Moodle"],
    ]
    add_table(doc, headers, rows, caption="Tableau 5. Registre des risques et mesures de mitigation",
              col_widths=[6.5, 8.5])


# ---------------------------------------------------------------
# Assemblage du document
# ---------------------------------------------------------------

def build_document():
    print("Génération des figures...")
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
    gantt_path = generate_gantt()
    print(f"  Gantt: {gantt_path}")
    risk_path = generate_risk_matrix()
    print(f"  Matrice risques: {risk_path}")

    print("Construction du document...")
    doc = Document()
    set_margins(doc)
    set_document_font_theme(doc)
    add_page_numbers(doc)

    page_de_garde(doc)
    section_sommaire(doc)
    section_comprehension_contexte(doc)
    doc.add_page_break()
    section_comprehension_mission(doc)
    doc.add_page_break()
    section_approche_methodologique(doc)
    doc.add_page_break()
    section_architecture(doc, arch_path, usecase_path, classes_path, sequence_path)
    doc.add_page_break()
    section_equipe(doc, org_path)
    doc.add_page_break()
    section_planning(doc, gantt_path)
    doc.add_page_break()
    section_chronogramme(doc)
    doc.add_page_break()
    section_reversibilite(doc)
    doc.add_page_break()
    section_valeur_ajoutee(doc)
    doc.add_page_break()
    section_gestion_risques(doc, risk_path)

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
