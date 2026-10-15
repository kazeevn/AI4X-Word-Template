#!/usr/bin/env python3
"""Builds the AI4X 2027 Word abstract template, AI4X-2027.docx and AI4X-2027.dotx.

The layout follows the LaTeX template in AI4X-Abstract-Template (ai4x.cls): A4, 2 cm margins,
two columns 0.8 cm apart, Source Serif 10/12 pt. Source Serif 4 is embedded in the document, so it
does not need to be installed. Only the Python standard library is used:

    python3 build_template.py
"""
import uuid
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parent
FONTS = ROOT / 'fonts'
ASSETS = ROOT / 'assets'

VENUE = 'AI4X Conference 2027, Bremen, Germany, June 14–17'
TITLE = 'AI4X Conference 2027 abstract template and formatting guidelines'
AUTHORS = 'First Author, Presenting Author, Middle Author, Last Author'
KEYWORDS = 'AI for Science, AI4X, AI, machine learning, science'
FONT = 'Source Serif 4'
LANG = 'en-US'

# Page geometry, in twentieths of a point (dxa); 1 cm = 567 dxa
PAGE_W, PAGE_H = 11906, 16838
MARGIN = 1134
COL_SEP = 454
COL_W = (PAGE_W - 2 * MARGIN - COL_SEP) // 2
EMU_PER_DXA = 635

# Source Serif 4 has a 1.371 em line, LaTeX sets 10 pt type on 12 pt: 0.875 of a single line
LINE = 210

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'


# ------------------------------------------------------------------ relationships
class Rels:
    def __init__(self):
        self.items = []

    def add(self, kind, target, external=False):
        rid = f'rId{len(self.items) + 1}'
        self.items.append((rid, kind, target, external))
        return rid

    def xml(self):
        out = [XML_HEAD, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for rid, kind, target, external in self.items:
            mode = ' TargetMode="External"' if external else ''
            out.append(f'<Relationship Id="{rid}" Type="{REL}{kind}" Target="{escape(target, {chr(34): "&quot;"})}"{mode}/>')
        out.append('</Relationships>')
        return ''.join(out)


doc_rels = Rels()
for kind, target in [('styles', 'styles.xml'), ('settings', 'settings.xml'), ('numbering', 'numbering.xml'),
                     ('fontTable', 'fontTable.xml')]:
    doc_rels.add(kind, target)
HEADER_RID = doc_rels.add('header', 'header1.xml')
FOOTER_RID = doc_rels.add('footer', 'footer1.xml')


# ------------------------------------------------------------------ inline content
def run(text, b=False, i=False, u=False, sup=False, size=None, style=None, nbsp=False):
    props = ''
    if style:
        props += f'<w:rStyle w:val="{style}"/>'
    if b:
        props += '<w:b/><w:bCs/>'
    if i:
        props += '<w:i/><w:iCs/>'
    if u:
        props += '<w:u w:val="single"/>'
    if size:
        props += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    if sup:
        props += '<w:vertAlign w:val="superscript"/>'
    rpr = f'<w:rPr>{props}</w:rPr>' if props else ''
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def tab():
    return '<w:r><w:tab/></w:r>'


def br():
    return '<w:r><w:br/></w:r>'


def field(instr, result):
    """A complex field with its cached result, so it shows correctly before an update."""
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {escape(instr)} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'{result}<w:r><w:fldChar w:fldCharType="end"/></w:r>')


_bookmark_id = [0]


def bookmark(name, content):
    _bookmark_id[0] += 1
    n = _bookmark_id[0]
    return f'<w:bookmarkStart w:id="{n}" w:name="{name}"/>{content}<w:bookmarkEnd w:id="{n}"/>'


def ref(name, text):
    return field(f'REF {name} \\h', run(text))


def link(url, content):
    rid = doc_rels.add('hyperlink', url, external=True)
    return f'<w:hyperlink r:id="{rid}" w:history="1">{content}</w:hyperlink>'


def url_run(url, text=None):
    return link(url, run(text or url, style='Hyperlink'))


def cite(*numbers):
    """Numeric citations as plain text, like natbib's [2, 3]."""
    return run('[' + ', '.join(str(n) for n in numbers) + ']')


_drawing_id = [0]


def drawing(rid, cx, cy, name, descr, link_rid=None, position=None):
    _drawing_id[0] += 1
    n = _drawing_id[0]
    hlink = f'<a:hlinkClick r:id="{link_rid}"/>' if link_rid else ''
    rpr = f'<w:rPr><w:noProof/><w:position w:val="{position}"/></w:rPr>' if position else '<w:rPr><w:noProof/></w:rPr>'
    return (f'<w:r>{rpr}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{n}" name="{name}" descr="{escape(descr, {chr(34): "&quot;"})}">{hlink}</wp:docPr>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            f'<pic:nvPicPr><pic:cNvPr id="{n}" name="{name}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


ORCID_IMG = doc_rels.add('image', 'media/orcid-id.png')
ORCID_SIZE = 59000  # EMU, about the height of a capital letter in a 7 pt superscript, as in orcidlink


def orcid(orcid_id):
    link_rid = doc_rels.add('hyperlink', f'https://orcid.org/{orcid_id}', external=True)
    return drawing(ORCID_IMG, ORCID_SIZE, ORCID_SIZE, 'ORCID iD', f'ORCID iD {orcid_id}',
                   link_rid=link_rid, position=6)


# ------------------------------------------------------------------ paragraphs
def para(style, *content, ppr=''):
    style_xml = f'<w:pStyle w:val="{style}"/>' if style else ''
    return f'<w:p><w:pPr>{style_xml}{ppr}</w:pPr>{"".join(content)}</w:p>'


def text_para(style, text):
    return para(style, run(text))


# ------------------------------------------------------------------ math
def mr(text, plain=False, normal=False):
    sty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if plain else '<m:rPr><m:nor/></m:rPr>' if normal else ''
    return f'<m:r>{sty}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>{escape(text)}</m:t></m:r>'


def m_limit_term(a, result):
    """( lim_{x -> a^+} 1/(x - a) = result ); the + is normal text, as LibreOffice cannot import a lone math +"""
    lim = (f'<m:func><m:fName><m:limLow><m:e>{mr("lim", plain=True)}</m:e>'
           f'<m:lim>{mr("x→")}<m:sSup><m:e>{mr(a)}</m:e><m:sup>{mr("+", normal=True)}</m:sup></m:sSup></m:lim>'
           '</m:limLow></m:fName>'
           f'<m:e><m:f><m:num>{mr("1")}</m:num><m:den>{mr("x−" + a)}</m:den></m:f>{mr("=" + result)}</m:e></m:func>')
    return f'<m:d><m:e>{lim}</m:e></m:d>'


EQUATION = f'<m:oMath>{m_limit_term("8", "∞")}{mr("⇒")}{m_limit_term("3", "ω")}</m:oMath>'


# ------------------------------------------------------------------ table
def cell(width, content, align='left', borders='', mar='', keep_next=False):
    tcpr = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if borders:
        tcpr += f'<w:tcBorders>{borders}</w:tcBorders>'
    if mar:
        tcpr += f'<w:tcMar>{mar}</w:tcMar>'
    jc = '<w:jc w:val="center"/>' if align == 'center' else ''
    keep = '<w:keepNext/>' if keep_next else ''
    return (f'<w:tc><w:tcPr>{tcpr}</w:tcPr><w:p><w:pPr><w:pStyle w:val="TableText"/>'
            f'{keep}<w:keepLines/>{jc}</w:pPr>{content}</w:p></w:tc>')


def results_table():
    widths = [1572, 1220, 800, 1000]
    assert sum(widths) == COL_W
    heavy = '<w:{side} w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
    light = '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    top_pad, bottom_pad = '<w:top w:w="57" w:type="dxa"/>', '<w:bottom w:w="57" w:type="dxa"/>'
    header = [run('Model'),
              run('MAE ↓') + br() + run('meV/atom'),
              run('R', i=True) + run('2', sup=True) + run(' ↑'),
              run('Time ↓') + br() + run('ms')]
    rows = [
        [run('Random forest'), run('85 ± 3'), run('0.91'), run('0.4', b=True)],
        [run('GNN'), run('33 ± 2'), run('0.98'), run('2.1')],
        [run('Ours'), run('28 ± 1', b=True), run('0.99', b=True), run('1.5')],
    ]
    out = ['<w:tbl><w:tblPr><w:tblStyle w:val="Booktabs"/>'
           f'<w:tblW w:w="{COL_W}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
           '<w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" '
           'w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid>']
    out += [f'<w:gridCol w:w="{w}"/>' for w in widths]
    out.append('</w:tblGrid>')
    # Keep this short example table together, without tying its final row to the next section.
    out.append('<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>')
    for k, (w, c) in enumerate(zip(widths, header)):
        out.append(cell(w, c, 'left' if k == 0 else 'center',
                        heavy.format(side='top') + light, top_pad + bottom_pad, keep_next=True))
    out.append('</w:tr>')
    for r, row in enumerate(rows):
        out.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>')
        mar = (top_pad if r == 0 else '') + (bottom_pad if r == len(rows) - 1 else '')
        border = heavy.format(side='bottom') if r == len(rows) - 1 else ''
        for k, (w, c) in enumerate(zip(widths, row)):
            out.append(cell(w, c, 'left' if k == 0 else 'center', border, mar,
                            keep_next=r < len(rows) - 1))
        out.append('</w:tr>')
    out.append('</w:tbl>')
    return ''.join(out)


# ------------------------------------------------------------------ document body
def sect_pr(columns, continuous):
    cols = f'<w:cols w:num="2" w:space="{COL_SEP}"/>' if columns == 2 else '<w:cols w:space="720"/>'
    kind = '<w:type w:val="continuous"/>' if continuous else ''
    return (f'<w:sectPr><w:headerReference w:type="default" r:id="{HEADER_RID}"/>'
            f'<w:footerReference w:type="default" r:id="{FOOTER_RID}"/>{kind}'
            f'<w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}"/>'
            f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}" '
            'w:header="550" w:footer="500" w:gutter="0"/>'
            f'{cols}<w:docGrid w:linePitch="240"/></w:sectPr>')


FIG = '_Ref270000001'
TAB = '_Ref270000002'
EQ = '_Ref270000003'
REF_AUDIT = 'https://github.com/constructorfabric/reference-audit/'
LATEX_TEMPLATE = 'https://github.com/kazeevn/AI4X-Abstract-Template'
IAC_TEMPLATE = 'https://github.com/grande-dev/International-Astronautical-Congress-LaTeX-template'

FIGURE_IMG = doc_rels.add('image', 'media/bremen-lab.jpeg')


def body():
    p = []
    # ---- title block, one column
    p.append(text_para('Title', TITLE))
    p.append(para('Authors',
                  run('First Author', b=True), run('\u2009'), orcid('0000-0000-0000-0000'),
                  run('*\u20091', sup=True), run(' '),
                  run('Presenting Author', b=True, u=True), run('\u2009'), run('*\u20092', sup=True), run(' '),
                  run('Middle Author', b=True), run('\u2009'), orcid('0000-0000-0000-0000'),
                  run('2\u20091', sup=True), run(' '),
                  run('Last Author', b=True), run('\u2009'), run('3', sup=True)))
    # Two balanced lines, with hyphenation suppressed, keep each affiliation intact in Word.
    def keep(text):
        return text.replace(' ', '\u00a0')
    p.append(para('Affiliations',
                  run('*', sup=True), run(keep('Equal contribution') + '\u2003'),
                  run('1', sup=True), run(keep('Olympus Academy of Mountaineering, Mytikas, Greece 402 00'), i=True),
                  br(),
                  run('2', sup=True), run(keep('University of Underwater Studies, Atlantis 341'), i=True),
                  run('\u2003'),
                  run('3', sup=True), run(keep('University of Nowhere'), i=True)))
    p.append(para('Affiliations',
                  run('Correspondence to: Last Author '),
                  link('mailto:dropides@void.ai', run('dropides@void.ai', style='Hyperlink', u=True)),
                  run(', Middle Author '),
                  link('mailto:boss@underworld.ai', run('boss@underworld.ai', style='Hyperlink', u=True)),
                  ppr=sect_pr(1, continuous=False)))

    # ---- two-column body
    p.append(text_para('Heading1', 'Introduction'))
    p.append(text_para('FirstParagraph',
                       'The page limit is two pages, not counting the acknowledgments, references, and appendices. '
                       'Please do not change the font sizes, margins or line spacing to squeeze more text.'))
    p.append(text_para('BodyText',
                       'The review is single-blind: the reviewers see the author names, so do not anonymize your '
                       'submission. At least one author must be listed after “Correspondence to”. ORCID links '
                       'for the authors will make it easier for the readers to find you, but are not mandatory; to '
                       'link an iD icon to your record, right-click it and choose Link.'))
    p.append(text_para('BodyText',
                       'AI4X is non-archival: an accepted abstract is not a formal publication, so you are free to '
                       'publish the work later in a journal or in conference proceedings.'))

    p.append(text_para('Heading1', 'Substantial section'))
    p.append(para('FirstParagraph',
                  run('If you want to, you can split the text into sections. An AI4X abstract is not a journal paper, '
                      'so there is no hard requirement for the work to be reproducible from the text. At the same '
                      'time, the reviewers and, ultimately, your fellow attendees look for novelty, significance, and '
                      'soundness – the abstract must convey what was done, why, and the validation results. It '
                      'can include equations, such as Equation ('),
                  ref(EQ, '1'),
                  run('). If the abstract is based on your previously published work, it should be clearly stated '
                      'and cited.')))
    p.append(para('Equation', tab(), EQUATION, tab(), run('('),
                  bookmark(EQ, field('SEQ Equation \\* ARABIC', run('1'))), run(')')))

    p.append(text_para('Heading2', 'Related work'))
    p.append(para('FirstParagraph',
                  run('A good way to show novelty is to state the difference from the related work, for example, on '
                      'protein structure prediction '), cite(1),
                  run(', graph neural networks (GNNs) for molecules and materials '), cite(2, 3),
                  run(', materials discovery at scale '), cite(4),
                  run(', or autonomous laboratories '), cite(5),
                  run('. An overview of the previous state of the field can be found in the AI4X Roadmap from a '
                      'previous conference edition '), cite(6), run('.')))

    p.append(text_para('Heading2', 'Figures and tables'))
    p.append(para('FirstParagraph',
                  ref(FIG, 'Figure 1'), run(' and '), ref(TAB, 'Table 1'),
                  run(' are single-column, but two-column layout can also be used: put the figure or table in a '
                      'one-column section between two continuous section breaks (Layout → Breaks, '
                      'Layout → Columns).')))
    size = COL_W * EMU_PER_DXA
    p.append(para('Figure', drawing(FIGURE_IMG, size, size, 'Figure 1',
                                    'The Bremen Town Musicians, a donkey, a dog, a cat, and a rooster, '
                                    'operating an automated laboratory')))
    p.append(para('Caption',
                  bookmark(FIG, run('Figure ') + field('SEQ Figure \\* ARABIC', run('1'))),
                  run(': Collaboration was just as important to the success of the study as automation. The '
                      'illustration was generated with Google Gemini (Nano Banana), and its defects are kept on '
                      'purpose.')))
    p.append(para('Caption',
                  bookmark(TAB, run('Table ') + field('SEQ Table \\* ARABIC', run('1'))),
                  run(': Formation energy prediction on a held-out test set: mean absolute error (MAE, mean ± '
                      'standard deviation over five random seeds), '),
                  run('R', i=True), run('2', sup=True),
                  run(', and inference time per structure. The best values are in bold; the numbers are '
                      'illustrative.'),
                  ppr='<w:keepNext/><w:keepLines/>'))
    p.append(results_table())

    # Keep the example content ahead of Word-specific instructions, as in the LaTeX layout.
    p.append(text_para('Heading1', 'Working with the template'))
    p.append(text_para('FirstParagraph',
                       'Use the styles of the template: Heading 1 and Heading 2 number the sections, Unnumbered '
                       'Heading is for the acknowledgments and references. Use Appendix Heading for appendices, '
                       'and Appendix Heading 2 and Appendix Heading 3 for their subsections. '
                       'The numbers of the figures, tables, and equations, and the references to them, are fields: '
                       'after moving things around, select all (Ctrl+A) and press F9 to update them.'))
    p.append(text_para('BodyText',
                       'The template embeds its font, Source Serif 4, so it looks the same on computers that do not '
                       'have the font installed. Please submit a PDF, and check in it that the bold and italic text '
                       'and the equations look right. Fill the title, the authors, and the keywords in '
                       'File → Info → Properties: they become the PDF metadata, which improves search '
                       'discoverability.'))

    p.append(text_para('UnnumberedHeading', 'AI Acknowledgments'))
    p.append(para('FirstParagraph',
                  ref(FIG, 'Figure 1'),
                  run(' was generated with Google Gemini (Nano Banana), and this Word template was converted from '
                      'the LaTeX template '), cite(8), run(' with Claude Code (Claude Opus 5.5).')))
    p.append(para('BodyText',
                  run('At AI4X, we welcome the new AI age and believe that, used well, AI can do much more than '
                      'produce “AI slop”. You are welcome to share how you used AI to prepare the work, be '
                      'it grammar checking or doing the study end-to-end. This is not mandatory and is not included '
                      'in the page limit. Whatever tools you use, the authors are fully responsible for the '
                      'submission, including any text, code, figures, and references produced with AI; in '
                      'particular, check that every cited work exists '), cite(7),
                  run('. You can use '), url_run(REF_AUDIT, 'reference-audit'),
                  run(' to run the checks yourself on the PDF of your abstract.')))

    p.append(text_para('UnnumberedHeading', 'Acknowledgments'))
    p.append(para('FirstParagraph',
                  run('Acknowledgments and references are not numbered and not included in the page limit – '
                      'feel free to give credit where it’s due. This Word template follows the AI4X 2027 LaTeX '
                      'template '), cite(8),
                  run(', which is based on the unofficial International Astronautical Congress template '), cite(9),
                  run('.')))

    p.append(text_para('UnnumberedHeading', 'References'))
    for n, entry in enumerate(REFERENCES, 1):
        p.append(para('Bibliography', run(f'[{n}]'), tab(), *entry))

    p.append(text_para('AppendixHeading', 'My title'))
    p.append(text_para('FirstParagraph',
                       'A paper may contain an unlimited number of appendices which don’t count towards the page '
                       'limit, but the reviewers are not required to read them.'))
    return ''.join(p) + sect_pr(2, continuous=True)


def doi(d):
    return url_run(f'https://doi.org/{d}', f'doi:{d}')


# The entries as ai4x.bst prints them
REFERENCES = [
    [run('John Jumper et al. Highly accurate protein structure prediction with AlphaFold. '),
     run('Nature', i=True), run(', 596(7873):583–589, 2021. '), doi('10.1038/s41586-021-03819-2'), run('.')],
    [run('Kristof Schütt, Pieter-Jan Kindermans, Huziel Enoc Sauceda Felix, Stefan Chmiela, Alexandre '
         'Tkatchenko, and Klaus-Robert Müller. SchNet: A continuous-filter convolutional neural network for '
         'modeling quantum interactions. In '),
     run('Advances in Neural Information Processing Systems', i=True),
     run(', volume 30. Curran Associates, Inc., 2017. URL: '),
     url_run('https://proceedings.neurips.cc/paper_files/paper/2017/file/'
             '303ed4c69846ab36c2904d3ba8573050-Paper.pdf'), run('.')],
    [run('Tian Xie and Jeffrey C. Grossman. Crystal graph convolutional neural networks for an accurate and '
         'interpretable prediction of material properties. '),
     run('Physical Review Letters', i=True), run(', 120(14):145301, 2018. '),
     doi('10.1103/PhysRevLett.120.145301'), run('.')],
    [run('Amil Merchant, Simon Batzner, Samuel S. Schoenholz, Muratahan Aykol, Gowoon Cheon, and Ekin Dogus '
         'Cubuk. Scaling deep learning for materials discovery. '),
     run('Nature', i=True), run(', 624(7990):80–85, 2023. '), doi('10.1038/s41586-023-06735-9'), run('.')],
    [run('Benjamin Burger et al. A mobile robotic chemist. '),
     run('Nature', i=True), run(', 583(7815):237–241, 2020. '), doi('10.1038/s41586-020-2442-2'), run('.')],
    [run('Stephen G. Dale, Nikita Kazeev, Alastair J. A. Price, Victor Posligua, Stephan Roche, O. Anatole von '
         'Lilienfeld, Konstantin S. Novoselov, Xavier Bresson, Gianmarco Mengaldo, Xudong Chen, Terence J. '
         'O’Kane, Emily R. Lines, Matthew J. Allen, Amandine E. Debus, Clayton Miller, Jiayu Zhou, Hiroko H. '
         'Dodge, David Rousseau, Andrey Ustyuzhanin, Ziyun Yan, Mario Lanza, Fabio Sciarrino, Ryo Yoshida, Zhidong '
         'Leong, Teck Leong Tan, Qianxiao Li, Adil Kabylda, Igor Poltavsky, Alexandre Tkatchenko, Sherif Abdulkader '
         'Tawfik, Prathami Divakar Kamath, Theo Jaffrelot Inizan, Kristin A. Persson, Bryant Y. Li, Vir Karan, '
         'Chenru Duan, Haojun Jia, Qiyuan Zhao, Hiroyuki Hayashi, Atsuto Seko, Isao Tanaka, Omar M. Yaghi, Tim '
         'Gould, Bun Chan, Stefan Vuckovic, Tianbo Li, Min Lin, Zehcen Tang, Yang Li, Yong Xu, Amrita Joshi, '
         'Xiaonan Wang, Leonard W. T. Ng, Sergei V. Kalinin, Mahshid Ahmadi, Jiyizhe Zhang, Shuyuan Zhang, Alexei '
         'Lapkin, Ming Xiao, Zhe Wu, Kedar Hippalgaonkar, Limsoon Wong, Lorenzo Bastonero, Nicola Marzari, Dorye '
         'Luis Esteras Cordoba, Andrei Tomut, Alba Quinones Andrade, and Jose-Hugo Garcia. Ai4x roadmap: '
         'Artificial intelligence for the advancement of scientific pursuit and its future directions, 2025. '
         'URL: '), url_run('https://arxiv.org/abs/2511.20976'), run(', '),
     url_run('https://arxiv.org/abs/2511.20976', 'arXiv:2511.20976'), run('.')],
    [run('AI4X Program Committee. Hallucinated citations and Desk Rejections – is there a connection? '),
     run('Nature', i=True), run(', 2028.')],
    [run('Nikita Kazeev. AI4X Conference 2027 abstract template. '), url_run(LATEX_TEMPLATE), run(', 2026.')],
    [run('John O. Woods, Davide Grande, and Stefano Farì. International Astronautical Congress LaTeX '
         'template. '), url_run(IAC_TEMPLATE), run(', 2025.')],
]


def document_xml():
    return f'{XML_HEAD}<w:document {NS}><w:body>{body()}</w:body></w:document>'


# ------------------------------------------------------------------ styles
def spacing(before=0, after=0, line=LINE):
    return f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'


def pstyle(sid, name, ppr='', rpr='', based='Normal', nxt=None, extra='', builtin=True):
    custom = '' if builtin else ' w:customStyle="1"'
    based_xml = f'<w:basedOn w:val="{based}"/>' if based else ''
    next_xml = f'<w:next w:val="{nxt}"/>' if nxt else ''
    return (f'<w:style w:type="paragraph"{custom} w:styleId="{sid}"><w:name w:val="{name}"/>{based_xml}{next_xml}'
            f'{extra}<w:qFormat/><w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>')


def styles_xml():
    fonts = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/>'
    def heading_ppr(num=''):
        return f'<w:keepNext/><w:keepLines/>{num}' + spacing(300, 20)
    styles = [
        f'<w:docDefaults><w:rPrDefault><w:rPr>{fonts}<w:sz w:val="20"/><w:szCs w:val="20"/>'
        f'<w:lang w:val="{LANG}" w:eastAsia="{LANG}" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
        f'<w:pPrDefault><w:pPr>{spacing()}</w:pPr></w:pPrDefault></w:docDefaults>',
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>'
        f'<w:pPr><w:widowControl/>{spacing()}<w:jc w:val="both"/></w:pPr><w:rPr>{fonts}<w:sz w:val="20"/>'
        '<w:szCs w:val="20"/></w:rPr></w:style>',
        '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>',
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>',
        '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>',
        pstyle('Title', 'Title', '<w:keepNext/>' + spacing(270, 300) + '<w:jc w:val="center"/>',
               '<w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/>', nxt='Authors'),
        pstyle('Authors', 'Authors', '<w:keepNext/>' + spacing(0, 110, 240) + '<w:jc w:val="center"/>',
               nxt='Affiliations', builtin=False),
        pstyle('Affiliations', 'Affiliations', '<w:keepNext/><w:suppressAutoHyphens/>'
               + spacing(0, 0, 214) + '<w:jc w:val="center"/>',
               '<w:sz w:val="18"/><w:szCs w:val="18"/>', builtin=False),
        pstyle('FirstParagraph', 'First Paragraph', nxt='BodyText', builtin=False),
        pstyle('BodyText', 'Body Text', '<w:ind w:firstLine="240"/>', nxt='BodyText'),
        pstyle('Heading1', 'heading 1',
               heading_ppr('<w:numPr><w:numId w:val="1"/></w:numPr>') + '<w:jc w:val="left"/><w:outlineLvl w:val="0"/>',
               '<w:b/><w:bCs/>', nxt='FirstParagraph'),
        pstyle('Heading2', 'heading 2',
               heading_ppr('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>') + '<w:jc w:val="left"/>'
               '<w:outlineLvl w:val="1"/>', '<w:i/><w:iCs/>', nxt='FirstParagraph'),
        pstyle('Heading3', 'heading 3',
               heading_ppr('<w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr>') + '<w:jc w:val="left"/>'
               '<w:outlineLvl w:val="2"/>', '<w:i/><w:iCs/>', nxt='FirstParagraph'),
        pstyle('UnnumberedHeading', 'Unnumbered Heading',
               heading_ppr() + '<w:jc w:val="left"/><w:outlineLvl w:val="0"/>', '<w:b/><w:bCs/>',
               nxt='FirstParagraph', builtin=False),
        pstyle('AppendixHeading', 'Appendix Heading',
               heading_ppr('<w:numPr><w:numId w:val="2"/></w:numPr>') + '<w:jc w:val="left"/><w:outlineLvl w:val="0"/>',
               '<w:b/><w:bCs/>', nxt='FirstParagraph', builtin=False),
        pstyle('AppendixHeading2', 'Appendix Heading 2',
               heading_ppr('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>')
               + '<w:jc w:val="left"/><w:outlineLvl w:val="1"/>',
               '<w:i/><w:iCs/>', nxt='FirstParagraph', builtin=False),
        pstyle('AppendixHeading3', 'Appendix Heading 3',
               heading_ppr('<w:numPr><w:ilvl w:val="2"/><w:numId w:val="2"/></w:numPr>')
               + '<w:jc w:val="left"/><w:outlineLvl w:val="2"/>',
               '<w:i/><w:iCs/>', nxt='FirstParagraph', builtin=False),
        pstyle('Caption', 'caption', spacing(200, 200) + '<w:ind w:left="200" w:hanging="200"/>',
               nxt='BodyText'),
        pstyle('Figure', 'Figure', '<w:keepNext/>' + spacing(120, 0, 240) + '<w:jc w:val="center"/>',
               nxt='Caption', builtin=False),
        pstyle('Equation', 'Equation',
               f'<w:tabs><w:tab w:val="center" w:pos="{COL_W // 2}"/><w:tab w:val="right" w:pos="{COL_W}"/>'
               '</w:tabs>' + spacing(160, 160, 240) + '<w:jc w:val="left"/>',
               nxt='BodyText', builtin=False),
        pstyle('TableText', 'Table Text', '<w:jc w:val="left"/>', builtin=False),
        pstyle('Bibliography', 'Bibliography', '<w:tabs><w:tab w:val="left" w:pos="340"/></w:tabs>'
               + spacing(0, 80) + '<w:ind w:left="340" w:hanging="340"/><w:jc w:val="left"/>',
               nxt='Bibliography'),
        pstyle('Header', 'header', spacing() + '<w:jc w:val="center"/>', '<w:sz w:val="16"/><w:szCs w:val="16"/>'),
        pstyle('Footer', 'footer', spacing() + '<w:jc w:val="center"/>', '<w:sz w:val="16"/><w:szCs w:val="16"/>'),
        # Links are not coloured, as in the LaTeX template (hyperref with colorlinks=false)
        '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>'
        '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr/></w:style>',
        # booktabs: heavy top and bottom rules, a light rule under the header row
        '<w:style w:type="table" w:customStyle="1" w:styleId="Booktabs"><w:name w:val="Booktabs"/>'
        '<w:basedOn w:val="TableNormal"/><w:qFormat/><w:pPr><w:jc w:val="left"/></w:pPr><w:tblPr>'
        '<w:tblBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tblBorders>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        '<w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders></w:tcPr></w:tblStylePr>'
        '</w:style>',
    ]
    return (f'{XML_HEAD}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + ''.join(styles) + '</w:styles>')


def numbering_xml():
    def lvl(i, fmt, text, style):
        return (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:pStyle w:val="{style}"/>'
                f'<w:suff w:val="space"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>')
    headings = (lvl(0, 'decimal', '%1.', 'Heading1') + lvl(1, 'decimal', '%1.%2.', 'Heading2')
                + lvl(2, 'decimal', '%1.%2.%3.', 'Heading3'))
    appendices = (lvl(0, 'upperLetter', 'Appendix %1.', 'AppendixHeading')
                  + lvl(1, 'decimal', '%1.%2.', 'AppendixHeading2')
                  + lvl(2, 'decimal', '%1.%2.%3.', 'AppendixHeading3'))
    return (f'{XML_HEAD}<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>{headings}</w:abstractNum>'
            f'<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/>{appendices}</w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
            '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>')


def settings_xml():
    return (f'{XML_HEAD}<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            '<w:zoom w:percent="100"/><w:embedTrueTypeFonts/><w:defaultTabStop w:val="720"/>'
            '<w:autoHyphenation/><w:characterSpacingControl w:val="doNotCompress"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '</w:compat>'
            '<m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/>'
            '<m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/>'
            '<m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/>'
            '<m:naryLim m:val="undOvr"/></m:mathPr>'
            f'<w:themeFontLang w:val="{LANG}"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>'
            '</w:settings>')


def header_xml():
    return (f'{XML_HEAD}<w:hdr {NS}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>{run(VENUE)}</w:p></w:hdr>')


def footer_xml():
    return (f'{XML_HEAD}<w:ftr {NS}><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>'
            f'{field("PAGE", run("1"))}</w:p></w:ftr>')


# ------------------------------------------------------------------ embedded fonts
FONT_FILES = [('Regular', 'SourceSerif4-Regular.ttf'), ('Bold', 'SourceSerif4-Bold.ttf'),
              ('Italic', 'SourceSerif4-It.ttf'), ('BoldItalic', 'SourceSerif4-BoldIt.ttf')]


def obfuscate(data, guid):
    """ECMA-376 Part 1, 17.8.1: the first 32 bytes are XORed with the GUID key, its bytes in reverse order."""
    h = guid.replace('-', '')
    key = bytes(int(h[30 - 2 * i:32 - 2 * i], 16) for i in range(16))
    out = bytearray(data)
    for i in range(32):
        out[i] ^= key[i % 16]
    return bytes(out)


def embedded_fonts():
    rels = Rels()
    embeds, files = [], {}
    for n, (kind, filename) in enumerate(FONT_FILES, 1):
        # A fixed GUID per font keeps the build reproducible
        guid = str(uuid.uuid5(uuid.NAMESPACE_URL, f'ai4x-2027/{filename}')).upper()
        rid = rels.add('font', f'fonts/font{n}.odttf')
        embeds.append(f'<w:embed{kind} r:id="{rid}" w:fontKey="{{{guid}}}"/>')
        files[f'word/fonts/font{n}.odttf'] = obfuscate((FONTS / filename).read_bytes(), guid)
    table = (f'{XML_HEAD}<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
             f'<w:font w:name="{FONT}"><w:panose1 w:val="02040603050405020204"/><w:charset w:val="00"/>'
             '<w:family w:val="roman"/><w:pitch w:val="variable"/>'
             f'{"".join(embeds)}</w:font>'
             '<w:font w:name="Cambria Math"><w:panose1 w:val="02040503050406030204"/><w:charset w:val="00"/>'
             '<w:family w:val="roman"/><w:pitch w:val="variable"/></w:font></w:fonts>')
    files['word/fontTable.xml'] = table
    files['word/_rels/fontTable.xml.rels'] = rels.xml()
    return files


# ------------------------------------------------------------------ package
def content_types(template):
    main = 'template.main+xml' if template else 'document.main+xml'
    wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
    overrides = {
        '/word/document.xml': wml + main, '/word/styles.xml': wml + 'styles+xml',
        '/word/settings.xml': wml + 'settings+xml', '/word/numbering.xml': wml + 'numbering+xml',
        '/word/fontTable.xml': wml + 'fontTable+xml', '/word/header1.xml': wml + 'header+xml',
        '/word/footer1.xml': wml + 'footer+xml',
        '/docProps/core.xml': 'application/vnd.openxmlformats-package.core-properties+xml',
    }
    return (f'{XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>'
            '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>'
            + ''.join(f'<Override PartName="{k}" ContentType="{v}"/>' for k, v in overrides.items())
            + '</Types>')


def core_xml():
    return (f'{XML_HEAD}<cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{escape(TITLE)}</dc:title><dc:subject>{escape(VENUE)}</dc:subject>'
            f'<dc:creator>{escape(AUTHORS)}</dc:creator><cp:keywords>{escape(KEYWORDS)}</cp:keywords>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-14T00:00:00Z</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-14T00:00:00Z</dcterms:modified>'
            '</cp:coreProperties>')


def package_rels():
    rels = Rels()
    rels.add('officeDocument', 'word/document.xml')
    rels.items.append(('rId2', None, 'docProps/core.xml', False))
    xml = rels.xml().replace(
        f'Type="{REL}None"',
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"')
    return xml


def parts():
    files = {
        '_rels/.rels': package_rels(),
        'docProps/core.xml': core_xml(),
        'word/document.xml': document_xml(),
        'word/styles.xml': styles_xml(),
        'word/numbering.xml': numbering_xml(),
        'word/settings.xml': settings_xml(),
        'word/header1.xml': header_xml(),
        'word/footer1.xml': footer_xml(),
        'word/media/orcid-id.png': (ASSETS / 'orcid-id.png').read_bytes(),
        'word/media/bremen-lab.jpeg': (ASSETS / 'bremen-lab.jpeg').read_bytes(),
        **embedded_fonts(),
    }
    # After document_xml(), which registers the hyperlinks
    files['word/_rels/document.xml.rels'] = doc_rels.xml()
    return files


def write(path, files, template):
    files = {'[Content_Types].xml': content_types(template), **files}
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            info = zipfile.ZipInfo(name, date_time=(2026, 10, 14, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data.encode('utf-8') if isinstance(data, str) else data)
    print('wrote', path.name)


if __name__ == '__main__':
    files = parts()
    write(ROOT / 'AI4X-2027.docx', files, template=False)
    write(ROOT / 'AI4X-2027.dotx', files, template=True)
