from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZipFile

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree
from xlsxwriter import Workbook

from .extract import STATUS
from .ledger import status_counts


FONT = "Microsoft YaHei"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def xml(tag: str, **attrs):
    node = OxmlElement(f"w:{tag}")
    for key, value in attrs.items():
        node.set(qn(f"w:{key}"), str(value))
    return node


def add_chart(doc, counts: list[tuple[str, int]]) -> None:
    buffer = BytesIO()
    with Workbook(buffer, {"in_memory": True}) as book:
        sheet = book.add_worksheet("状态")
        sheet.write_row(0, 0, ("状态", "行动项数"))
        for index, pair in enumerate(counts, start=1):
            sheet.write_row(index, 0, pair)
        chart = book.add_chart({"type": "bar"})
        chart.add_series({
            "categories": [sheet.name, 1, 0, len(counts), 0],
            "values": [sheet.name, 1, 1, len(counts), 1],
            "fill": {"color": "#32658A"}, "border": {"none": True},
            "data_labels": {"value": True, "position": "outside_end"},
        })
        axis_font = {"name": FONT, "size": 10}
        chart.set_y_axis({"reverse": True, "num_font": axis_font, "major_tick_mark": "none",
                          "minor_tick_mark": "none", "line": {"none": True}})
        maximum = max(value for _, value in counts)
        chart.set_x_axis({"min": 0, "max": maximum + 1, "visible": False,
                          "major_gridlines": {"visible": False}})
        chart.set_legend({"none": True})
        chart.set_chartarea({"border": {"none": True}})
        sheet.insert_chart("D2", chart)

    with ZipFile(buffer) as archive:
        chart_xml = etree.fromstring(archive.read("xl/charts/chart1.xml"))
    labels = chart_xml.find(f".//{{{CHART_NS}}}dLbls")
    show_value = labels.find(f"{{{CHART_NS}}}showVal")
    show_value.addprevious(etree.Element(f"{{{CHART_NS}}}showLegendKey", val="0"))
    for flag in ("showCatName", "showSerName", "showPercent", "showBubbleSize"):
        etree.SubElement(labels, f"{{{CHART_NS}}}{flag}", val="0")
    package = doc.part.package
    workbook = Part(PackURI("/word/embeddings/action-status.xlsx"),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer.getvalue(), package)
    chart_part = Part(PackURI("/word/charts/chart1.xml"),
                      "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", package=package)
    external = etree.SubElement(chart_xml, f"{{{CHART_NS}}}externalData")
    external.set(f"{{{REL_NS}}}id", chart_part.relate_to(workbook, RT.PACKAGE))
    etree.SubElement(external, f"{{{CHART_NS}}}autoUpdate", val="0")
    chart_part._blob = etree.tostring(chart_xml, encoding="UTF-8", xml_declaration=True)
    rel_id = doc.part.relate_to(chart_part, RT.CHART)
    height = Inches(max(1.35, 0.8 + 0.27 * len(counts)))
    width = Inches(6.65)
    drawing = parse_xml(f'''<w:drawing {nsdecls("w", "wp", "a", "r")} xmlns:c="{CHART_NS}">
      <wp:inline><wp:extent cx="{width}" cy="{height}"/>
      <wp:docPr id="1" name="行动项状态分布"/>
      <a:graphic><a:graphicData uri="{CHART_NS}"><c:chart r:id="{rel_id}"/>
      </a:graphicData></a:graphic></wp:inline></w:drawing>''')
    drawing.find(f".//{qn('wp:docPr')}").set("descr", "，".join(f"{label} {count} 项" for label, count in counts))
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    p.add_run()._r.append(drawing)


def link_to_source(paragraph, number: int) -> None:
    link = xml("hyperlink", anchor=f"P{number:04d}")
    run = xml("r")
    props = xml("rPr")
    props.append(xml("color", val="234A70"))
    props.append(xml("u", val="single"))
    run.append(props)
    text = xml("t")
    text.text = f"P{number:04d}"
    run.append(text)
    link.append(run)
    paragraph._p.append(link)


def action_table(doc, items: list[dict]) -> None:
    table = doc.add_table(rows=1, cols=5)
    table.autofit = False
    widths = (2.67, 0.80, 1.16, 1.1, 0.97)
    for column, width in zip(table.columns, widths):
        column.width = Inches(width)
    borders = xml("tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(xml(edge, val="single", sz=4, color="D9D9D9"))
    table._tbl.tblPr.append(borders)
    table.rows[0]._tr.get_or_add_trPr().append(xml("tblHeader", val="true"))
    data = [("行动项", "负责人", "截止日", "状态", "原文")]
    for item in items:
        snap = item["snapshot"]
        data.append((f"{item['action_id']}  {item['text']}", snap["owner"] or "待确认",
                     snap["due"] or "待确认", STATUS[snap["status"]], item["paragraph"]))
    for index, cells in enumerate(data):
        row = table.rows[0] if index == 0 else table.add_row()
        for col, (cell, width, value) in enumerate(zip(row.cells, widths, cells)):
            cell.width = Inches(width)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            if col:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if isinstance(value, int):
                link_to_source(p, value)
            else:
                p.add_run(value).bold = index == 0
            if index == 0:
                p.runs[0].font.color.rgb = RGBColor.from_string("000000")
                cell._tc.get_or_add_tcPr().append(xml("shd", fill="E7EEF6"))


def write_word(path: Path, title: str, day: str, items: list[dict], revision: int = 1) -> None:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = section.right_margin = Inches(0.9)
    section.top_margin = section.bottom_margin = Inches(0.7)
    grid = section._sectPr.find(qn("w:docGrid"))
    if grid is not None:
        section._sectPr.remove(grid)
    for name, size in (("Title", 25), ("Heading 1", 13), ("Normal", 11)):
        style = doc.styles[name]
        style.font.name = FONT
        style.font.size = Pt(size)
        style.font.bold = name != "Normal"
        style.font.color.rgb = RGBColor.from_string("000000" if name != "Normal" else "25313D")
        fonts = style.element.get_or_add_rPr().rFonts
        for attr in tuple(fonts.attrib):
            if "theme" in attr.lower():
                del fonts.attrib[attr]
        fonts.set(qn("w:eastAsia"), FONT)
        style.paragraph_format.space_after = Pt(5)
        style.paragraph_format.line_spacing = Pt(size + 6)
        if name != "Normal":
            style.paragraph_format.keep_with_next = True
        props = style.element.get_or_add_pPr()
        border = props.find(qn("w:pBdr"))
        if border is not None:
            props.remove(border)
    doc.styles["Heading 1"].paragraph_format.space_before = Pt(12)
    title = " ".join("".join(c if c.isalnum() else " " for c in title).split())
    heading = title if title.endswith("会议纪要") else f"{title}会议纪要"
    doc.core_properties.title = heading
    doc.add_paragraph(heading, "Title")
    actions = list({item["action_id"]: item for item in items if item["kind"] == "action"}.values())
    decisions = [item for item in items if item["kind"] == "decision"]
    questions = [item for item in items if item["kind"] == "question"]
    topics = [item for item in items if item["kind"] == "topic"]
    pending = [item for item in questions if item.get("operation") != "close"]
    closed = [item for item in questions if item.get("operation") == "close"]
    doc.add_paragraph(f"{day}   ·   第 {revision} 版   ·   {len(decisions)} 项决定   ·   "
                      f"{len(actions)} 项行动   ·   {len(pending)} 项待确认")
    overview = topics[:3] or decisions[:1]
    detailed_decisions = decisions if topics else decisions[1:]
    for label, group in (("会议摘要", overview), ("会议决定", detailed_decisions),
                         ("行动项", actions), ("待确认事项", pending), ("已关闭事项", closed)):
        if not group:
            continue
        doc.add_heading(label, level=1)
        if label == "行动项":
            action_table(doc, actions)
            doc.add_heading("行动项状态分布", level=1)
            add_chart(doc, status_counts(actions))
        else:
            for item in group:
                prefix = f"{item['question_id']}  " if "question_id" in item else ""
                p = doc.add_paragraph(prefix + item["text"] + "  ")
                link_to_source(p, item["paragraph"])
    excerpts = {}
    for item in items:
        excerpts.setdefault(item["paragraph"], {})[item["quote"]] = None
    if excerpts:
        doc.add_heading("原文摘录", level=1)
        for number, quotes in excerpts.items():
            p = doc.add_paragraph()
            p._p.append(xml("bookmarkStart", id=number, name=f"P{number:04d}"))
            p.add_run(f"P{number:04d}  ").bold = True
            p._p.append(xml("bookmarkEnd", id=number))
            p.add_run(" / ".join(quote for quote in quotes if not any(quote in other for other in quotes if other != quote)))

    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=path.parent, prefix=".minutes-", suffix=".docx", delete=False) as tmp:
        temporary_path = Path(tmp.name)
    try:
        doc.save(temporary_path)
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)
