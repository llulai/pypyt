import re
from functools import singledispatch
from pptx import Presentation
from pptx.chart.chart import Chart
from pptx.chart.data import ChartData
from pptx.shapes.base import BaseShape
from pptx.shapes.table import Table
from pptx.text.text import TextFrame


# BASE
def open_ppt(filename: str) -> Presentation:
    """Opens a pptx file given the filename and returns it"""
    return Presentation(filename)


def render_ppt(template_name: str, values: dict) -> Presentation:
    # opens the presentation
    prs = open_ppt(template_name)

    # checks each given item
    for k, v in values.items():

        # gets all the instances of the item in the presentation
        for shape in get_shapes_by_name(prs, k):

            # depending on what kind of item it is, it renders it
            if is_table(shape): render_table(v, shape.table)
            elif is_paragraph(shape): render_paragraph(v, shape.text_frame)
            elif is_chart(shape): render_chart(v, shape.chart)

    return prs


def save_ppt(prs: Presentation, filename: str):
    prs.save(filename)


def render_and_save_ppt(template_name: str, values: dict, filename: str):
    save_ppt(render_ppt(template_name, values), filename)


def get_shapes_by_name(prs: Presentation, name: str) -> list:
    return [shape for slide in prs.slides for shape in slide.shapes if shape.name == name]


def get_shapes(prs: Presentation) -> dict:
    """Returns a dictionary with the shape name as keys and shape type as values"""
    return {shape.name: get_shape_type(shape) for slide in prs.slides for shape in slide.shapes}


def get_shape_type(shape: BaseShape) -> str:
    if is_paragraph(shape): return 'paragraph'
    elif is_table(shape): return 'table'
    elif is_chart(shape): return 'chart'


def is_table(shape: BaseShape) -> bool:
    """Checks whether the given shape is has a table"""
    return shape.has_table


def is_paragraph(shape: BaseShape) -> bool:
    return shape.has_text_frame


def is_chart(shape: BaseShape) -> bool:
    return shape.has_chart


# RENDER TABLE
@singledispatch
def render_table(_, __):
    raise NotImplemented


@render_table.register(dict)
def _(values: dict, table: Table):
    """In the case you want to render placeholders within the table, it will call render_paragraph for each cell"""
    for row in table.rows:
        for cell in row.cells:
            render_paragraph(values, cell.text_frame)


@render_table.register(list)
def _(values: list, table: Table):
    """In the case you want to replace the whole table, it will set the value for each cell in the list"""
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            render_paragraph(values[i][j], cell.text_frame)


# RENDER PARAGRAPH
@singledispatch
def render_paragraph(values, text_frame: TextFrame):
    """In the case you want to replace the whole text"""
    paragraph = text_frame.paragraphs[0]
    p = paragraph._p
    for idx, run in enumerate(paragraph.runs):
        if idx == 0:
            continue
        p.remove(run._r)
    else:
        paragraph.text = 't'
    paragraph.runs[0].text = values


@render_paragraph.register(dict)
def _(values, text_frame: TextFrame):
    """In the case you want to replace placeholders within the paragraph"""
    for paragraph in text_frame.paragraphs:
        new_text_template = paragraph.text
        keywords = re.findall(r"\{(\w+)\}", new_text_template)
        if keywords:
            new_text = new_text_template.format(**{k: values[k] for k in keywords})
            p = paragraph._p
            for idx, run in enumerate(paragraph.runs):
                if idx == 0:
                    continue
                p.remove(run._r)
            paragraph.runs[0].text = new_text


# RENDER GRAPH
def render_chart(values, chart: Chart):
    chart_data = ChartData()

    chart_data.categories = values['categories']

    for label, series in values['data'].items():
        chart_data.add_series(label, series)

    if 'title' in values:
        chart.chart_title.text_frame.text = values['title']

    chart.replace_data(chart_data)
