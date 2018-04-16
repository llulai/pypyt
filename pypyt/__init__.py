import re
from pptx import Presentation
from pptx.chart.data import ChartData


def replace_text(shape, new_text):
    paragraph = shape.text_frame.paragraphs[0]
    p = paragraph._p
    for idx, run in enumerate(paragraph.runs):
        if idx == 0:
            continue
        p.remove(run._r)
    paragraph.runs[0].text = new_text


def replace_data(shape, data):
    chart_data = ChartData()

    chart_data.categories = data['categories']

    for label, serie in data['data'].items():
        chart_data.add_series(label, serie)

    if 'title' in data:
        shape.chart.chart_title.text_frame.text = data['title']

    shape.chart.replace_data(chart_data)


def render_text(shape, **kwargs):
    for paragraph in shape.text_frame.paragraphs:
        new_text_template = paragraph.text
        keywords = re.findall(r"\{(\w+)\}", new_text_template)
        if keywords:
            new_text = new_text_template.format(**{k: kwargs[k] for k in keywords})
            p = paragraph._p
            for idx, run in enumerate(paragraph.runs):
                if idx == 0:
                    continue
                p.remove(run._r)
            paragraph.runs[0].text = new_text


def render_table(self, shape, **kwargs):
    for row in shape.table.rows:
        for cell in row.cells:
            self._render_text(cell, **kwargs)


def get_shape_by_name(prs, name):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.name == name:
                return shape


def render_ppt(template_name: str, values: dict) -> Presentation:
    prs = Presentation(template_name)
    for k, v in values.items():
        shape = get_shape_by_name(prs, k)
        if shape.has_table:
            render_table(shape, **v)
        elif shape.has_text_frame:
            if isinstance(v, dict):
                render_text(shape, **v)
            elif isinstance(v, str):
                replace_text(shape, v)
        elif shape.has_chart:
            replace_data(shape, v)
    return prs


def save_ppt(prs: Presentation, filename: str):
    prs.save(filename)


def render_and_save_ppt(template_name: str, values: dict, filename: str):
    save_ppt(render_ppt(template_name, values), filename)
