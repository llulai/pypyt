"""Renders PowerPoint presentations easily with Python"""

import re
from functools import singledispatch
from pptx import Presentation
from pptx.chart.chart import Chart
from pptx.chart.data import CategoryChartData
from pptx.shapes.base import BaseShape
from pptx.shapes.table import Table
from pptx.text.text import TextFrame
from pandas import DataFrame
from typing import Union


# BASE
def open_ppt(filename: str) -> Presentation:
    """
    Opens a pptx file given the filename and returns it.

    Parameters
    ----------
    filename: str
        The name of the file to be open.

    Returns
    -------
    pptx.presentation.Presentation
        Presentation object

    Examples
    --------
    >>> open_ppt('template.pptx')
    <pptx.presentation.Presentation at ...>
    """
    return Presentation(filename)


def render_ppt(prs: Presentation, values: dict) -> Presentation:
    """
    Returns a rendered presentation given the template name and values to be rendered.

    Parameters
    ----------
    prs: pptx.presentation.Presentation
        Presentation to be rendered.

    values: dict
        Dictionary with the values to render on the template.

    Returns
    -------
    pptx.presentation.Presentation
        Rendered presentation

    Examples
    --------
    >>> prs = open_ppt('template.pptx')
    >>> values = {'presentation_title': "My Cool Presentation"}
    >>> render_ppt(prs, values)
    <pptx.presentation.Presentation at ...>
    """

    # checks each given item
    for key, value in values.items():

        # gets all the instances of the item in the presentation
        for shape in get_shapes_by_name(prs, key):

            # depending on what kind of item it is, it renders it
            if is_table(shape):
                render_table(value, shape.table)
            elif is_paragraph(shape):
                render_paragraph(value, shape.text_frame)
            elif is_chart(shape):
                render_chart(value, shape.chart)

    return prs


def render_template(template_name: str, values: dict) -> Presentation:
    """
    Returns a rendered presentation given the template name and values to be rendered.

    Parameters
    ----------
    template_name: str
        Name of the presentation to be rendered.

    values: dict
        Dictionary with the values to render on the template.

    Returns
    -------
    pptx.presentation.Presentation
        Rendered presentation

    Examples
    --------
    >>> values = {'presentation_title': "My Cool Presentation"}
    >>> render_template('template.pptx', values)
    <pptx.presentation.Presentation at ...>
    """
    return render_ppt(open_ppt(template_name), values)


def save_ppt(prs: Presentation, filename: str) -> None:
    """
    Saves the given presentation with the given filename.

    Parameters
    ----------
    prs: pptx.presentation.Presentation
        Presentation to be saved.

    filename: str
        Name of the file to be saved.

    Examples
    --------
    >>> values = {'presentation_title': "My Cool Presentation"}
    >>> rendered_prs = render_template('template.pptx', values)
    >>> save_ppt(rendered_prs, 'presentation.pptx')
    """
    prs.save(filename)


def render_and_save_ppt(template_name: str, values: dict, filename: str) -> None:
    """
    Renders and save a presentation with the given template name,
    values to be rendered, and filename.
    Parameters
    ----------
    template_name: str
        Name of the presentation to be saved.

    values: dict
        Dictionary with the values to render on the template.

    filename: str
        Name of the file to be saved.

    Examples
    --------
    >>> values = {'presentation_title': "My Cool Presentation"}
    >>> render_and_save_ppt('template.pptx', values, 'presentation.pptx')
    """
    save_ppt(render_template(template_name, values), filename)


def get_shapes_by_name(prs: Presentation, name: str) -> list:
    """
    Returns a list of shapes with the given name in the given presentation.
    Parameters
    ----------
    prs: pptx.presentation.Presentation
        Presentation to be saved.

    name: str
        Name of the shape(s) to be returned.

    Examples
    --------
    >>> prs = open_ppt('template.pptx')
    >>> get_shapes_by_name(prs, 'chart')
    [<pptx.shapes.placeholder.PlaceholderGraphicFrame at ...>]

    """
    return [shape for slide in prs.slides for shape in slide.shapes if shape.name == name]


def get_shapes(prs: Presentation) -> tuple:
    """
    Returns a tuple of tuples with the shape name and shape type.

    Parameters
    ----------
    prs: pptx.presentation.Presentation
        Presentation to get the shapes from.

    Returns
    -------
    tuple:
        Tuple with the shapes.

    Examples
    --------
    >>> prs = open_ppt('template.pptx')
    >>> get_shapes(prs)
    (('client_name', 'paragraph'),
     ('presentation_title', 'paragraph'),
     ('slide_text', 'paragraph'),
     ('slide_title', 'paragraph'),
     ('chart', 'chart'),
     ('Title 1', 'paragraph'),
     ('table', 'table'),
     ('Title 1', 'paragraph'))
    """
    return tuple((shape.name, get_shape_type(shape)) for slide in prs.slides for shape in slide.shapes)


def get_shape_type(shape: BaseShape) -> str:
    """
    Returns a string with the kind of the given shape.

    Parameters
    ----------
    shape: BaseShape
        Shape to get the type from.

    Returns
    -------
    string:
        String representing the type of the shape

    Examples
    --------
    >>> prs = open_ppt('template.pptx')
    >>> shapes = get_shapes_by_name(prs, 'client_name')
    >>> get_shape_type(shapes[0])
    'paragraph'
    """
    if is_paragraph(shape):
        return 'paragraph'
    elif is_table(shape):
        return 'table'
    elif is_chart(shape):
        return 'chart'
    return ''


def is_table(shape: BaseShape) -> bool:
    """Checks whether the given shape is has a table"""
    return shape.has_table


def is_paragraph(shape: BaseShape) -> bool:
    """Checks whether the given shape is has a paragraph"""
    return shape.has_text_frame


def is_chart(shape: BaseShape) -> bool:
    """Checks whether the given shape is has a chart"""
    return shape.has_chart


# RENDER TABLE
@singledispatch
def render_table(values: Union[dict, list, DataFrame], table: Table) -> None:  # pylint: disable=unused-argument
    """
    Renders a table with the given values.

    Parameters
    ----------
    values: Union[dict, list, DataFrame]
        Values to render the table

    table: Table
        Table object to be rendered.
    """
    raise NotImplementedError


@render_table.register(dict)
def _(values: dict, table: Table) -> None:
    """In the case you want to render placeholders within the table,
    it will call render_paragraph for each cell"""
    for row in table.rows:
        for cell in row.cells:
            render_paragraph(values, cell.text_frame)


@render_table.register(DataFrame)
def _(values: DataFrame, table: Table) -> None:
    # TODO: raise error if size(values) != size(table)
    table_rows = iter(table.rows)

    if hasattr(values, 'header') and values.header:
        for values_cell, table_cell in zip(list(values), next(table_rows)):
            render_paragraph(values_cell, table_cell.text_frame)

    for values_row, table_row in zip(list(values.values), table_rows):
        for values_cell, table_cell in zip(list(values_row), table_row.cells):
            render_paragraph(values_cell, table_cell.text_frame)


@render_table.register(list)
def _(values: list, table: Table) -> None:
    """In the case you want to replace the whole table,
    it will set the value for each cell in the list"""
    # TODO: raise error if size(values) != size(table)
    for values_row, table_row in zip(values, table.rows):
        for values_cell, table_cell in zip(values_row, table_row.cells):
            render_paragraph(values_cell, table_cell.text_frame)


# RENDER PARAGRAPH
@singledispatch
def render_paragraph(values, text_frame: TextFrame) -> None:
    """
    In the case you want to replace the whole text.

    Parameters
    ----------
    values: Union[dict, str, int, float]
        Values to render the text.

    text_frame: TextFrame
        TextFrame object to be rendered.
    """
    paragraph = text_frame.paragraphs[0]
    p = paragraph._p  # pylint: disable=protected-access,invalid-name
    for idx, run in enumerate(paragraph.runs):
        if idx == 0:
            continue
        p.remove(run._r)  # pylint: disable=protected-access
    else:
        paragraph.text = 't'
    paragraph.runs[0].text = values


@render_paragraph.register(dict)
def _(values: dict, text_frame: TextFrame) -> None:
    """In the case you want to replace placeholders within the paragraph"""
    for paragraph in text_frame.paragraphs:
        new_text_template = paragraph.text
        keywords = re.findall(r"\{(\w+)\}", new_text_template)
        if keywords:
            new_text = new_text_template.format(**{k: values[k] for k in keywords})
            p = paragraph._p  # pylint: disable=protected-access,invalid-name
            for idx, run in enumerate(paragraph.runs):
                if idx == 0:
                    continue
                p.remove(run._r)  # pylint: disable=protected-access
            paragraph.runs[0].text = new_text


# RENDER GRAPH
@singledispatch
def render_chart(values: Union[dict, DataFrame], chart: Chart) -> None:  # pylint: disable=unused-argument
    """
    Renders the given values into the given chart.

    Parameters
    ----------
    values: Union[dict, DataFrame]
        Values to render the chart.

    chart: Chart
        Chart object to be rendered.
    """
    raise NotImplementedError(f"Method not implemented for {type(values)} object type")


@render_chart.register(dict)
def _(values: DataFrame, chart: Chart) -> None:
    """Renders into the given chart the values in the DataFrame."""

    chart_data = CategoryChartData()

    chart_data.categories = values['categories']

    for label, series in values['data'].items():
        chart_data.add_series(label, series)

    if 'title' in values:
        chart.chart_title.text_frame.text = values['title']

    chart.replace_data(chart_data)


@render_chart.register(DataFrame)
def _(values: DataFrame, chart: Chart) -> None:
    """Renders into the given chart the values in the DataFrame."""

    chart_data = CategoryChartData()

    chart_data.categories = list(values.index)

    for label in list(values):
        chart_data.add_series(label, list(values[label].values))

    if hasattr(values, 'title'):
        chart.chart_title.text_frame.text = values.title

    chart.replace_data(chart_data)
