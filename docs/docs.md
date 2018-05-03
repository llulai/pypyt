<h1 id="pypyt">pypyt</h1>

Renders PowerPoint presentations easily with Python
<h2 id="pypyt.open_ppt">open_ppt</h2>

```python
open_ppt(filename:str) -> <function Presentation at 0x000001A85D04CD08>
```

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

<h2 id="pypyt.render_ppt">render_ppt</h2>

```python
render_ppt(prs:<function Presentation at 0x000001A85D04CD08>, values:dict) -> <function Presentation at 0x000001A85D04CD08>
```

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

<h2 id="pypyt.render_template">render_template</h2>

```python
render_template(template_name:str, values:dict) -> <function Presentation at 0x000001A85D04CD08>
```

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

<h2 id="pypyt.save_ppt">save_ppt</h2>

```python
save_ppt(prs:<function Presentation at 0x000001A85D04CD08>, filename:str) -> None
```

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

<h2 id="pypyt.render_and_save_ppt">render_and_save_ppt</h2>

```python
render_and_save_ppt(template_name:str, values:dict, filename:str) -> None
```

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

<h2 id="pypyt.get_shapes_by_name">get_shapes_by_name</h2>

```python
get_shapes_by_name(prs:<function Presentation at 0x000001A85D04CD08>, name:str) -> list
```

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


<h2 id="pypyt.get_shapes">get_shapes</h2>

```python
get_shapes(prs:<function Presentation at 0x000001A85D04CD08>) -> tuple
```

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

<h2 id="pypyt.get_shape_type">get_shape_type</h2>

```python
get_shape_type(shape:pptx.shapes.base.BaseShape) -> str
```

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

<h2 id="pypyt.is_table">is_table</h2>

```python
is_table(shape:pptx.shapes.base.BaseShape) -> bool
```
Checks whether the given shape is has a table
<h2 id="pypyt.is_paragraph">is_paragraph</h2>

```python
is_paragraph(shape:pptx.shapes.base.BaseShape) -> bool
```
Checks whether the given shape is has a paragraph
<h2 id="pypyt.is_chart">is_chart</h2>

```python
is_chart(shape:pptx.shapes.base.BaseShape) -> bool
```
Checks whether the given shape is has a chart
<h2 id="pypyt.render_table">render_table</h2>

```python
render_table(values:Union[dict, list, pandas.core.frame.DataFrame], table:pptx.shapes.table.Table) -> None
```

Renders a table with the given values.

Parameters
----------
values: Union[dict, list, DataFrame]
    Values to render the table

table: Table
    Table object to be rendered.

<h2 id="pypyt.render_paragraph">render_paragraph</h2>

```python
render_paragraph(values, text_frame:pptx.text.text.TextFrame) -> None
```

In the case you want to replace the whole text.

Parameters
----------
values: Union[dict, str, int, float]
    Values to render the text.

text_frame: TextFrame
    TextFrame object to be rendered.

<h2 id="pypyt.render_chart">render_chart</h2>

```python
render_chart(values:Union[dict, pandas.core.frame.DataFrame], chart:pptx.chart.chart.Chart) -> None
```

Renders the given values into the given chart.

Parameters
----------
values: Union[dict, DataFrame]
    Values to render the chart.

chart: Chart
    Chart object to be rendered.

