.. _api:

Library Documentation
=====================

.. module:: pypyt

For normal use, all you will have to do to get started is::

    from pypyt import *

This will import the following functions:


Template Functions

* The :func:`open_template` function opens a template.
* The :func:`render_template` function to renders a template.
* The :func:`render_and_save_template` function renders and saves a template.


Presentation Functions:

* The :func:`get_shapes` function gets all the shapes and its type in a presentation.
* The :func:`get_shapes_by_name` function gets all the shapes with the given name in a presentation.
* The :func:`render_ppt` function renders a presentation.
* The :func:`render_and_save_ppt` function renders and save a presentation.
* The :func:`save_ppt` function saves a presentation.


Shape Functions:

* The :func:`get_shape_type` function gets the type of a shape.
* The :func:`is_chart` function returns ``True`` if the given shape is a chart.
* The :func:`is_hyperlink` function returns ``True`` if the given shape is a hyperlink.
* The :func:`is_paragraph` function returns ``True`` if the given shape is a paragraph.
* The :func:`is_picture` function returns ``True`` if the given shape is a picture.
* The :func:`is_table` function returns ``True`` if the given shape is a table.

Render Functions:

* The :func:`render_chart` function renders a chart object.
* The :func:`render_hyperlink` function renders a hyperlink object.
* The :func:`render_paragraph` function renders a paragraph object.
* The :func:`render_picture` function renders a picture object.
* The :func:`render_table` function renders a table object.

Utility Functions:

* The :func:`picture_from_url` function gets an image from the given url.


Documentation Functions:

* The :func:`pypyt_doc` function opens this documentation.


Template Functions
------------------

.. autofunction:: pypyt.open_template(template_name)
.. autofunction:: render_template(template_name, values)
.. autofunction:: render_and_save_template(template_name, values, filename)


Presentation Functions
----------------------

.. autofunction:: get_shapes(prs)
.. autofunction:: get_shapes_by_name(prs, name)
.. autofunction:: render_ppt(prs, values)
.. autofunction:: render_and_save_ppt(template_name, values, filename)
.. autofunction:: save_ppt(prs, filename)


Shape Functions
---------------

.. autofunction:: get_shape_type(shape)
.. autofunction:: is_chart(shape)
.. autofunction:: is_hyperlink(shape)
.. autofunction:: is_paragraph(shape)
.. autofunction:: is_picture(shape)
.. autofunction:: is_table(shape)



Render Functions
----------------

.. autofunction:: render_chart(values, chart)
.. autofunction:: render_hyperlink(values, hyperlink)
.. autofunction:: render_paragraph(values, text_frame)
.. autofunction:: render_picture(values, table)
.. autofunction:: render_table(values, table)


Utility Functions
-----------------

.. autofunction:: picture_from_url(url)

Documentation Functions
-----------------------

.. autofunction:: pypyt_doc()

