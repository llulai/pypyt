.. _api:

API Documentation
=================

.. module:: pypyt

For normal use, all you will have to do to get started is::

    from pypyt import *

This will import the following:

* The :func:`render_and_save_ppt` function to render and save a presentation.


Rendering a presentation
------------------------

.. autofunction:: render_and_save_ppt()

Call this function with the template_name, values and file_name and will render and save a presentation.::

    values = {'presentation_title': "My Cool Presentation"}
    render_and_save_ppt('__template__.pptx', values, 'presentation.pptx')

Rendering Paragraphs
--------------------

todo

Rendering Charts
----------------

todo

Rendering Tables
----------------

todo