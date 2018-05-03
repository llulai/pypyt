.. pypyt documentation master file, created by
   sphinx-quickstart on Thu May  3 16:49:34 2018.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to pypyt's documentation!
=================================
Pypyt is a library to render PowerPoint presentations from python in an easy and intuitive way.

How to Install it
-----------------
::

   pip install pypyt


How to use it
-------------

::

   from pypyt import *
   values = {'presentation_title': "My Cool Presentation"}
   render_and_save_ppt('__template__.pptx', values, 'presentation.pptx')


For more information about how to use this library, see the :ref:`api`.


.. toctree::
   :maxdepth: 2
   :caption: Contents:

   api


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
