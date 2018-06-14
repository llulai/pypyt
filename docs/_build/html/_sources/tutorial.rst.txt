.. _tutorial:

Tutorial
--------

First of all, you need to create a template file with the objects names as shown in
`this video <https://www.youtube.com/watch?v=IhES3of_9Nw>`_

Lets assume that you have a template file named ``__template__.pptx`` with two shapes: ``presentation_title`` and
``client_name`` as shown in the image below.

.. image:: images/template1.png

In order to render it you might use the following code::

   from pypyt import render_and_save_template

       values = {
           'presentation_title': "This is a cool presentation",
           'client_name': "Cool Client"
       }

       render_and_save_template('__template__.pptx', values, 'rendered_ppt.pptx')


This will render a presentation like the one below.

.. image:: images/output1.png


For more information about how to use this library, see the :ref:`use`.

For the full documentation, see the :ref:`api`.