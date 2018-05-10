# pypyt 0.0.5
It is a simple library to render PowerPoint presentations using python

# How to install it:


    pip install pypyt



# How to use it:
- You need to create a template file with the objects named as shown in [this video](https://www.youtube.com/watch?v=IhES3of_9Nw).

Lets assume that you have a template file named \_\_template\_\_.pptx with two shapes: *presentation_title* and
*client_name* (as shown in the image below), in order to render it you might use the following code:

![](images/template1.png)
```python
from pypyt import render_and_save_ppt

values = {
    'presentation_title': "This is a cool presentation",
    'client_name': "Cool Client"
}

render_and_save_ppt('__template__.pptx', values, 'rendered_ppt.pptx')
```
    
This will render a presentation like the one below.
![](images/output1.png)


# Reference

```python
from pypyt import pypyt_doc
pypyt_doc()
```

# Contact
If you're trying to use this and want to extend it, have any request or need help, just ping me on slack (j.gajardo) or
send me an email to j.gajardo@criteo.com