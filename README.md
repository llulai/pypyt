# pypyt
pypyt (pronounced **pie pie tee**, not pee wide pee wide tee), is a simple library to render PowerPoint presentations using python


version |coverage 
--- | ---
[![PyPI version](https://badge.fury.io/py/pypyt.svg)](https://badge.fury.io/py/pypyt) | [![Coverage Status](https://coveralls.io/repos/github/llulai/pypyt/badge.svg?branch=master)](https://coveralls.io/github/llulai/pypyt?branch=master)
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

The documentation is available <a href="https://pypyt.readthedocs.io/en/latest/index.html">online</a> and in
[pdf](https://media.readthedocs.org/pdf/pypyt/latest/pypyt.pdf) through read the docs.

Also, you can open it with the following command.

```python
from pypyt import pypyt_doc
pypyt_doc()
```

# Contact
If you're trying to use this and want to extend it, have any request or need help, just ping me on slack (j.gajardo) or
send me an email to j.gajardo@criteo.com

# Acknowledgments
- **(Criteo - New York) Sebastian Riera:** For finding bugs, suggest fixes and new features.
- **(Criteo - New York) Mehdi Rifai:** For suggesting fixes to bugs.
- **(Criteo - Brazil) Caio Camilli:** For being the first user of the package and bring up possible points of improvement.

# License
The contents of this repository are covered under the [MIT License](LICENSE.txt).