from setuptools import setup


setup(
    name='pypyt',
    version="0.0.5",
    author='Julio Gajardo',
    author_email='j.gajardo@criteo.com',
    description='Simple library to render ppt templates in python',
    license='MIT',
    url='https://gitlab.criteois.com/j.gajardo/pypyt',
    packages=['pypyt', 'html_docs'],
    install_requires=['python-pptx'],
)

__author__ = 'Julio Gajardo'
