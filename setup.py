from setuptools import setup


setup(
    name='pypyt',
    version="0.0.20",
    author='Julio Gajardo',
    author_email='j.gajardo@criteo.com',
    description='Simple library to render ppt templates in python',
    license='MIT',
    url='https://gitlab.criteois.com/j.gajardo/pypyt',
    packages=['pypyt'],
    install_requires=['python-pptx==0.6.15', 'pandas==0.23.4', 'numpy==1.14.2']
)

__author__ = 'Julio Gajardo'
