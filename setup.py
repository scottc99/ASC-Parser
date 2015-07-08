#### Setup and installation for apprpriate modules to use ASC-Parser ####
#!/usr/bin/env python2.7

from setuptools import setup
from distutils.command.install import install
from codecs import open 

setup(
    name = 'ASC_Parser',
    py_modules = ['os', 'glob', 'distutils', 'distutils.core', 'collections', 'Tkinter', 'pprint', 'codecs'],
    package_dir = {'': 'site-packages'},
    packages = ['xlrd', 'xlwt', 'tkFileDialog', 'dicttoxml', 'lxml']
)
