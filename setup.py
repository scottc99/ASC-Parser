<<<<<<< HEAD
#### System and module requirements for adsorption analysis software ####
#!/usr/bin/env python

from setuptools import setup, find_packages
from distutils.command.install import install
import os
import sys


setup(
    name = 'ASC-Parser',
    version = "0.0.1.dev1",
    package_dir = {'': 'python'},
    packages = find_packages(),
    url = 'https://github.com/scottc99/ASC-Parser',
    long_description = open('README.md').read()

    # cmdclass = {'sdist': sdist_git},
)
=======
#### Setup and installation of apprpriate modules to use ASC-Parser ####
>>>>>>> 25880b362b4ed87ad98f232d8da74db7d7a86ae1
