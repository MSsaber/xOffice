# !/usr/python
# -*- coding:utf-8 -*-

from setuptools import setup,find_packages

__version__ = '1.0.0'
requirements = open('requirements.txt').readlines()

setup(
    name = 'xlc',
    version = __version__,
    author = 'bxiao',
    author_email = '1752615737@qq.com',
    url = '',
    description = 'xlc : excel tool',
    packages = find_packages(exclude=["test"]),
    python_requires = '>=3.5.0',
    install_requires = requirements
)