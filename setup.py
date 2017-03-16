from setuptools import setup, find_packages
from setuptools.command.test import test as TestCommand
import sys

setup(
    name='sse_invoice_entry',
    version='1.0',
    packages=find_packages(),
    entry_points={'console_scripts': ['sse-invoice-entry = sse_invoice_entry:main']},
    install_requires=['pywin32'],
    license='',
    author='Mike Pagel',
    author_email='mike@mpagel.de',
    description='Invoice data entry in SSE.'
)
