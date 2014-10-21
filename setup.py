import os
import re
from setuptools import setup, find_packages

def get_version():
    init = open(os.path.join(os.path.dirname(__file__), 'aeroolib',
                             '__init__.py')).read()
    return re.search(r"""__version__ = '([0-9.\sA-Za-z]*)'""", init).group(1)

setup(
    name="aeroolib",
    url="http://www.alistek.com/",
    author="Alistek Ltd",
    author_email="info@alistek.com",
    maintainer=u"Alistek Ltd",
    maintainer_email="info@alistek.com",
    description="A templating library able to output ODF files",
    long_description="""
aeroolib
=========

A templating library which provides a way to easily output all kind of
different files (odt, ods). Adding support for more filetype is easy:
you just have to create a plugin for this.

aeroolib also provides a report repository allowing you to link python objects
and report together, find reports by mimetypes/name/python objects.
    """,
    license="GPL License",
    version=get_version(),
    packages=find_packages(),
    install_requires=[
        "Genshi >= 0.5",
        "lxml >= 2.0"
    ],
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: GNU General Public License (GPL)",
        "Operating System :: OS Independent",
        "Programming Language :: Python",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Text Processing",
    ],
    test_suite="nose.collector",
    )
