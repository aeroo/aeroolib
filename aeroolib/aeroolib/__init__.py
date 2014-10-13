"""
Aeroo Library
=========

A templating library which provides a way to easily output all kind of
different files (odt, ods). Adding support for more filetype is
easy: you just have to create a plugin for this.

Is based on Relatorio project

aeroolib also provides a report repository allowing you to link python objects
and report together, find reports by mimetypes/name/python objects.
"""
from aeroolib.reporting import MIMETemplateLoader, ReportRepository, Report
import plugins

__version__ = '1.2.0'
