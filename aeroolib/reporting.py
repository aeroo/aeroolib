##############################################################################
#
# Copyright (c) 2010 Alistek Ltd. (http://www.alistek.com) All Rights
# Reserved.
#                    General contacts <info@alistek.com>
#
# DISCLAIMER: This module is licensed under GPLv3 or newer and 
# is considered incompatible with OpenERP SA "AGPL + Private Use License"!
#
# Copyright (c) 2007, 2008 OpenHex SPRL. (http://openhex.com) All Rights
# Reserved.
#
# WARNING: This program as such is intended to be used by professional
# programmers who take the whole responsability of assessing all potential
# consequences resulting from its eventual inadequacies and bugs
# End users who are looking for a ready-to-use solution with commercial
# garantees and support are strongly adviced to contract a Free Software
# Service Company
#
# This program is Free Software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 3
# of the License, or (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
#
##############################################################################

__metaclass__ = type

import os, sys

import pkg_resources
from genshi.template import TemplateLoader

def _absolute(path):
    "Compute the absolute path of path relative to the caller file"
    if os.path.isabs(path):
        return path
    caller_fname = sys._getframe(2).f_globals['__file__']
    caller_dir = os.path.dirname(caller_fname)
    return os.path.abspath(os.path.join(caller_dir, path))

def _guess_type(mime):
    """
    Returns the codename used by aeroolib to identify which template plugin
    it should use to render a mimetype
    """
    mime = mime.lower()
    major, stype = mime.split('/', 1)
    if major == 'application':
        if 'opendocument' in stype:
            return 'oo.org'
        else:
            return stype
    elif major == 'text':
        if stype in ('xml', 'html', 'xhtml'):
            return 'markup'
        else:
            return 'text'


class MIMETemplateLoader(TemplateLoader):
    """This subclass of TemplateLoader use mimetypes to search and find
    templates to load.
    """

    factories = {}

    mime_func = [_guess_type]

    def get_type(self, mime):
        "finds the codename used by aeroolib to work on a mimetype"
        for func in reversed(self.mime_func):
            codename = func(mime)
            if codename is not None:
                return codename

    def load(self, path, mime=None, relative_to=None, cls=None):
        "returns a template object based on path"
        assert mime is not None or cls is not None

        if mime is not None:
            cls = self.factories[self.get_type(mime)]

        return super(MIMETemplateLoader, self).load(
            path, cls=cls, relative_to=relative_to)

    @classmethod
    def add_factory(cls, abbr_mimetype, template_factory, id_function=None):
        """adds a template factory to the already known factories"""
        cls.factories[abbr_mimetype] = template_factory
        if id_function is not None:
            cls.mime_func.append(id_function)

default_loader = MIMETemplateLoader(auto_reload=True)


class DefaultFactory:
    """This is the default factory used by aeroolib.

    It just returns a copy of the data it receives"""

    def __init__(self, *args, **kwargs):
        pass

    def __call__(self, **kwargs):
        data = kwargs.copy()
        return data

default_factory = DefaultFactory()


class Report:
    """Report is a simple interface on top of a rendering template.
    """

    def __init__(self, path, mimetype,
                 factory=default_factory, loader=default_loader):
        self.fpath = path
        self.mimetype = mimetype
        self.data_factory = factory
        self.tmpl_loader = loader
        self.filters = []

    def __call__(self, **kwargs):
        template = self.tmpl_loader.load(self.fpath, self.mimetype)
        data = self.data_factory(**kwargs)
        return template.generate(**data).filter(*self.filters)

    def __repr__(self):
        return '<aeroo report on %s>' % self.fpath



class ReportDict:

    def __init__(self, *args, **kwargs):
        self.mimetypes = {}
        self.ids = {}


class ReportRepository:
    """ReportRepository stores the report definition associated to objects.

    The report are indexed in this object by the object class they are working
    on and the name given to it by the user.
    """

    def __init__(self, datafactory=DefaultFactory):
        self.classes = {}
        self.default_factory = datafactory
        self.loader = default_loader

    def add_report(self, klass, mimetype, template_path, data_factory=None,
                   report_name='default', description=''):
        """adds a report to the repository.

        You will be able to find the report via
            - the class it is working on
            - the mimetype it outputs
            - the name of the report

        You also have the opportunity to define a specific data_factory.
        """
        if data_factory is None:
            data_factory = self.default_factory
        reports = self.classes.setdefault(klass, ReportDict())
        report = Report(_absolute(template_path), mimetype,
                        data_factory(klass, mimetype), self.loader)
        reports.ids[report_name] = report, mimetype, description
        reports.mimetypes.setdefault(mimetype, []) \
                         .append((report, report_name))

    def by_mime(self, klass, mimetype):
        """gets a list of report related to a class by specifying the mimetype
        """
        return self.classes[klass].mimetypes[mimetype]

    def by_id(self, klass, id):
        """get a report related to a class by its id
        """
        return self.classes[klass].ids[id]
