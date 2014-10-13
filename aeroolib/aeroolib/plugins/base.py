##############################################################################
#
# Copyright (c) 2012 Alistek Ltd (http://www.alistek.com) All Rights
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
#
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

from cStringIO import OutputType

import genshi.core
from genshi.template import NewTextTemplate, MarkupTemplate

from aeroolib.reporting import Report, MIMETemplateLoader


class AerooStream(genshi.core.Stream):
    "Base class for the aeroo streams."

    def render(self, method=None, encoding='utf-8', out=None, **kwargs):
        "calls the serializer to render the template"
        return self.serializer(self.events)

    def serialize(self, method='xml', **kwargs):
        "generates the bitstream corresponding to the template"
        return self.render(method, **kwargs)

    def __or__(self, function):
        "Support for the bitwise operator"
        return AerooStream(self.events | function, self.serializer)

    def __str__(self):
        val = self.render()
        if isinstance(val, OutputType):
            return val.getvalue()
        else:
            return val

MIMETemplateLoader.add_factory('text', NewTextTemplate)
MIMETemplateLoader.add_factory('xml', MarkupTemplate)
