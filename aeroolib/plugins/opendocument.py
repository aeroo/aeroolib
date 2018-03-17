##############################################################################
#
# Copyright (c) 2012 Alistek Ltd (http://www.alistek.com) All Rights
# Reserved.
#                    General contacts <info@alistek.com>
#
# DISCLAIMER: This module is licensed under GPLv3 or newer and 
# is considered incompatible with OpenERP SA "AGPL + Private Use License"!
#
# Copyright (c) 2009 Cedric Krier.
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

import re
from hashlib import md5

import time
import urllib
import zipfile
from io import StringIO, BytesIO
from copy import copy, deepcopy


import warnings
warnings.filterwarnings('always', module='aeroolib.plugins.opendocument')

import lxml.etree
import genshi
import genshi.output
from genshi.template import MarkupTemplate
from genshi.filters import Transformer
from genshi.filters.transform import ENTER, EXIT
from genshi.core import Stream
from genshi.template.interpolation import PREFIX


from aeroolib.plugins.base import AerooStream
from aeroolib.reporting import Report, MIMETemplateLoader

GENSHI_EXPR = re.compile(r'''
        (/)?                                 # is this a closing tag?
        (for|if|choose|when|otherwise|with|def|match|attrs|content|replace|strip)  # tag directive
        \s*
        (?:\s(\w+)=["'](.*)["']|$)           # match a single attr & its value
        |
        .*                                   # or anything else
        ''', re.VERBOSE)

EXTENSIONS = {'image/png': 'png',
              'image/jpeg': 'jpg',
              'image/bmp': 'bmp',
              'image/gif': 'gif',
              'image/tiff': 'tif',
              'image/xbm': 'xbm',
             }

AEROO_URI = 'http://alistek.com/'
GENSHI_URI = 'http://genshi.edgewall.org/'
MANIFEST = 'META-INF/manifest.xml'
META = 'meta.xml'
STYLES = 'styles.xml'
output_encode = genshi.output.encode
EtreeElement = lxml.etree.Element

TAB = "[%s]" % md5(b"&#x9;").hexdigest()
NEW_LINE = "[%s]" % md5(b"&#xA;").hexdigest()

# A note regarding OpenDocument namespaces:
#
# The current code assumes the original OpenOffice document uses default
# namespace prefix ("table", "xlink", "draw", ...). We derive the actual
# namespaces URIs from their prefix, instead of the other way round. This has
# the advantage that if a new version of the format use different namespaces
# (this is not the case for ODF 1.1 but could be the case in the future since
# there is a version number in those namespaces after all), Aeroo will
# support those new formats out of the box.


# A note about attribute namespaces:
#
# Ideally, we should update the namespace map of all the nodes we add
# (Genshi) attributes to, so that the attributes use a nice "py" prefix instead
# of a generated one (eg. "ns0", which is correct but ugly) in the case no
# parent node defines it. Unfortunately, lxml doesn't support this:
# the nsmap attribute of Element objects is (currently) readonly.


class OOTemplateError(genshi.template.base.TemplateSyntaxError):
    "Error to raise when there is a SyntaxError in the genshi template"

class ImageHref:
    "A class used to add images in the odf zipfile"

    def __init__(self, namespaces, zfile, manifest, context):
        self.namespaces = namespaces
        self.zip = zfile
        self.manifest = manifest
        self.context = context.copy()

    def __call__(self, expr):
        bitstream, mimetype = expr[:2]
        if isinstance(bitstream, Report):
            bitstream = bitstream(**self.context).render()
        bitstream.seek(0)
        file_content = bitstream.read()
        if len(file_content)>0:
            name = md5(file_content).hexdigest()
            path = 'Pictures/%s.%s' % (name, EXTENSIONS[mimetype])
            if path not in self.zip.namelist():
                self.zip.writestr(path, file_content)
                self.manifest.add_file_entry(path, mimetype)

            if len(expr) == 4:
                width, height = expr[2:]
            else:
                width = height = False
            attrs = {'{%s}href' % self.namespaces['xlink']: path}
            if width:
                attrs['{%s}width' % self.namespaces['svg']] = width
            if height:
                attrs['{%s}height' % self.namespaces['svg']] = height

            return attrs
        else:
            return {}

class ColumnCounter:
    """A class used to count the actual maximum number of cells (and thus
    columns) a table contains accross its rows.
    """
    def __init__(self):
        self.temp_counters = {}
        self.counters = {}

    def reset(self, loop_id):
        self.temp_counters[loop_id] = 0

    def inc(self, loop_id):
        self.temp_counters[loop_id] += 1

    def store(self, loop_id, table_name):
        self.counters[table_name] = max(self.temp_counters.pop(loop_id),
                                        self.counters.get(table_name, 0))


def wrap_nodes_between(first, last, new_parent):
    """An helper function to move all nodes between two nodes to a new node
    and add that new node to their former parent. The boundary nodes are
    removed in the process.
    """
    old_parent = first.getparent()

    # Any text after the opening tag (and not within a tag) need to be handled
    # explicitly. For example in <if>xxx<span>yyy</span>zzz</if>, zzz is
    # copied along the span tag, but not xxx, which corresponds to the tail
    # attribute of the opening tag.
    if first.tail:
        new_parent.text = first.tail
    for node in first.itersiblings():
        if node is last:
            break
        # appending a node to a new parent also
        # remove it from its previous parent
        new_parent.append(node)
    old_parent.replace(first, new_parent)
    new_parent.tail = last.tail
    old_parent.remove(last)


def update_py_attrs(node, value):
    """An helper function to update py_attrs of a node.
    """
    if not value:
        return
    py_attrs_attr = '{%s}attrs' % GENSHI_URI
    if not py_attrs_attr in node.attrib:
        node.attrib[py_attrs_attr] = value
    else:
        node.attrib[py_attrs_attr] = \
                "(lambda x, y: x.update(y) or x)(%s or {}, %s or {})" % \
                (node.attrib[py_attrs_attr], value)

def _filter(val):
    if type(val)==bool:
        return ''
    elif isinstance(val, str):
        return val.replace('\t', TAB).replace('\r\n', NEW_LINE).replace('\n', NEW_LINE)
    else:
        return val

def _hyperlink(namespaces):
    def _hyperlink(expr):
        attrs = {'{%s}href' % namespaces['xlink']:expr}
        return attrs
    return _hyperlink

class Template(MarkupTemplate):

    def __init__(self, source, serializer, filepath=None, styles=None,
                filename=None, loader=None, encoding=None, lookup='strict',
                allow_exec=True):
        self.namespaces = {}
        self.inner_docs = []
        self.has_col_loop = False
        self.Serializer = serializer #OOSerializer(filepath or source, oo_styles=styles)
        super(Template, self).__init__(source, filepath, filename, loader,
                                       encoding, lookup, allow_exec)

    def _parse(self, source, encoding):
        """parses the odf file.

        It adds genshi directives and finds the inner docs.
        """
        zf = self.Serializer.inzip

        content = self.Serializer.content_xml
        styles = self.Serializer.styles_xml

        template = super(Template, self)
        content = template._parse(self.insert_directives(content), encoding)
        styles = template._parse(self.insert_directives(styles), encoding)
        content_files = [('content.xml', content)]
        styles_files = [('styles.xml', styles)]

        while self.inner_docs:
            doc = self.inner_docs.pop()
            c_path, s_path = doc + '/content.xml', doc + '/styles.xml'
            content = zf.read(c_path)
            styles = zf.read(s_path)
            c_parsed = template._parse(self.insert_directives(content),
                                       encoding)
            s_parsed = template._parse(self.insert_directives(styles),
                                       encoding)
            content_files.append((c_path, c_parsed))
            styles_files.append((s_path, s_parsed))

        parsed = []
        for fpath, fparsed in content_files + styles_files:
            parsed.append((genshi.core.PI, ('aeroo', fpath), None))
            parsed += fparsed

        return parsed

    def insert_directives(self, content):
        """adds the genshi directives, handle the images and the innerdocs.
        """
        tree = lxml.etree.parse(BytesIO(content))
        root = tree.getroot()

        # assign default/fake namespaces so that documents do not need to
        # define them if they don't use them
        self.namespaces = {
            "text": "urn:text",
            "draw": "urn:draw",
            "table": "urn:table",
            "office": "urn:office",
            "xlink": "urn:xlink",
            "svg": "urn:svg",
        }
        # but override them with the real namespaces
        self.namespaces.update(root.nsmap)

        # remove any "root" namespace as lxml.xpath do not support them
        self.namespaces.pop(None, None)

        self.namespaces['py'] = GENSHI_URI
        self.namespaces['aeroo'] = AEROO_URI

        self._invert_style(tree)
        self._handle_aeroo_tags(tree)
        self._handle_images(tree)
        self._handle_hyperlinks(tree)
        self._handle_innerdocs(tree)
        self._escape_values(tree)
        return BytesIO(lxml.etree.tostring(tree))

    def _invert_style(self, tree):
        "inverts the text:a and text:span"
        xpath_expr = "//text:a[starts-with(@xlink:href, 'python://')]" \
                     "/text:span"
        for span in tree.xpath(xpath_expr, namespaces=self.namespaces):
            text_a = span.getparent()
            outer = text_a.getparent()
            text_a.text = span.text
            span.text = ''
            text_a.remove(span)
            outer.replace(text_a, span)
            span.append(text_a)

    def _aeroo_statements(self, tree):
        def check_except_directive( expr):
            if GENSHI_EXPR.match(expr).groups()[1] is not None:
                return False
            return True
        "finds the aeroo statements (text:a/text:placeholder)"
        # If this node href matches the aeroo URL it is kept.
        # If this node href matches a genshi directive it is kept for further
        # processing.
        xlink_href_attrib = '{%s}href' % self.namespaces['xlink']
        text_a = '{%s}a' % self.namespaces['text']
        placeholder = '{%s}placeholder' % self.namespaces['text']
        text_input = '{%s}text-input' % self.namespaces['text']
        s_xpath = "//text:a[starts-with(@xlink:href, 'python://')]" \
                  "| //text:placeholder" \
                  "| //text:text-input[starts-with(@text:description, '<') and substring(@text:description, string-length(@text:description))='>']"

        r_statements = []
        opened_tags = []
        # We map each opening tag with its closing tag
        closing_tags = {}
        dir_sequence = [] # its need for follow syntax checking in template
        for statement in tree.xpath(s_xpath, namespaces=self.namespaces):
            if statement.tag == placeholder:
                if not statement.text:
                    continue
                expr = statement.text[1:-1]
                expr = check_except_directive(expr) and "__filter(%s)" % statement.text[1:-1] or expr
            elif statement.tag == text_input:
                expr = statement.attrib["{%s}description" % self.namespaces['text']][1:-1]
                if expr=='_()':
                    expr = "_('%s')" % statement.text
                    expr = check_except_directive(expr) and "__filter(%s)" % expr
                else:
                    expr = check_except_directive(expr) and "__filter(%s)" % statement.attrib["{%s}description" % self.namespaces['text']][1:-1] or expr
            elif statement.tag == text_a:
                expr = urllib.unquote(statement.attrib[xlink_href_attrib][9:])
                expr = check_except_directive(expr) and "__filter(%s)" % urllib.unquote(statement.attrib[xlink_href_attrib][9:]) or expr

            if not expr:
                raise OOTemplateError("No expression in the tag",
                                      self.filepath)
            closing, directive, attr, attr_val = \
                    GENSHI_EXPR.match(expr).groups()
            is_opening = closing != '/'

            if directive is not None:
                dir_sequence.append((directive, expr, is_opening))
                # map closing tags with their opening tag
                if is_opening:
                    opened_tags.append(statement)
                else:
                    if not opened_tags:
                        continue
                    closing_tags[id(opened_tags.pop())] = statement
            # - we only need to return opening statements
            if is_opening:
                r_statements.append((statement,(expr, directive, attr, attr_val)))

        ##### Template syntax checking (has no opening/closing tags) #####
        check_opening = {}
        open_stm_by_type = {}
        while(dir_sequence):
            directive = dir_sequence.pop()
            check_opening.setdefault(directive[0], 0)
            if directive[2]: # only for opening tags
                if directive[0] not in open_stm_by_type:
                    open_stm_by_type.setdefault(directive[0], [directive[1]])
                else:
                    open_stm_by_type[directive[0]].append(directive[1])
            check_opening[directive[0]] += directive[2] and -1 or 1

        for stm in check_opening:
            if check_opening[stm]<0: # has no closing tag
                error_stm = open_stm_by_type[stm][check_opening[stm]-1]
                raise Exception("Statement has no closing tag. <%s>" % error_stm)
            if check_opening[stm]>0: # has no opening tag
                error_stm = "/%s" % stm
                raise Exception("Statement has no opening tag. <%s>" % error_stm)
        ##################################################################

        return r_statements, closing_tags

    def _handle_aeroo_tags(self, tree):
        """
        Will treat all aeroo tag (py:if/for/choose/when/otherwise)
        tags
        """
        # Some tag/attribute name constants
        table_namespace = self.namespaces['table']
        table_row_tag = '{%s}table-row' % table_namespace
        table_cell_tag = '{%s}table-cell' % table_namespace

        office_name = '{%s}value' % self.namespaces['office']
        office_valuetype = '{%s}value-type' % self.namespaces['office']

        py_attrs_attr = '{%s}attrs' % GENSHI_URI
        py_replace = '{%s}replace' % GENSHI_URI

        r_statements, closing_tags = self._aeroo_statements(tree)

        for r_node, parsed in r_statements:
            expr, directive, attr, a_val = parsed

            # If the node is a genshi directive statement:
            if directive is not None:
                opening = r_node
                closing = closing_tags[id(r_node)]

                # - we find the nearest common ancestor of the closing and
                #   opening statements
                o_ancestors = [opening]
                c_ancestors = [closing] + list(closing.iterancestors())
                ancestor = None
                for node in opening.iterancestors():
                    try:
                        idx = c_ancestors.index(node)
                        assert c_ancestors[idx] == node
                        # we only need ancestors up to the common one
                        del c_ancestors[idx:]
                        ancestor = node
                        break
                    except ValueError:
                        # c_ancestors.index(node) raise ValueError if node is
                        # not a child of c_ancestors
                        pass
                    o_ancestors.append(node)
                assert ancestor is not None, \
                       "No common ancestor found for opening and closing tag"

                outermost_o_ancestor = o_ancestors[-1]
                outermost_c_ancestor = c_ancestors[-1]

                # handle horizontal repetitions (over columns)
                if directive == "for" and ancestor.tag == table_row_tag:
                    a_val = self._handle_column_loops(parsed, ancestor,
                                                      opening,
                                                      outermost_o_ancestor,
                                                      outermost_c_ancestor)

                # - we create a <py:xxx> node
                if attr is not None:
                    attribs = {attr: a_val}
                else:
                    attribs = {}
                genshi_node = EtreeElement('{%s}%s' % (GENSHI_URI,
                                                       directive),
                                           attrib=attribs,
                                           nsmap={'py': GENSHI_URI})

                # - we move all the nodes between the opening and closing
                #   statements to this new node (append also removes from old
                #   parent)
                # - we replace the opening statement by the <py:xxx> node
                # - we delete the closing statement (and its ancestors)

                ######## This is need for removing only closing genshi tag except other children tags of <p> tag ########
                if len(outermost_c_ancestor)>1 and closing in outermost_c_ancestor.getchildren():
                    outermost_c_ancestor_copy = deepcopy(outermost_c_ancestor)
                    outermost_c_ancestor.remove(closing)
                    outermost_c_ancestor.getparent().append(outermost_c_ancestor_copy)
                    outermost_c_ancestor = outermost_c_ancestor_copy
                #########################################################################################################
                wrap_nodes_between(outermost_o_ancestor, outermost_c_ancestor, genshi_node)
            else:
                # It's not a genshi statement it's a python expression
                r_node.attrib[py_replace] = expr
                parent = r_node.getparent().getparent()
                if parent is None or parent.tag != table_cell_tag:
                    continue
                update_py_attrs(parent, "{'guess_type':1}")

    def _handle_column_loops(self, statement, ancestor, opening,
                             outer_o_node, outer_c_node):
        _, directive, attr, a_val = statement

        self.has_col_loop = True

        table_namespace = self.namespaces['table']
        table_col_tag = '{%s}table-column' % table_namespace
        table_num_col_attr = '{%s}number-columns-repeated' % table_namespace

        py_attrs_attr = '{%s}attrs' % GENSHI_URI
        repeat_tag = '{%s}repeat' % AEROO_URI

        # table node (it is not necessarily the direct parent of ancestor)
        table_node = ancestor.iterancestors('{%s}table' % table_namespace) \
                             .next()
        table_name = table_node.attrib['{%s}name' % table_namespace]

        # add counting instructions
        loop_id = id(opening)

        # 1) add reset counter code on the row opening tag
        #    (through a py:attrs attribute).
        # Note that table_name is not needed in the first two
        # operations, but a unique id within the table is required
        # to support nested column repetition
        update_py_attrs(ancestor, "__aeroo_reset_col_count(%d)" % loop_id)

        # 2) add increment code (through a py:attrs attribute) on
        #    the first cell node after the opening (cell node)
        #    ancestor
        enclosed_cell = outer_o_node.getnext()
        assert enclosed_cell.tag == '{%s}table-cell' % table_namespace
        update_py_attrs(enclosed_cell, "__aeroo_inc_col_count(%d)" %
                loop_id)

        # 3) add "store count" code as a py:replace node, as the
        #    last child of the row
        attr_value = "__aeroo_store_col_count(%d, %r)" \
                     % (loop_id, table_name)
        replace_node = EtreeElement('{%s}replace' % GENSHI_URI,
                                    attrib={'value': attr_value},
                                    nsmap={'py': GENSHI_URI})
        ancestor.append(replace_node)

        # find the position in the row of the cells holding the
        # <for> and </for> instructions
        # We use "*" so as to count both normal cells and covered/hidden cells
        position_xpath_expr = 'count(preceding-sibling::*)'
        opening_pos = \
            int(outer_o_node.xpath(position_xpath_expr,
                                   namespaces=self.namespaces))
        closing_pos = \
            int(outer_c_node.xpath(position_xpath_expr,
                                   namespaces=self.namespaces))

        # check whether or not the opening tag spans several rows
        a_val = self._handle_row_spanned_column_loops(
                    statement, outer_o_node, opening_pos, closing_pos)

        # check if this table's headers were already processed
        repeat_node = table_node.find(repeat_tag)
        if repeat_node is not None:
            prev_pos = (int(repeat_node.attrib['opening']),
                        int(repeat_node.attrib['closing']))
            if (opening_pos, closing_pos) != prev_pos:
                raise Exception(
                    'Incoherent column repetition found! '
                    'If a table has several lines with repeated '
                    'columns, the repetition need to be on the '
                    'same columns across all lines.')
        else:
            # compute splits: oo collapses the headers of adjacent
            # columns which use the same style. We need to split
            # any column header which is repeated so many times
            # that it encompasses any of the column headers that
            # we need to repeat
            to_split = []
            idx = 0
            childs = list(table_node.iterchildren(table_col_tag))
            for tag in childs:
                inc = int(tag.attrib.get(table_num_col_attr, 1))
                oldidx = idx
                idx += inc
                if oldidx < opening_pos < idx or \
                   oldidx < closing_pos < idx:
                    to_split.append(tag)

            # split tags
            for tag in to_split:
                tag_pos = table_node.index(tag)
                num = int(tag.attrib.pop(table_num_col_attr))
                new_tags = [deepcopy(tag) for _ in range(num)]
                table_node[tag_pos:tag_pos+1] = new_tags

            # recompute the list of column headers as it could
            # have changed.
            coldefs = list(table_node.iterchildren(table_col_tag))

            # compute the column header nodes corresponding to
            # the opening and closing tags.
            first = table_node[opening_pos]
            last = table_node[closing_pos]

            # add a <aeroo:repeat> node around the column
            # definitions nodes
            attribs = {
               "opening": str(opening_pos),
               "closing": str(closing_pos),
               "table": table_name
            }
            repeat_node = EtreeElement(repeat_tag, attrib=attribs,
                                       nsmap={'aeroo': AEROO_URI})
            wrap_nodes_between(first, last, repeat_node)
        return a_val

    def _handle_row_spanned_column_loops(self, statement, outer_o_node,
                                         opening_pos, closing_pos):
        """handles column repetitions which span several rows, by duplicating
        the py:for node for each row, and make the loops work on a copy of the
        original iterable as to not exhaust generators."""

        _, directive, attr, a_val = statement
        table_namespace = self.namespaces['table']
        table_rowspan_attr = '{%s}number-rows-spanned' % table_namespace

        # checks wether there is a (meaningful) rowspan
        rows_spanned = int(outer_o_node.attrib.get(table_rowspan_attr, 1))
        if rows_spanned == 1:
            return a_val

        table_row_tag = '{%s}table-row' % table_namespace
        table_cov_cell_tag = '{%s}covered-table-cell' % table_namespace

        # if so, we need to:

        # 1) create a with node to define a temporary variable
        temp_var = "__aeroo_temp%d" % id(outer_o_node)
        # a_val == "target in iterable"
        target, iterable = a_val.split(' in ', 1)
        vars = "%s = list(%s)" % (temp_var, iterable.strip())
        with_node = EtreeElement('{%s}with' % GENSHI_URI,
                                 attrib={"vars": vars},
                                 nsmap={'py': GENSHI_URI})

        # 2) transform a_val to use that temporary variable
        a_val = "%s in %s" % (target, temp_var)

        # 3) wrap the corresponding cells on the next row(s)
        #    (those should be covered-table-cell) inside a
        #    duplicate py:for node (looping on the temporary
        #    variable).
        row_node = outer_o_node.getparent()
        row_node.addprevious(with_node)
        rows_to_wrap = [row_node]
        assert row_node.tag == table_row_tag
        next_rows = row_node.itersiblings(table_row_tag)
        for row_idx in range(rows_spanned-1):
            next_row_node = next_rows.next()
            rows_to_wrap.append(next_row_node)
            # compute the start and end nodes
            first = next_row_node[opening_pos]
            last = next_row_node[closing_pos]
            assert first.tag == table_cov_cell_tag
            assert last.tag == table_cov_cell_tag
            # wrap them
            tag = '{%s}%s' % (GENSHI_URI, directive)
            for_node = EtreeElement(tag,
                                    attrib={attr: a_val},
                                    nsmap={'py': GENSHI_URI})
            wrap_nodes_between(first, last, for_node)

        # 4) wrap all the corresponding rows indide the "with"
        #    node
        for node in rows_to_wrap:
            with_node.append(node)
        return a_val

    def _handle_images(self, tree):
        "replaces all draw:frame named 'image: ...' by draw:image nodes"
        draw_namespace = self.namespaces['draw']
        draw_name = '{%s}name' % draw_namespace
        draw_image = '{%s}image' % draw_namespace
        py_attrs = '{%s}attrs' % self.namespaces['py']
        svg_namespace = self.namespaces['svg']
        svg_width = '{%s}width' % svg_namespace
        svg_height = '{%s}height' % svg_namespace
        xpath_expr = "//draw:frame[starts-with(@draw:name, 'image:')]"
        for draw in tree.xpath(xpath_expr, namespaces=self.namespaces):
            d_name = draw.attrib[draw_name][6:].strip()
            draw.attrib[draw_name] = "Aeroo picture "
            attr_expr = "__aeroo_make_href(%s)" % d_name
            image_node = EtreeElement(draw_image,
                                      attrib={py_attrs: attr_expr},
                                      nsmap={'draw': draw_namespace,
                                             'py': GENSHI_URI})
            draw.replace(draw[0], image_node)

    def _handle_hyperlinks(self, tree):
        xpath_href_expr = "//draw:a[starts-with(@xlink:href, 'python://')] | //text:a[starts-with(@xlink:href, 'pythonuri://')]"
        href_attrib = '{%s}href' % self.namespaces['xlink']
        py_attrs = '{%s}attrs' % self.namespaces['py']
        for a in tree.xpath(xpath_href_expr, namespaces=self.namespaces):
            a.attrib[py_attrs] = "__aeroo_hyperlink(%s)" % urllib.unquote(a.attrib[href_attrib]).replace('python://','').replace('pythonuri://','')#[9:]
            del a.attrib[href_attrib]

    def _handle_innerdocs(self, tree):
        "finds inner_docs and adds them to the processing stack."
        href_attrib = '{%s}href' % self.namespaces['xlink']
        xpath_expr = "//draw:object[starts-with(@xlink:href, './')" \
                     "and @xlink:show='embed']"
        for draw in tree.xpath(xpath_expr, namespaces=self.namespaces):
            self.inner_docs.append(draw.attrib[href_attrib][2:])

    def _escape_values(self, tree):
        "escapes element values"
        for element in tree.iter():
            for attrs in element.keys():
                if not attrs.startswith('{%s}' % GENSHI_URI):
                    element.attrib[attrs] = element.attrib[attrs]\
                            .replace(PREFIX, PREFIX * 2)
            if element.text:
                element.text = element.text.replace(PREFIX, PREFIX * 2)

    def generate(self, *args, **kwargs):
        "creates the AerooStream."
        #serializer = OOSerializer(self.filepath)

        self.Serializer.new_oo = BytesIO()
        outzip = zipfile.ZipFile(self.Serializer.new_oo, 'w')
        self.Serializer.outzip = outzip
        kwargs['__aeroo_make_href'] = ImageHref(self.namespaces, outzip, self.Serializer.manifest, kwargs)
        kwargs['__aeroo_hyperlink'] = _hyperlink(self.namespaces)
        if '__filter' not in kwargs:
            kwargs['__filter'] = _filter

        counter = ColumnCounter()
        kwargs['__aeroo_reset_col_count'] = counter.reset
        kwargs['__aeroo_inc_col_count'] = counter.inc
        kwargs['__aeroo_store_col_count'] = counter.store

        stream = super(Template, self).generate(*args, **kwargs)
        if self.has_col_loop:
            # Note that we can't simply add a "number-columns-repeated"
            # attribute and then fill it with the correct number of columns
            # because that wouldn't work if more than one column is repeated.
            transformation = DuplicateColumnHeaders(counter)
            col_filter = Transformer('//repeat[namespace-uri()="%s"]'
                                     % AEROO_URI)
            col_filter = col_filter.apply(transformation)
            stream = Stream(list(stream), self.serializer) | col_filter
        return AerooStream(stream, self.Serializer)


class DuplicateColumnHeaders(object):
    def __init__(self, counter):
        self.counter = counter

    def __call__(self, stream):
        for mark, (kind, data, pos) in stream:
            # for each repeat tag found
            if mark is ENTER:
                # get the number of columns for that table
                attrs = data[1]
                table = attrs.get('table')
                col_count = self.counter.counters[table]

                # collect events (column header tags) to repeat
                events = []
                for submark, event in stream:
                    if submark is EXIT:
                        break
                    events.append(event)

                # repeat them
                for _ in range(col_count):
                    for event in events:
                        yield None, event
            else:
                yield mark, (kind, data, pos)


class Manifest(object):

    def __init__(self, content):
        self.tree = lxml.etree.parse(BytesIO(content))
        self.root = self.tree.getroot()
        self.namespaces = self.root.nsmap
        for el in filter(lambda child: child.tag=="{%s}file-entry" % self.namespaces['manifest'], self.root.iterchildren()):
            if list(filter(lambda a: a[0]=='{%s}full-path' % self.namespaces['manifest'] and a[1].startswith("Thumbnails"), el.items())):
                self.root.remove(el)

    def __str__(self):
        return lxml.etree.tostring(self.tree, encoding='UTF-8',
                                   xml_declaration=True).decode("utf-8") 

    def add_file_entry(self, path, mimetype=None):
        manifest_namespace = self.namespaces['manifest']
        attribs = {'{%s}media-type' % manifest_namespace: mimetype or '',
                   '{%s}full-path' % manifest_namespace: path}
        entry_node = EtreeElement('{%s}%s' % (manifest_namespace,
                                              'file-entry'),
                                  attrib=attribs,
                                  nsmap={'manifest': manifest_namespace})
        self.root.append(entry_node)


class Meta(object):

    def __init__(self, content):
        self.tree = lxml.etree.parse(BytesIO(content))
        self.root = self.tree.getroot()
        self.namespaces = self.root.nsmap
        self.meta_root = self.root.getchildren()[0]
        self.add_entry('editing-cycles', 'meta', '1')
        self.add_entry('editing-duration', 'meta', '')

    def __str__(self):
        return lxml.etree.tostring(self.tree, encoding='UTF-8',
                                   xml_declaration=True).decode("utf-8") 

    def add_entry(self, tag_name, namespace, data=None, attribs={}):
        meta_namespace = self.namespaces['meta']
        entry_node = self.meta_root.find('{%s}%s' % (self.namespaces[namespace],tag_name))
        if entry_node is None:
            entry_node = EtreeElement('{%s}%s' % (self.namespaces[namespace],tag_name),
                                      nsmap={'meta': meta_namespace},
                                      attrib=attribs)
            self.meta_root.append(entry_node)
        else:
            entry_node.attrib.update(attribs)
        entry_node.text = data

    def add_property(self, data=None, ptype=''):
        meta_namespace = self.namespaces['meta']
        entry_node = EtreeElement('{%s}%s' % (self.namespaces['meta'],'user-defined'),
                                  nsmap={'meta': meta_namespace},
                                  attrib={'{%s}name' % self.namespaces['meta']:ptype})
        self.meta_root.append(entry_node)
        entry_node.text = data

class TStyle(object):

    def __init__(self, content):
        self.tree = lxml.etree.parse(BytesIO(content))
        self.root = self.tree.getroot()
        self.namespaces = self.root.nsmap

    def __str__(self):
        return lxml.etree.tostring(self.tree, encoding='UTF-8',
                                   xml_declaration=True).decode("utf-8") 

class TContent(object):

    def __init__(self, content):
        self.tree = lxml.etree.parse(BytesIO(content))
        self.root = self.tree.getroot()
        self.namespaces = self.root.nsmap
        ns = lxml.etree.FunctionNamespace(None)

    def __str__(self):
        return lxml.etree.tostring(self.tree, encoding='UTF-8',
                                   xml_declaration=True).decode("UTF-8")

class OOSerializer:

    def __init__(self, oo_template, oo_styles=None):
        self.template = oo_template
        self.inzip = zipfile.ZipFile(self.template)
        self.manifest = Manifest(self.inzip.read(MANIFEST))
        self.meta = Meta(self.inzip.read(META))
        self.content_xml = self.inzip.read('content.xml')
        self.styles_xml = self.inzip.read('styles.xml')
        self.apply_style(oo_styles)

        self.new_oo = None
        self.xml_serializer = genshi.output.XMLSerializer()



    def check_tabs(self, tree, namespaces):
        tags = tree.xpath("//text:*/text()[contains(string(), '%s')]/.." % TAB, namespaces=namespaces)
        for tag in tags:
            children = tag.getchildren()
            if tag.text:
                tabs_text=tag.text.split(TAB)
                tag.text=tabs_text.pop(0)
                pos = 0
                for text in tabs_text:
                    new_node = EtreeElement('{%s}tab' % namespaces['text'],
                                        attrib={},
                                        nsmap=namespaces)
                    new_node.tail=text
                    tag.insert(pos,new_node)
                    pos += 1
            else:
                new_node = EtreeElement('{%s}tab' % namespaces['text'],
                    attrib=tag.attrib,
                    nsmap=namespaces)
                tag.addnext(new_node)
            for child in tag.iterchildren():
                if not child.tail:
                    continue
                child_tail = child.tail
                child.tail = ''
                tail_text = child_tail.split(TAB)
                n = len(tail_text)-1
                while n:
                    new_node = EtreeElement('{%s}tab' % namespaces['text'],
                                attrib={},
                                nsmap=namespaces)
                    child.addnext(new_node)
                    n -= 1
                child.tail = tail_text.pop(0)
                last_node = child.getnext()
                for text in tail_text:
                    last_node.tail = text
                    last_node = last_node.getnext()


    def check_spaces(self, tree, namespaces):
        tags = tree.xpath("//text:*[starts-with(text(), ' ') or substring(name(),string-length(name())-1)=' ']", namespaces=namespaces)
        for tag in tags:
            children = tag.getchildren()
            if tag.text:
                tabs_text=tag.text.split(' ')
                tag.text=tabs_text.pop(0)
                pos = 0
                for text in tabs_text:
                    new_node = EtreeElement('{%s}s' % namespaces['text'],
                                        attrib={},
                                        nsmap=namespaces)
                    new_node.tail=text
                    tag.insert(pos,new_node)
                    pos += 1
            else:
                new_node = EtreeElement('{%s}s' % namespaces['text'],
                    attrib=tag.attrib,
                    nsmap=namespaces)
                tag.addnext(new_node)
            for child in tag.iterchildren():
                if not child.tail:
                    continue
                child_tail = child.tail
                child.tail = ''
                tail_text = child_tail.split(' ')
                n = len(tail_text)-1
                while n:
                    new_node = EtreeElement('{%s}s' % namespaces['text'],
                                attrib={},
                                nsmap=namespaces)
                    child.addnext(new_node)
                    n -= 1
                child.tail = tail_text.pop(0)
                last_node = child.getnext()
                for text in tail_text:
                    last_node.tail = text
                    last_node = last_node.getnext()

    def check_new_lines(self, tree, namespaces):
        #tags = tree.xpath("//text:*[contains(text(), '%s')]" % NEW_LINE, namespaces=namespaces)
        #tags = tree.xpath("//text:*[text()[contains(., '%s')]]" % NEW_LINE, namespaces=namespaces)
        span_tags = tree.xpath("//text:span[text()[contains(., '%s')]]" % NEW_LINE, namespaces=namespaces)
        tags = tree.xpath("//text:*[text()[contains(., '%s')]]" % NEW_LINE, namespaces=namespaces)
        tags = list(set(tags)-set(span_tags))
        #tags = tree.xpath("//text:*[text()[contains(., '%s')]]" % NEW_LINE, namespaces=namespaces)
        for tag in tags:
            children = tag.getchildren()
            if tag.text:
                for text in tag.text.split(NEW_LINE):
                    new_node = EtreeElement('%s' % tag.tag,
                                        attrib=tag.attrib,
                                        nsmap=namespaces)
                    new_node.text=text
                    tag.addprevious(new_node)
                last_node=new_node
            else:
                new_node = EtreeElement('%s' % tag.tag,
                    attrib=tag.attrib,
                    nsmap=namespaces)
                last_node = new_node
                tag.addnext(new_node)
            if children:
                for child in children:
                    if tag.text and tag.text.endswith(NEW_LINE):
                        new_node = EtreeElement('%s' % tag.tag,
                                        attrib=tag.attrib,
                                        nsmap=namespaces)
                        new_node.append(child)
                        last_node.addnext(new_node)
                        last_node = new_node
                        last_child = None
                    else:
                        last_node.append(child)
                        last_child = child

                    if not child.tail:
                        continue
                    child_tail = child.tail
                    child.tail = ''
                    tail_text = child_tail.split(NEW_LINE)
                    if tail_text[0] is not None and last_child is not None:
                        last_child.tail = tail_text.pop(0)
                    for text in tail_text:
                        new_node = EtreeElement('%s' % tag.tag,
                                    attrib=tag.attrib,
                                    nsmap=namespaces)
                        new_node.text=text
                        last_node.addnext(new_node)
                        last_node = new_node
            if tag.text:
                tag.getparent().remove(tag)

        for tag in span_tags:
            if tag.tag=='{%s}span' % namespaces['text']:
                span_texts = tag.text.split(NEW_LINE)
                parent = tag.getparent()
                parent_attrib = parent.attrib
                parent_children = parent.iterchildren()
                parent_itertext = parent.itertext()
                parent_text = parent.text
                parent_tail = parent.tail
                
                new_parent_node = EtreeElement('%s' % parent.tag,
                                    attrib=parent_attrib,
                                    nsmap=namespaces)
                parent.getparent().replace(parent, new_parent_node)
                new_parent_node.text = parent_text
                new_parent_node.tail = parent_tail
                next_child = new_parent_node
                try:
                    while(True):
                        curr_child = parent_children.next()
                        if curr_child.tag=='{%s}span' % namespaces['text'] and tag.text==curr_child.text:
                            new_span_node = EtreeElement('{%s}span' % namespaces['text'],
                                                attrib=tag.attrib,
                                                nsmap=namespaces)
                            new_span_node.text = span_texts.pop(0)
                            next_child.append(new_span_node)

                            curr_child = next_child
                            for text in span_texts:
                                new_node = EtreeElement('%s' % parent.tag,
                                                    attrib=parent_attrib,
                                                    nsmap=namespaces)
                                new_span_node = EtreeElement('{%s}span' % namespaces['text'],
                                                    attrib=tag.attrib,
                                                    nsmap=namespaces)
                                new_span_node.text=text
                                new_node.append(new_span_node)
                                curr_child.addnext(new_node)
                                curr_child = new_node
                            new_span_node.tail = tag.tail # set tail of source node in tail of latest new node
                            break
                        else:
                            next_child.append(curr_child)
                except StopIteration:
                    pass

                try:
                    next_text = True
                    while(next_text):
                        next_child = parent_children.next()
                        curr_child.append(next_child)
                        next_text = next_child.text
                except StopIteration:
                    pass
                try:
                    while(True):
                        next_child = parent_children.next()
                        if not next_child.text:
                            curr_child.append(next_child)
                        else:
                            new_node = EtreeElement('%s' % parent.tag,
                                                attrib=parent_attrib,
                                                nsmap=namespaces)
                            new_node.append(next_child)
                            curr_child.addnext(new_node)
                            curr_child = new_node
                        if next_child.tail:
                            next_child_texts = next_child.tail.split(NEW_LINE)
                            next_child.tail = next_child_texts.pop(0)
                            for next_text in next_child_texts:
                                new_node = EtreeElement('%s' % parent.tag,
                                            attrib=parent_attrib,
                                            nsmap=namespaces)
                                new_node.text=next_text
                                curr_child.addnext(new_node)
                                curr_child = new_node
                except StopIteration:
                    pass

    def check_guess_type(self, tree, namespaces):
        tags = tree.xpath('//table:table-cell[@guess_type]', namespaces=namespaces)
        for tag in tags:
            if len(tag)==0 or len(tag)>1 or tag[0].getchildren():
                guess_type = 'string'
            else:
                try:
                    float(tag[0].text)
                    guess_type = 'float'
                    tag.attrib['{%s}value' % namespaces['office']] = tag[0].text
                    # AKRETION HACK https://github.com/aeroo/aeroolib/issues/7
                    tag.attrib['{%s}value-type' % namespaces['calcext']] = guess_type
                except (ValueError,TypeError):
                    guess_type = 'string'
            tag.attrib['{%s}value-type' % namespaces['office']] = guess_type
            del tag.attrib['guess_type']

    def check_images(self, tree, namespaces):
        tags = tree.xpath('//draw:frame/draw:image[@svg:height and @svg:width]', namespaces=namespaces)
        for tag in tags:
            height = tag.attrib['{%s}height' % namespaces['svg']]
            width = tag.attrib['{%s}width' % namespaces['svg']]
            del tag.attrib['{%s}height' % namespaces['svg']]
            del tag.attrib['{%s}width' % namespaces['svg']]
            
            tag.getparent().attrib.update({'{%s}height' % namespaces['svg']:height,
                                           '{%s}width' % namespaces['svg']:width})



    def apply_style(self, oo_styles):
        if oo_styles:
            self.styles_zip = zipfile.ZipFile(oo_styles)
            self.styles_new = TStyle(self.styles_zip.read(STYLES))
            self.styles_new_xml = str(self.styles_new)
            self.styles = TStyle(self.inzip.read(STYLES))
            ##### font-face #####
            font_face_new = self.styles_new.root.xpath('//office:font-face-decls/style:font-face', \
                                                namespaces=self.styles_new.namespaces)
            font_face_orig = self.styles.root.xpath('//office:font-face-decls/style:font-face', \
                                                namespaces=self.styles.namespaces)
            self._replace_style_by_attrib(font_face_new, font_face_orig, 'name', self.styles.namespaces['style'])
            ###### styles ######
            new_styles = self.styles_new.root.xpath('//office:styles/style:style', \
                                                namespaces=self.styles_new.namespaces)
            orig_styles = self.styles.root.xpath('//office:styles/style:style', \
                                                namespaces=self.styles.namespaces)
            self._replace_style_by_attrib(new_styles, orig_styles, 'name', self.styles.namespaces['style'])
            ##### master-styles #####
            new_master_page_styles = self.styles_new.root.xpath('//office:master-styles/style:master-page', \
                                        namespaces=self.styles_new.namespaces)
            orig_master_page_styles = self.styles.root.xpath('//office:master-styles/style:master-page', \
                                        namespaces=self.styles.namespaces)
            self._replace_style_by_attrib(new_master_page_styles, orig_master_page_styles, 'name', \
                                        self.styles.namespaces['style'])
            ##### automatic-styles #####
            new_automatic_page_styles = self.styles_new.root.xpath('//office:automatic-styles/style:style', \
                                        namespaces=self.styles_new.namespaces)
            orig_automatic_page_styles = self.styles.root.xpath('//office:automatic-styles/style:style', \
                                        namespaces=self.styles.namespaces)
            if orig_automatic_page_styles:
                self._replace_style_by_attrib(new_automatic_page_styles, orig_automatic_page_styles, 'name', self.styles.namespaces['style'])
            else:
                dest_node = self.styles.root.xpath('//office:automatic-styles', namespaces=self.styles_new.namespaces)[0]
                self._add_styles(new_automatic_page_styles, dest_node)

            new_page_layout = self.styles_new.root.xpath('//office:automatic-styles/style:page-layout', \
                                                namespaces=self.styles_new.namespaces)
            orig_page_layout = self.styles.root.xpath('//office:automatic-styles/style:page-layout', \
                                                namespaces=self.styles.namespaces)
            self._replace_style_by_attrib(new_page_layout, orig_page_layout, 'name', self.styles.namespaces['style'])
            ###########################
            self.styles_xml = str(self.styles).encode("UTF-8")

    def add_title(self, title):
        self.meta.add_entry('title', 'dc', title)

    def add_creation_date(self, date):
        self.meta.add_entry('date', 'dc', date)
        self.meta.add_entry('creation-date', 'meta', date)

    def add_creation_user(self, user):
        self.meta.add_entry('creator', 'dc', user)

    def add_generator_info(self, data):
        self.meta.add_entry('generator', 'meta', data)

    def add_custom_property(self, data, ptype):
        self.meta.add_property(data, ptype)

    def _add_styles(self, node_list, dest_node):
        for node in node_list:
            dest_node.append(deepcopy(node))
        return dest_node

    def _replace_style_by_attrib(self, node_list1, node_list2, name, namespace):
        for node1 in node_list1:
            curr_node = None
            node_to_update = []
            for node2 in node_list2:
                if node2.get('{%s}%s' % (namespace,name))==node1.get('{%s}%s' % (namespace,name)):
                    curr_node = node2
                    node_to_update.append(node2)
                    break
            if curr_node is not None and curr_node.getchildren():
                for child_node in curr_node.iterchildren():
                    curr_node.attrib.update(dict(node1.attrib))
                    orig_node = node1.find(child_node.tag)
                    if orig_node is not None:
                        new_attribs = dict(child_node.attrib)
                        new_attribs.update(orig_node.attrib)
                        new_child_node = EtreeElement(child_node.tag, attrib=new_attribs)
                        new_child_node.append(deepcopy(child_node))
                        orig_subchildren = orig_node.getchildren()
                        if orig_subchildren:
                            self._add_styles(orig_subchildren, new_child_node)
                        curr_node.replace(child_node, new_child_node)
                        node1.remove(orig_node)
                if node1 is not None:
                    self._add_styles(node1, curr_node)
            else:
                node_list2[0].getparent().append(deepcopy(node1))
        return node_list2

    def __call__(self, stream):
        files = {}
        for kind, data, pos in stream:
            if kind == genshi.core.PI and data[0] == 'aeroo':
                stream_for = data[1]
                continue
            files.setdefault(stream_for, []).append((kind, data, pos))

        now = time.localtime()[:6]
        outzip = self.outzip
        self.template.seek(0)
        inzip = zipfile.ZipFile(BytesIO(self.template.read()))
        for f_info in inzip.infolist():
            if f_info.filename.startswith('ObjectReplacements'):
                continue
            elif f_info.filename in files:
                stream = files[f_info.filename]
                # create a new file descriptor, copying some attributes from
                # the original file
                new_info = zipfile.ZipInfo(f_info.filename, now)
                for attr in ('compress_type', 'flag_bits', 'create_system'):
                    setattr(new_info, attr, getattr(f_info, attr))
                serialized_stream = output_encode(self.xml_serializer(stream), encoding='utf-8')
                ############### Styles usage #################
                if f_info.filename == STYLES:
                    self.styles_orig = TStyle(serialized_stream)
                    self.check_guess_type(self.styles_orig.tree, self.styles_orig.namespaces)
                    self.check_images(self.styles_orig.tree, self.styles_orig.namespaces)
                    self.check_new_lines(self.styles_orig.tree, self.styles_orig.namespaces)
                    self.check_tabs(self.styles_orig.tree, self.styles_orig.namespaces)
                    self.check_spaces(self.styles_orig.tree, self.styles_orig.namespaces)
                    if hasattr(self, 'styles_new'):
                        pictures = []
                        for file_name in self.styles_zip.namelist():
                            if file_name.startswith('Pictures'):
                                picture = self.styles_zip.read(file_name)
                                pictures.append((file_name, picture))
                                self.manifest.add_file_entry(file_name)
                        for ffile, picture in pictures:
                            outzip.writestr(ffile, picture)

                    outzip.writestr(new_info, str(self.styles_orig))
                ###############################################
                elif f_info.filename == 'content.xml':
                    content = TContent(serialized_stream)
                    self.check_guess_type(content.tree, content.namespaces)
                    self.check_images(content.tree, content.namespaces)
                    self.check_new_lines(content.tree, content.namespaces)
                    self.check_tabs(content.tree, content.namespaces)
                    self.check_spaces(content.tree, content.namespaces)
                    outzip.writestr(new_info, str(content))
                else:
                    outzip.writestr(new_info, serialized_stream)
            elif f_info.filename == MANIFEST:
                outzip.writestr(f_info, str(self.manifest))
            elif f_info.filename == META:
                outzip.writestr(f_info, str(self.meta))
            elif f_info.filename=='Thumbnails/thumbnail.png':
                continue
            else:
                outzip.writestr(f_info, inzip.read(f_info.filename))
        inzip.close()
        outzip.close()

        return self.new_oo is not None and self.new_oo or BytesIO()

MIMETemplateLoader.add_factory('oo.org', Template)

