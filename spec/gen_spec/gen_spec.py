#!/usr/bin/env python
# -*- coding: utf-8 -*-

# gen_spec.py
#

import argparse
import os
import sqlite3
import sys

from lxml import objectify

sys.path.append('/Users/scanny/Dropbox/src/python-pptx')

# import pdb; pdb.set_trace()


"""
Utility code to load spec database and generate portions of constants.py and
spec.py

Fields
----------------------------------------

id_
   Distinct integer identifier, defined in the MS API specification, although
   that spec is incomplete. When not documented, it is discovered by
   inspection in the PowerPoint built-in VBA IDE.

ms_name
   name assigned to the auto shape type in the MS API, e.g.
   'msoShapeRoundedRectangle'.

desc
   description of auto shape type from MS API specification, invented for
   auto shapes not documented by Microsoft.

const_name
   upper snake case name used for symbolic constant representing id of the
   auto shape, e.g. 'LINE_CALLOUT_1_ACCENT_BAR'.

base_name
   base of name an instance of this auto shape type is assigned in the XML,
   e.g. '10-Point Star'. In the XML, an integer suffix is appended to create
   the shape name, e.g. 'Rounded Rectangle 9'.

prst
   string key used in the XML to specify the type of an auto shape element,
   e.g. 'roundRect'.

adj_vals
   sequence of (name, val) tuples specifying the available adjustment values
   for the auto shape type along with the default value for each.


Sources
----------------------------------------

src_data/msoAutoShapeType.py
   Tuples containing all the components of the full definition from the
   various sources, documented, generated, and hand-coded where necessary.

presetShapeDefinitions.xml
   The adjustment values are pulled from the XML file in which they are
   defined, distributed as part of the ISO/IEC spec.


Outputs
----------------------------------------

constants.py
spec.py
autoshapetypes.rst (not implemented yet)

"""


class MsoAutoShapeTypeCollection(list):
    """auto shape type documented by Microsoft"""
    def __init__(self):
        super(MsoAutoShapeTypeCollection, self).__init__()

    @classmethod
    def load(cls, sort=None):
        conn = sqlite3.connect('spec.db')
        c = conn.cursor()
        c.execute(
            '  SELECT id, prst, const_name, base_name, ms_name, desc\n'
            '    FROM auto_shape_types\n'
            'ORDER BY const_name'
        )
        mastc = cls()
        for id_, prst, const_name, base_name, ms_name, desc in c:
            mast = MsoAutoShapeType(id_, prst, const_name, base_name,
                                    ms_name, desc)
            mastc.append(mast)
        mastc.load_adjustment_values(c)
        conn.close()
        return mastc

    def load_adjustment_values(self, c):
        """load adjustment values for auto shape types in self"""
        # retrieve auto shape types in const_name order --------
        for mast in self:
            # retriev adj vals for this auto shape type --------
            c.execute(
                '  SELECT name, val\n'
                '    FROM adjustment_values\n'
                '   WHERE prst = ?\n'
                'ORDER BY seq_nmbr', (mast.prst,)
            )
            for name, val in c:
                mast.adj_vals.append(AdjustmentValue(name, val))


class MsoAutoShapeType(object):
    def __init__(self, id_, prst, const_name, base_name, ms_name, desc):
        super(MsoAutoShapeType, self).__init__()
        self.id_ = id_
        self.prst = prst
        self.const_name = const_name
        self.base_name = base_name
        self.ms_name = ms_name
        self.desc = desc
        self.adj_vals = []


class AdjustmentValue(object):
    def __init__(self, name, val):
        super(AdjustmentValue, self).__init__()
        self.name = name
        self.val = val


def create_tables(c):
    """create (or recreate) the auto shape type tables"""
    # auto_shape_types ---------------------
    c.execute(
        'DROP TABLE IF EXISTS auto_shape_types'
    )
    c.execute(
        'CREATE TABLE auto_shape_types (\n'
        '    id         integer,\n'
        '    prst       text,\n'
        '    const_name text,\n'
        '    base_name  text,\n'
        '    ms_name    text,\n'
        '    desc       text\n'
        ')\n'
    )
    c.execute(
        'DROP TABLE IF EXISTS adjustment_values'
    )
    c.execute(
        'CREATE TABLE adjustment_values (\n'
        '    prst     text,\n'
        '    seq_nmbr integer,\n'
        '    name     text,\n'
        '    val      integer\n'
        ')\n'
    )


def insert_adjustment_values(c):
    """insert adjustment values into their table"""
    adjustment_values = load_adjustment_values()
    # insert into table ------------------------------------
    q_insert = (
        'INSERT INTO adjustment_values\n'
        '            (prst, seq_nmbr, name, val)\n'
        '     VALUES (?, ?, ?, ?)\n'
    )
    c.executemany(q_insert, adjustment_values)


def load_adjustment_values():
    """load adjustment values and their default values from XML"""
    # parse XML --------------------------------------------
    thisdir = os.path.split(__file__)[0]
    prst_defs_relpath = (
        'ISO-IEC-29500-1/schemas/dml-geometries/OfficeOpenXML-DrawingMLGeomet'
        'ries/presetShapeDefinitions.xml'
    )
    prst_defs_path = os.path.join(thisdir, prst_defs_relpath)
    presetShapeDefinitions = objectify.parse(prst_defs_path).getroot()
    # load individual records into tuples to return --------
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    avLst_qn = '{%s}avLst' % ns
    adjustment_values = []
    for shapedef in presetShapeDefinitions.iterchildren():
        prst = shapedef.tag
        try:
            avLst = shapedef[avLst_qn]
        except AttributeError:
            continue
        for idx, gd in enumerate(avLst.gd):
            name = gd.get('name')
            val = int(gd.get('fmla')[4:])  # strip off leading 'val '
            record = (prst, idx+1, name, val)
            adjustment_values.append(record)
    return adjustment_values


def print_mso_auto_shape_type_constants():
    """print symbolic constant definitions for msoAutoShapeType"""
    auto_shape_types = MsoAutoShapeTypeCollection.load(sort='const_name')
    out = render_mso_auto_shape_type_constants(auto_shape_types)
    print out


def print_mso_auto_shape_type_enum():
    """print symbolic constant definitions for msoAutoShapeType"""
    auto_shape_types = MsoAutoShapeTypeCollection.load(sort='const_name')
    out = render_mso_auto_shape_type_enum(auto_shape_types)
    print out


def print_mso_auto_shape_type_spec():
    """print spec dictionary for msoAutoShapeType"""
    auto_shape_types = MsoAutoShapeTypeCollection.load(sort='const_name')
    out = render_mso_auto_shape_type_spec(auto_shape_types)
    print out


def render_adj_vals(adj_vals):
    # calculate adj_vals string, possibly multi-line ---
    if adj_vals:
        lines = []
        for av in adj_vals:
            indent = ' ' * 12
            line = "%s('%s', %d)," % (indent, av.name, av.val)
            lines.append(line)
        lines_str = '\n'.join(lines)
        out = '(\n%s\n        )' % lines_str
    else:
        out = '()'
    return out


def render_desc(desc):
    """calculate desc string, wrapped if too long"""
    desc = desc + '.'
    desc_lines = split_len(desc, 54)
    if len(desc_lines) > 1:
        join_str = "'\n%s'" % (' '*21)
        lines_str = join_str.join(desc_lines)
        out = "('%s')" % lines_str
    else:
        out = "'%s'" % desc_lines[0]
    return out


def render_mso_auto_shape_type_constants(auto_shape_types):
    out = (
        'class MSO_AUTO_SHAPE_TYPE(object):\n'
        '    """\n'
        '    Constants corresponding to the msoAutoShapeType enumeration in '
        'the\n    MS API. Standard abbreviation is \'MAST\', e.g.:\n\n      '
        '  from pptx.spec import MSO_AUTO_SHAPE_TYPE as MAST\n\n'
        '    """\n'
        '    # msoAutoShapeType -----------------\n'
    )
    for ast in auto_shape_types:
        out += '    %s = %d\n' % (ast.const_name, ast.id_)
    return out


def render_mso_auto_shape_type_enum(auto_shape_types):
    out = (
        'class MSO_AUTO_SHAPE_TYPE(XmlEnumeration):\n'
        '    """\n'
        '    Specifies a type of AutoShape, e.g. DOWN_ARROW\n'
        '    """\n'
        '\n'
        '    __members__ = (\n'
    )
    for ast in auto_shape_types:
        tmpl = (
            "        XmlMappedEnumMember(\n"
            "            '%s', %d, '%s', '%s'\n"
            "        ),\n"
        )
        args = (ast.const_name, ast.id_, ast.prst, ast.desc)
        out += tmpl % args
        # break
    out += '    )'
    return out


def render_mso_auto_shape_type_spec(auto_shape_types):
    out = 'autoshape_types = {\n'
    for idx, ast in enumerate(auto_shape_types):
        out += (
            "    MAST.%s: {\n"
            "        'basename': '%s',\n"
            "        'prst':     '%s',\n"
            "        'avLst':    %s\n"
            "    },\n" %
            (ast.const_name, ast.base_name, ast.prst,
             render_adj_vals(ast.adj_vals))
        )
    out += '}'
    return out


def split_len(s, length):
    """split string *s* into list of strings no longer than *length*"""
    return [s[i:i+length] for i in range(0, len(s), length)]


def to_mixed_case(s):
    """
    convert upper snake case string to mixed case, e.g. MIXED_CASE becomes
    MixedCase
    """
    out = ''
    last_c = ''
    for c in s:
        if c == '_':
            pass
        elif last_c in ('', '_'):
            out += c.upper()
        else:
            out += c.lower()
        last_c = c
    return out


# ===========================================================================
# load auto shape types into database
# ===========================================================================
# this code is broken now, in particular the insert functions need to be
# connected to the new file src_data/msoAutoShapeType.py
# ===========================================================================

# conn = sqlite3.connect('spec.db')
# c = conn.cursor()

# # create_tables(c)
# # insert_adjustment_values(c)

# conn.commit()
# conn.close()


# ============================================================================
# CLI parser
# ============================================================================

def parse_args(spectypes):
    """
    Return arguments object formed by parsing the command line used to launch
    the program.
    """
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument(
        "-c", "--constants",
        help="emit constants instead of spec dict",
        action="store_true"
    )
    arg_parser.add_argument(
        "spectype",
        help="specifies the spec type to be generated",
        choices=spectypes
    )
    return arg_parser.parse_args()


# ============================================================================
# main
# ============================================================================

spectypes = ('msoAutoShapeType', 'MSO_AUTO_SHAPE_TYPE')

args = parse_args(spectypes)

# call print function for requested spectype ---------------
if args.constants:
    if args.spectype == 'msoAutoShapeType':
        print_mso_auto_shape_type_constants()
else:
    if args.spectype == 'msoAutoShapeType':
        print_mso_auto_shape_type_spec()
    elif args.spectype == 'MSO_AUTO_SHAPE_TYPE':
        print_mso_auto_shape_type_enum()

# avoid error message "close failed ... sys.excepthook is missing ...
sys.stdout.flush()
