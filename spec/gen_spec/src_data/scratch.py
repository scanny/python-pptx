#!/usr/bin/env python
# -*- coding: utf-8 -*-

# scratch.py
#
# utility code to generate one-time text runs from spec.db

import sqlite3


# establish database connection ------------------------
conn = sqlite3.connect('../spec.db')
c = conn.cursor()

# # retrieve ms_name to const_name mapping in prst order -
# c.execute(
#     '  SELECT ms_name, const_name\n'
#     '    FROM auto_shape_types\n'
#     'ORDER BY ms_name'
# )
# lines = []
# for ms_name, const_name in c:
#     lines.append("    ('%s', '%s')" % (ms_name, const_name))
# const_name_list_str = ',\n'.join(lines)

# out = 'const_name_map = (\n%s\n)' % const_name_list_str

# retrieve ms_name to PresentationML mappings --------------
c.execute(
    '  SELECT ms_name, prst, base_name\n'
    '    FROM auto_shape_types\n'
    'ORDER BY ms_name'
)
lines = []
for ms_name, prst, base_name in c:
    lines.append("        ('%s', '%s', '%s')" % (ms_name, prst, base_name))
pml_list_str = ',\n'.join(lines)

out = 'def pml_map():\n    return (\n%s\n    )' % pml_list_str

# tear down database connection ------------------------
conn.commit()
conn.close()

print out
