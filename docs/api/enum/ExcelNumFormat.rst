.. _ExcelNumFormat:

Excel Number Formats
====================

The following integer values correspond to built-in Excel number formats.
While they cannot be used directly in |pp|, this reference can be used to
determine the format strings that can be substituted in
:meth:`.ChartData.add_series` to specify a numeric display format for series
values.

Further information on string number format codes (such as '#,##0') can be
found on `this web page`_.

.. _`this web page`:
   https://support.office.com/en-GB/article/
   Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68

----

=====  ==========  ==========================
Value  Type        Format String
-----  ----------  --------------------------
0      General     General
1      Decimal     0
2      Decimal     0.00
3      Decimal     #,##0
4      Decimal     #,##0.00
5      Currency    $#,##0;$-#,##0
6      Currency    $#,##0;[Red]$-#,##0
7      Currency    $#,##0.00;$-#,##0.00
8      Currency    $#,##0.00;[Red]$-#,##0.00
9      Percentage  0%
10     Percentage  0.00%
11     Scientific  0.00E+00
12     Fraction    # ?/?
13     Fraction    # /
14     Date        m/d/yy
15     Date        d-mmm-yy
16     Date        d-mmm
17     Date        mmm-yy
18     Time        h:mm AM/PM
19     Time        h:mm:ss AM/PM
20     Time        h:mm
21     Time        h:mm:ss
22     Time        m/d/yy h:mm
37     Currency    #,##0;-#,##0
38     Currency    #,##0;[Red]-#,##0
39     Currency    #,##0.00;-#,##0.00
40     Currency    #,##0.00;[Red]-#,##0.00
41     Accounting  _ * #,##0_ ;_ * "_ ;_ @_
42     Accounting  _ $* #,##0_ ;_ $* "_ ;_ @_
43     Accounting  _ * #,##0.00_ ;_ * "??_ ;_ @_
44     Accounting  _ $* #,##0.00_ ;_ $* "??_ ;_ @_
45     Time        mm:ss
46     Time        h :mm:ss
47     Time        mm:ss.0
48     Scientific  ##0.0E+00
49     Text        @
=====  ==========  ==========================
