.. _XlDisplayBlanksAs:

``XL_DISPLAY_BLANKS_AS``
========================

Specifies how blank (empty) cells are displayed in a chart. Corresponds to the
PowerPoint "Hidden and Empty Cells" dialog option "Show empty cells as".

Example::

    from pptx.enum.chart import XL_DISPLAY_BLANKS_AS

    chart.display_blanks_as = XL_DISPLAY_BLANKS_AS.GAPS

----

GAPS
    Blank cells are not plotted (leave a gap in the chart).

INTERPOLATED
    Blank cells are interpolated (span connects surrounding points).

ZERO
    Blank cells are plotted as zero.
