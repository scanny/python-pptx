"""Legend of a chart."""

from __future__ import annotations

from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.text.text import Font
from pptx.util import lazyproperty


class Legend(object):
    """
    Represents the legend in a chart. A chart can have at most one legend.
    """

    def __init__(self, legend_elm):
        super(Legend, self).__init__()
        self._element = legend_elm

    def exclude_entry(self, idx):
        """Hide the legend entry at 0-based position *idx*.

        The corresponding series (or, for pie/doughnut/funnel style charts,
        data point) remains plotted — only its entry in the legend is
        suppressed. This is the programmatic equivalent of PowerPoint's
        right-click *Format Legend Entry → Delete* on a single legend entry.

        Calling :meth:`exclude_entry` for an ``idx`` that is already hidden
        is a no-op. Passing a non-int raises :class:`TypeError`. GitHub
        issue #649.

        .. versionadded:: 2026.05.0
        """
        idx = self._coerce_idx(idx)
        legendEntry = self._element.get_or_add_legendEntry_for_idx(idx)
        delete_ = legendEntry.get_or_add_delete_()
        # -- CT_Boolean omits `val` on the default True, but PowerPoint itself
        # -- always writes the attribute and some PowerPoint builds interpret
        # -- the missing form as False. Set it explicitly for interop.
        delete_.set("val", "1")

    def include_entry(self, idx):
        """Re-show a legend entry previously hidden via :meth:`exclude_entry`.

        If the entry at *idx* carries no other overrides (e.g. a ``c:txPr``
        formatting block) the whole ``c:legendEntry`` is removed, keeping
        the XML minimal; otherwise only the ``c:delete`` flag is dropped.
        Calling :meth:`include_entry` for an ``idx`` that is not currently
        excluded is a no-op. GitHub issue #649.

        .. versionadded:: 2026.05.0
        """
        idx = self._coerce_idx(idx)
        legendEntry = self._element.get_legendEntry_for_idx(idx)
        if legendEntry is None:
            return
        legendEntry._remove_delete_()
        # -- drop the whole c:legendEntry when only c:idx remains --
        if legendEntry.txPr is None and len(legendEntry) == 1:
            self._element.remove(legendEntry)

    @property
    def hidden_entries(self) -> tuple[int, ...]:
        """Read-only tuple of integer indices of currently hidden legend entries.

        Indices are 0-based and correspond to the ``c:idx/@val`` of each
        ``c:legendEntry`` carrying ``c:delete/@val="1"``. The tuple is
        emitted in document order (i.e. the order the ``c:legendEntry``
        elements appear in the XML). GitHub issue #649.

        .. versionadded:: 2026.05.0
        """
        return self._element.hidden_entry_idxs

    @lazyproperty
    def font(self) -> Font:
        """
        The |Font| object that provides access to the text properties for
        this legend, such as bold, italic, etc.
        """
        defRPr = self._element.defRPr
        font = Font(defRPr)
        return font

    @property
    def horz_offset(self) -> float | None:
        """
        Adjustment of the x position of the legend from its default.
        Expressed as a float between -1.0 and 1.0 representing a fraction of
        the chart width. Negative values move the legend left, positive
        values move it to the right. |None| if no setting is specified.
        """
        return self._element.horz_offset

    @horz_offset.setter
    def horz_offset(self, value):
        self._element.horz_offset = value

    @property
    def include_in_layout(self) -> bool:
        """|True| if legend should be located inside plot area.

        Read/write boolean specifying whether legend should be placed inside
        the plot area. In many cases this will cause it to be superimposed on
        the chart itself. Assigning |None| to this property causes any
        `c:overlay` element to be removed, which is interpreted the same as
        |True|. This use case should rarely be required and assigning
        a boolean value is recommended.
        """
        overlay = self._element.overlay
        if overlay is None:
            return True
        return overlay.val

    @include_in_layout.setter
    def include_in_layout(self, value):
        if value is None:
            self._element._remove_overlay()
            return
        self._element.get_or_add_overlay().val = bool(value)

    @property
    def position(self):
        """
        Read/write :ref:`XlLegendPosition` enumeration value specifying the
        general region of the chart in which to place the legend.
        """
        legendPos = self._element.legendPos
        if legendPos is None:
            return XL_LEGEND_POSITION.RIGHT
        return legendPos.val

    @position.setter
    def position(self, position):
        self._element.get_or_add_legendPos().val = position

    @staticmethod
    def _coerce_idx(idx):
        """Return *idx* as ``int``, rejecting ``bool`` and non-integer input."""
        if isinstance(idx, bool) or not isinstance(idx, int):
            raise TypeError("idx must be an int, got %r" % type(idx).__name__)
        if idx < 0:
            raise ValueError("idx must be non-negative, got %d" % idx)
        return idx
