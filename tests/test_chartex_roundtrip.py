# pyright: reportPrivateUsage=false

"""Regression test — ``chartEx`` fallback parts dropped on round-trip.

Opening a ``.pptx`` containing a ``cx:chartSpace`` (``chartEx``) chart and
saving it back produced an output that PowerPoint refused to render because
four reachable-by-PowerPoint parts were silently dropped from the saved
package:

* ``ppt/charts/chart1.xml`` — the traditional-chart fallback associated
  with the ``chartEx`` via the ``mc:AlternateContent`` mechanism. It is
  present as an ``Override`` entry in ``[Content_Types].xml`` but is *not*
  reachable via the relationships graph rooted at the package, so the
  rel-walk save path missed it.
* ``ppt/charts/_rels/chart1.xml.rels`` — the fallback chart's own rels
  file, which in turn points at the embedded workbook.
* ``ppt/charts/_rels/chartEx1.xml.rels`` — empty in this fixture but
  physically present in the input zip; its disappearance on output caused
  a visible byte-level diff even though its contents were empty.
* ``ppt/embeddings/Microsoft_Excel_Sheet1.xlsx`` — the workbook that
  backs the fallback chart, reachable only through the orphan chart's
  rels.

The fix widens the package load/save path to preserve *every* part
declared in ``[Content_Types].xml`` (including "orphan" parts that are
not reached via any relationship) and to preserve an empty ``.rels``
file when one was physically present in the input package.
"""

from __future__ import annotations

import io
import zipfile

from pptx import Presentation


class DescribeChartExRoundTrip:
    """Unit-test suite for ``chartEx`` fallback part preservation."""

    def it_preserves_the_chartEx_fallback_chart_and_embedded_workbook(self):
        src = "features/steps/test_files/cht-chartex.pptx"

        # -- round-trip via an in-memory buffer to avoid touching the filesystem --
        out = io.BytesIO()
        Presentation(src).save(out)
        out.seek(0)

        with zipfile.ZipFile(src) as zf:
            src_names = set(zf.namelist())
        with zipfile.ZipFile(out) as zf:
            out_names = set(zf.namelist())

        # -- the four parts PowerPoint needs to render the fallback chart must survive --
        required = {
            "ppt/charts/chart1.xml",
            "ppt/charts/_rels/chart1.xml.rels",
            "ppt/charts/_rels/chartEx1.xml.rels",
            "ppt/embeddings/Microsoft_Excel_Sheet1.xlsx",
        }
        assert required.issubset(src_names), (
            "fixture has changed; expected the four fallback-chart parts to be present "
            "in the input package (see test docstring)"
        )
        missing_on_output = required - out_names
        assert not missing_on_output, (
            f"round-trip dropped required chartEx fallback parts: "
            f"{sorted(missing_on_output)}"
        )

        # -- the chartEx part itself must also survive --
        assert "ppt/charts/chartEx1.xml" in out_names

    def and_it_preserves_the_entire_set_of_member_names_on_round_trip(self):
        src = "features/steps/test_files/cht-chartex.pptx"

        out = io.BytesIO()
        Presentation(src).save(out)
        out.seek(0)

        with zipfile.ZipFile(src) as zf:
            src_names = set(zf.namelist())
        with zipfile.ZipFile(out) as zf:
            out_names = set(zf.namelist())

        # -- every input zip member should be present in the output. (The reverse is
        # -- not asserted here: the writer may materialize lazily-created parts such
        # -- as docProps/app.xml that weren't physically present in the input.)
        only_in_src = src_names - out_names
        assert not only_in_src, (
            f"round-trip dropped input package members: {sorted(only_in_src)}"
        )
