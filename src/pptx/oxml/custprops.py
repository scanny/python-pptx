"""lxml custom element classes for custom-properties-related XML elements.

The custom-properties part corresponds to ``/docProps/custom.xml`` in the OPC package. Its
``<Properties>`` root element holds user-defined document properties (the "Custom" tab of
PowerPoint's Document-Properties dialog). Each child ``<property>`` element has a unique
``name`` attribute, a mandatory ``fmtid`` (always the PID_CUSTOM FMTID per ECMA-376
§15.2.12.2), and a ``pid`` starting at 2 and incrementing per property. The property's
value is carried by a single typed child element from the ``vt:`` namespace
(``vt:lpwstr`` for strings, ``vt:i4`` for integers, ``vt:r8`` for floats, ``vt:bool`` for
booleans, and ``vt:filetime`` for datetimes).

python-pptx exposes this part through :attr:`Presentation.custom_properties`, a
dict-like view surfaced via :class:`~pptx.parts.customprops.CustomPropertiesPart`. See
issue #259.
"""

from __future__ import annotations

import datetime as dt
from typing import Iterator, Union, cast

from lxml.etree import SubElement, _Element

from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import BaseOxmlElement

# -- the `custom-properties` namespace URI; children live here with the default namespace
# -- (no explicit prefix). The `vt:` namespace carries the typed value elements.
_CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
_VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

# -- The FMTID constant PID_CUSTOM identifies the default OOXML custom-properties set.
# -- Per ECMA-376 §15.2.12.2 every `<property>` element in `docProps/custom.xml` must
# -- carry exactly this value. The GUID is well-known and appears verbatim in every
# -- OOXML producer's output (LibreOffice, Microsoft Office, Apple Pages, etc.).
FMTID_CUSTOM = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

# -- The smallest valid `pid` per ECMA-376 §15.2.12.2 is 2 (pid=0 and pid=1 are reserved
# -- for the dictionary and code-page properties respectively, which are not used in the
# -- custom-properties part).
MIN_PID = 2

# -- A python value that can be round-tripped as a custom property. Mirrors the types
# -- supported by the six vt: elements emitted by :meth:`CT_CustomProperties.set_value`.
# -- Note: `datetime` subclasses `date`, so `dt.date` covers both — the encoder
# -- dispatches `datetime` to ``vt:filetime`` and a plain `date` to ``vt:date``.
CustomPropValue = Union[str, int, float, bool, dt.date]


class CT_CustomProperties(BaseOxmlElement):
    """`Properties` element, root of the custom-properties part (``/docProps/custom.xml``)."""

    _Properties_tmpl = '<Properties xmlns="%s" xmlns:vt="%s"/>\n' % (_CUSTOM_NS, _VT_NS)

    @staticmethod
    def new_customProperties() -> CT_CustomProperties:
        """Return a newly-created, empty `Properties` element.

        Contains no `<property>` children; callers add them via :meth:`set_value`.
        """
        return cast(CT_CustomProperties, parse_xml(CT_CustomProperties._Properties_tmpl))

    def __len__(self) -> int:
        return len(self._property_elms)

    def __iter__(self) -> Iterator[str]:
        """Yield each custom-property name in document order."""
        for prop in self._property_elms:
            name = prop.get("name")
            if name is not None:
                yield name

    def names(self) -> list[str]:
        """Return the list of custom-property names in document order."""
        return list(iter(self))

    def has_name(self, name: str) -> bool:
        """Return True if a `<property>` child has `name=name`."""
        return self._find_property(name) is not None

    def get_value(self, name: str) -> CustomPropValue | None:
        """Return the python value of the named custom property, or None if not present.

        The python type is inferred from the first typed-vt child of the matching
        `<property>` element. See `_python_value_of` for the mapping.
        """
        prop = self._find_property(name)
        if prop is None:
            return None
        return _python_value_of(prop)

    def set_value(self, name: str, value: CustomPropValue) -> None:
        """Set (or replace) the named custom property to `value`.

        The value's python type determines the vt: child element emitted:
        ``str`` -> ``vt:lpwstr``, ``bool`` -> ``vt:bool``, ``int`` -> ``vt:i4``
        (clipped at the 32-bit signed int range; wider ints raise ``OverflowError``),
        ``float`` -> ``vt:r8``, ``datetime`` -> ``vt:filetime``. Unsupported types
        raise ``TypeError``.

        A pre-existing property of the same name is updated in place, preserving the
        ``pid`` it was originally assigned (PowerPoint tolerates non-contiguous pids
        and matches on ``name``, so preserving the existing value avoids unnecessary
        churn in the XML diff between a round-trip open/save).
        """
        if not isinstance(name, str) or not name:  # pyright: ignore[reportUnnecessaryIsInstance]
            raise ValueError("custom-property name must be a non-empty string")
        vt_tag, vt_text = _vt_tag_and_text_for(value)
        prop = self._find_property(name)
        if prop is None:
            prop = SubElement(self, qn("cst:property"))
            prop.set("fmtid", FMTID_CUSTOM)
            prop.set("pid", str(self._next_pid()))
            prop.set("name", name)
        else:
            # -- remove any existing typed-vt child (there should be exactly one) --
            for child in list(prop):
                prop.remove(child)
        value_elm = SubElement(prop, qn(vt_tag))
        value_elm.text = vt_text

    def delete_value(self, name: str) -> bool:
        """Remove the named custom property; return True if a property was removed.

        Returns False when no `<property>` child matches `name`. ``pid`` values of the
        remaining properties are not renumbered — ECMA-376 only requires that each
        ``pid`` be unique within the part, not that they form a contiguous sequence,
        and PowerPoint matches properties by ``name``.
        """
        prop = self._find_property(name)
        if prop is None:
            return False
        self.remove(prop)
        return True

    # -- internal helpers -------------------------------------------------

    @property
    def _property_elms(self) -> list[_Element]:
        """List of `<property>` child elements in document order."""
        return self.findall(qn("cst:property"))

    def _find_property(self, name: str) -> _Element | None:
        """Return the `<property>` child with `name=name`, or None if absent."""
        for prop in self._property_elms:
            if prop.get("name") == name:
                return prop
        return None

    def _next_pid(self) -> int:
        """Return the smallest unused pid >= MIN_PID across existing properties.

        Per ECMA-376 §15.2.12.2 pids must be unique within the part and at least 2.
        This method scans existing children and returns max(pid) + 1 (with a floor of
        MIN_PID), which is what Office and LibreOffice do in practice. It does not
        attempt to reuse holes in the sequence — round-tripping a file with holes
        (e.g. after a delete) keeps the existing pids stable and appends new ones
        above the current maximum.
        """
        max_pid = MIN_PID - 1
        for prop in self._property_elms:
            try:
                pid = int(prop.get("pid", ""))
            except ValueError:
                continue
            if pid > max_pid:
                max_pid = pid
        return max(MIN_PID, max_pid + 1)


# -- value <-> XML conversion helpers (module-level so they can be unit-tested in
# -- isolation without constructing a full CT_CustomProperties) -------------


_INT32_MIN = -(2**31)
_INT32_MAX = 2**31 - 1


def _vt_tag_and_text_for(value: CustomPropValue) -> tuple[str, str]:
    """Return the `(vt_tag, text)` pair used to serialize `value`.

    Note the ``bool`` branch must precede the ``int`` branch because ``bool`` is a
    subclass of ``int`` in Python — an unguarded ``isinstance(x, int)`` would route
    ``True`` and ``False`` to ``vt:i4`` instead of ``vt:bool``. Similarly the
    ``datetime`` branch must precede the ``date`` branch because ``datetime``
    subclasses ``date``.
    """
    if isinstance(value, bool):
        return "vt:bool", "true" if value else "false"
    if isinstance(value, int):
        if value < _INT32_MIN or value > _INT32_MAX:
            raise OverflowError(
                "custom-property int value %r out of 32-bit signed range; "
                "use a float or string instead" % (value,)
            )
        return "vt:i4", str(value)
    if isinstance(value, float):
        return "vt:r8", repr(value)
    if isinstance(value, dt.datetime):
        # -- serialize as ISO-8601 UTC ("...Z"). A naive datetime is treated as UTC,
        # -- mirroring how PowerPoint writes these. Aware datetimes are converted to
        # -- UTC first so the on-disk representation is unambiguous.
        if value.tzinfo is not None:
            value = value.astimezone(dt.timezone.utc).replace(tzinfo=None)
        return "vt:filetime", value.strftime("%Y-%m-%dT%H:%M:%SZ")
    if isinstance(value, dt.date):
        # -- a plain calendar date (no time, no zone) per ECMA-376 §22.4.2.7 --
        return "vt:date", value.strftime("%Y-%m-%d")
    if isinstance(value, str):  # pyright: ignore[reportUnnecessaryIsInstance]
        return "vt:lpwstr", value
    raise TypeError(
        "custom-property value must be str, int, float, bool, date, or datetime; got %s"
        % type(value).__name__
    )


def _safe_int(text: str) -> int | None:
    try:
        return int(text)
    except ValueError:
        return None


def _safe_float(text: str) -> float | None:
    try:
        return float(text)
    except ValueError:
        return None


def _parse_iso_datetime(text: str) -> dt.datetime | None:
    """Parse the ISO-8601 string `text` into a naive-UTC ``datetime``.

    Returns |None| if `text` is not parseable. PowerPoint writes ``filetime`` as
    ``YYYY-MM-DDTHH:MM:SSZ`` (with an optional fractional-seconds component), which
    :meth:`datetime.fromisoformat` handles on 3.11+ once the trailing ``Z`` is
    normalized to ``+00:00``.
    """
    if not text:
        return None
    normalized = text.strip()
    if normalized.endswith("Z"):
        normalized = normalized[:-1] + "+00:00"
    try:
        parsed = dt.datetime.fromisoformat(normalized)
    except ValueError:
        return None
    # -- normalize aware datetimes to naive-UTC for round-trip symmetry with setters --
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(dt.timezone.utc).replace(tzinfo=None)
    return parsed


def _parse_iso_date(text: str) -> dt.date | None:
    """Parse a ``vt:date`` ISO-8601 value (``YYYY-MM-DD``) into a ``date``.

    Returns |None| when `text` is empty or unparseable. Tolerates producers
    that mistakenly tack on a ``T...`` time component or a trailing ``Z`` /
    timezone designator — the spec forbids both, but unzipping real-world
    ``.pptx`` files occasionally turns them up and reading should not fail.
    The return type is a plain :class:`datetime.date` (never a `datetime`),
    which lets round-trips distinguish `vt:date` from `vt:filetime`.
    """
    if not text:
        return None
    s = text.strip()
    if s.endswith("Z"):
        s = s[:-1]
    if len(s) >= 6 and s[-6] in ("+", "-") and s[-3] == ":":
        s = s[:-6]
    if "T" in s:
        s = s.split("T", 1)[0]
    try:
        return dt.datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


# -- Per-tag decoders for vt: value elements. Maps Clark-name -> (text -> python value).
# -- The str / int / float families share a few tag names each; factoring the decoder
# -- table out as a module-level dict keeps `_python_value_of` a flat lookup instead of
# -- a twelve-branch `if` chain.
_VT_DECODERS: dict[str, object] = {
    # -- strings --
    qn("vt:lpwstr"): lambda text: text,
    qn("vt:lpstr"): lambda text: text,
    qn("vt:bstr"): lambda text: text,
    # -- integers (all representations collapse to python int on read) --
    qn("vt:i1"): _safe_int,
    qn("vt:i2"): _safe_int,
    qn("vt:i4"): _safe_int,
    qn("vt:i8"): _safe_int,
    qn("vt:int"): _safe_int,
    qn("vt:ui1"): _safe_int,
    qn("vt:ui2"): _safe_int,
    qn("vt:ui4"): _safe_int,
    qn("vt:ui8"): _safe_int,
    qn("vt:uint"): _safe_int,
    # -- floating-point / decimal --
    qn("vt:r4"): _safe_float,
    qn("vt:r8"): _safe_float,
    qn("vt:decimal"): _safe_float,
    # -- boolean (ECMA-376 accepts "true"/"false"/"1"/"0" per xsd:boolean) --
    qn("vt:bool"): lambda text: text.strip().lower() in ("true", "1"),
    # -- datetimes --
    qn("vt:filetime"): _parse_iso_datetime,
    # -- calendar dates (ECMA-376 §22.4.2.7); decoded to `date` so the python type
    # -- preserves the on-disk distinction between `vt:date` and `vt:filetime` on
    # -- round-trip.
    qn("vt:date"): _parse_iso_date,
}


def _python_value_of(prop: _Element) -> CustomPropValue | None:
    """Return the python value carried by the first vt: child of `prop`, or None."""
    for child in prop:
        decoder = _VT_DECODERS.get(child.tag)
        if decoder is None:
            continue
        return decoder(child.text or "")  # type: ignore[operator, no-any-return]
    return None
