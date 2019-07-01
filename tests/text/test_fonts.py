# encoding: utf-8

"""
Test suite for pptx.text.fonts module
"""

from __future__ import absolute_import, print_function, unicode_literals

import io
import pytest

from struct import calcsize

from pptx.compat import BytesIO
from pptx.text.fonts import (
    _BaseTable,
    _Font,
    FontFiles,
    _HeadTable,
    _NameTable,
    _Stream,
    _TableFactory,
)

from ..unitutil.file import test_file_dir, testfile
from ..unitutil.mock import (
    call,
    class_mock,
    function_mock,
    initializer_mock,
    instance_mock,
    method_mock,
    open_mock,
    property_mock,
    var_mock,
)


class DescribeFontFiles(object):
    def it_can_find_a_system_font_file(self, find_fixture):
        family_name, is_bold, is_italic, expected_path = find_fixture
        path = FontFiles.find(family_name, is_bold, is_italic)
        assert path == expected_path

    def it_catalogs_the_system_fonts_to_help_find(self, installed_fixture):
        expected_call_args, expected_values = installed_fixture
        installed_fonts = FontFiles._installed_fonts()
        assert FontFiles._iter_font_files_in.call_args_list == (expected_call_args)
        assert installed_fonts == expected_values

    def it_generates_font_dirs_to_help_find(self, font_dirs_fixture):
        expected_values = font_dirs_fixture
        font_dirs = FontFiles._font_directories()
        assert font_dirs == expected_values

    def it_knows_os_x_font_dirs_to_help_find(self, osx_dirs_fixture):
        expected_dirs = osx_dirs_fixture
        font_dirs = FontFiles._os_x_font_directories()
        assert font_dirs == expected_dirs

    def it_knows_windows_font_dirs_to_help_find(self, win_dirs_fixture):
        expected_dirs = win_dirs_fixture
        font_dirs = FontFiles._windows_font_directories()
        assert font_dirs == expected_dirs

    def it_iterates_over_fonts_in_dir_to_help_find(self, iter_fixture):
        directory, _Font_, expected_calls, expected_paths = iter_fixture
        paths = list(FontFiles._iter_font_files_in(directory))

        print(directory)

        assert _Font_.open.call_args_list == expected_calls
        assert paths == expected_paths

    # fixtures ---------------------------------------------

    @pytest.fixture(
        params=[
            ("Foobar", False, False, "foobar.ttf"),
            ("Foobar", True, False, "foobarb.ttf"),
            ("Barfoo", False, True, "barfooi.ttf"),
        ]
    )
    def find_fixture(self, request, _installed_fonts_):
        family_name, is_bold, is_italic, expected_path = request.param
        return family_name, is_bold, is_italic, expected_path

    @pytest.fixture(params=[("darwin", ["a", "b"]), ("win32", ["c", "d"])])
    def font_dirs_fixture(
        self, request, _os_x_font_directories_, _windows_font_directories_
    ):
        platform, expected_dirs = request.param
        dirs_meth_mock = {
            "darwin": _os_x_font_directories_,
            "win32": _windows_font_directories_,
        }[platform]
        sys_ = var_mock(request, "pptx.text.fonts.sys")
        sys_.platform = platform
        dirs_meth_mock.return_value = expected_dirs
        return expected_dirs

    @pytest.fixture
    def installed_fixture(self, _iter_font_files_in_, _font_directories_):
        _font_directories_.return_value = ["d", "d_2"]
        _iter_font_files_in_.side_effect = [
            [(("A", True, False), "a.ttf")],
            [(("B", False, True), "b.ttf")],
        ]
        expected_call_args = [call("d"), call("d_2")]
        expected_values = {("A", True, False): "a.ttf", ("B", False, True): "b.ttf"}
        return expected_call_args, expected_values

    @pytest.fixture
    def iter_fixture(self, _Font_):
        directory = test_file_dir
        font_file_path = testfile("calibriz.ttf")
        font = _Font_.open.return_value.__enter__.return_value
        font.family_name, font.is_bold, font.is_italic = "Arial", True, True
        expected_calls = [call(font_file_path)]
        expected_paths = [(("Arial", True, True), font_file_path)]
        return directory, _Font_, expected_calls, expected_paths

    @pytest.fixture
    def osx_dirs_fixture(self, request):
        import os

        os_ = var_mock(request, "pptx.text.fonts.os")
        os_.path = os.path
        os_.environ = {"HOME": "/Users/fbar"}
        return [
            "/Library/Fonts",
            "/Network/Library/Fonts",
            "/System/Library/Fonts",
            "/Users/fbar/Library/Fonts",
            "/Users/fbar/.fonts",
        ]

    @pytest.fixture
    def win_dirs_fixture(self, request):
        return [r"C:\Windows\Fonts"]

    # fixture components -----------------------------------

    @pytest.fixture
    def _Font_(self, request):
        return class_mock(request, "pptx.text.fonts._Font")

    @pytest.fixture
    def _font_directories_(self, request):
        return method_mock(request, FontFiles, "_font_directories")

    @pytest.fixture
    def _installed_fonts_(self, request):
        _installed_fonts_ = method_mock(request, FontFiles, "_installed_fonts")
        _installed_fonts_.return_value = {
            ("Foobar", False, False): "foobar.ttf",
            ("Foobar", True, False): "foobarb.ttf",
            ("Barfoo", False, True): "barfooi.ttf",
        }
        return _installed_fonts_

    @pytest.fixture
    def _iter_font_files_in_(self, request):
        return method_mock(request, FontFiles, "_iter_font_files_in")

    @pytest.fixture
    def _os_x_font_directories_(self, request):
        return method_mock(request, FontFiles, "_os_x_font_directories")

    @pytest.fixture
    def _windows_font_directories_(self, request):
        return method_mock(request, FontFiles, "_windows_font_directories")


class Describe_Font(object):
    def it_can_construct_from_a_font_file_path(self, open_fixture):
        path, _Stream_, stream_ = open_fixture
        with _Font.open(path) as f:
            _Stream_.open.assert_called_once_with(path)
            assert isinstance(f, _Font)
        stream_.close.assert_called_once_with()

    def it_knows_its_family_name(self, family_fixture):
        font, expected_name = family_fixture
        family_name = font.family_name
        assert family_name == expected_name

    def it_knows_whether_it_is_bold(self, bold_fixture):
        font, expected_value = bold_fixture
        assert font.is_bold is expected_value

    def it_knows_whether_it_is_italic(self, italic_fixture):
        font, expected_value = italic_fixture
        assert font.is_italic is expected_value

    def it_provides_access_to_its_tables(self, tables_fixture):
        font, _TableFactory_, expected_calls, expected_tables = tables_fixture
        tables = font._tables
        assert _TableFactory_.call_args_list == expected_calls
        assert tables == expected_tables

    def it_generates_table_records_to_help_read_tables(self, iter_fixture):
        font, expected_values = iter_fixture
        values = list(font._iter_table_records())
        assert values == expected_values

    def it_knows_the_table_count_to_help_read(self, table_count_fixture):
        font, expected_value = table_count_fixture
        assert font._table_count == expected_value

    def it_reads_the_header_to_help_read_font(self, fields_fixture):
        font, expected_values = fields_fixture
        fields = font._fields
        font._stream.read_fields.assert_called_once_with(">4sHHHH", 0)
        assert fields == expected_values

    # fixtures ---------------------------------------------

    @pytest.fixture(
        params=[("head", True, True), ("head", False, False), ("foob", True, False)]
    )
    def bold_fixture(self, request, _tables_, head_table_):
        key, is_bold, expected_value = request.param
        head_table_.is_bold = is_bold
        _tables_.return_value = {key: head_table_}
        font = _Font(None)
        return font, expected_value

    @pytest.fixture
    def family_fixture(self, _tables_, name_table_):
        font = _Font(None)
        expected_name = "Foobar"
        _tables_.return_value = {"name": name_table_}
        name_table_.family_name = expected_name
        return font, expected_name

    @pytest.fixture
    def fields_fixture(self, read_fields_):
        stream = _Stream(None)
        font = _Font(stream)
        read_fields_.return_value = expected_values = ("foob", 42, 64, 7, 16)
        return font, expected_values

    @pytest.fixture(
        params=[("head", True, True), ("head", False, False), ("foob", True, False)]
    )
    def italic_fixture(self, request, _tables_, head_table_):
        key, is_italic, expected_value = request.param
        head_table_.is_italic = is_italic
        _tables_.return_value = {key: head_table_}
        font = _Font(None)
        return font, expected_value

    @pytest.fixture
    def iter_fixture(self, _table_count_, stream_read_):
        stream = _Stream(None)
        font = _Font(stream)
        _table_count_.return_value = 2
        stream_read_.return_value = (
            b"name"
            b"xxxx"
            b"\x00\x00\x00\x2A"
            b"\x00\x00\x00\x15"
            b"head"
            b"xxxx"
            b"\x00\x00\x00\x15"
            b"\x00\x00\x00\x2A"
        )
        expected_values = [("name", 42, 21), ("head", 21, 42)]
        return font, expected_values

    @pytest.fixture
    def open_fixture(self, _Stream_):
        path = "foobar.ttf"
        stream_ = _Stream_.open.return_value
        return path, _Stream_, stream_

    @pytest.fixture
    def tables_fixture(
        self, stream_, name_table_, head_table_, _iter_table_records_, _TableFactory_
    ):
        font = _Font(stream_)
        _iter_table_records_.return_value = iter([("name", 11, 22), ("head", 33, 44)])
        _TableFactory_.side_effect = [name_table_, head_table_]

        expected_calls = [call("name", stream_, 11, 22), call("head", stream_, 33, 44)]
        expected_tables = {"name": name_table_, "head": head_table_}
        return font, _TableFactory_, expected_calls, expected_tables

    @pytest.fixture
    def table_count_fixture(self, _fields_):
        font = _Font(None)
        _fields_.return_value = (-666, 42)
        expected_value = 42
        return font, expected_value

    # fixture components -----------------------------------

    @pytest.fixture
    def _fields_(self, request):
        return property_mock(request, _Font, "_fields")

    @pytest.fixture
    def head_table_(self, request):
        return instance_mock(request, _HeadTable)

    @pytest.fixture
    def _iter_table_records_(self, request):
        return method_mock(request, _Font, "_iter_table_records")

    @pytest.fixture
    def name_table_(self, request):
        return instance_mock(request, _NameTable)

    @pytest.fixture
    def read_fields_(self, request):
        return method_mock(request, _Stream, "read_fields")

    @pytest.fixture
    def _Stream_(self, request):
        return class_mock(request, "pptx.text.fonts._Stream")

    @pytest.fixture
    def stream_(self, request):
        return instance_mock(request, _Stream)

    @pytest.fixture
    def stream_read_(self, request):
        return method_mock(request, _Stream, "read")

    @pytest.fixture
    def _TableFactory_(self, request):
        return function_mock(request, "pptx.text.fonts._TableFactory")

    @pytest.fixture
    def _table_count_(self, request):
        return property_mock(request, _Font, "_table_count")

    @pytest.fixture
    def _tables_(self, request):
        return property_mock(request, _Font, "_tables")


class Describe_Stream(object):
    def it_can_construct_from_a_path(self, open_fixture):
        path, open_, _init_, file_ = open_fixture
        stream = _Stream.open(path)
        open_.assert_called_once_with(path, "rb")
        _init_.assert_called_once_with(file_)
        assert isinstance(stream, _Stream)

    def it_can_be_closed(self, close_fixture):
        stream, file_ = close_fixture
        stream.close()
        file_.close.assert_called_once_with()

    def it_can_read_fields_from_a_template(self, read_flds_fixture):
        stream, tmpl, offset, file_, expected_values = read_flds_fixture
        fields = stream.read_fields(tmpl, offset)
        file_.seek.assert_called_once_with(offset)
        file_.read.assert_called_once_with(calcsize(tmpl))
        assert fields == expected_values

    def it_can_read_bytes(self, read_fixture):
        stream, offset, length, file_, expected_value = read_fixture
        bytes_ = stream.read(offset, length)
        file_.seek.assert_called_once_with(offset)
        file_.read.assert_called_once_with(length)
        assert bytes_ == expected_value

    # fixtures ---------------------------------------------

    @pytest.fixture
    def close_fixture(self, file_):
        stream = _Stream(file_)
        return stream, file_

    @pytest.fixture
    def open_fixture(self, open_, _init_):
        path = "foobar.ttf"
        file_ = open_.return_value
        return path, open_, _init_, file_

    @pytest.fixture
    def read_fixture(self, file_):
        stream = _Stream(file_)
        offset, length = 42, 21
        file_.read.return_value = "foobar"
        expected_value = "foobar"
        return stream, offset, length, file_, expected_value

    @pytest.fixture
    def read_flds_fixture(self, file_):
        stream = _Stream(file_)
        tmpl, offset = b">4sHH", 0
        file_.read.return_value = b"foob" b"\x00\x2A" b"\x00\x15"
        expected_values = (b"foob", 42, 21)
        return stream, tmpl, offset, file_, expected_values

    # fixture components -----------------------------------

    @pytest.fixture
    def file_(self, request):
        return instance_mock(request, io.RawIOBase)

    @pytest.fixture
    def _init_(self, request):
        return initializer_mock(request, _Stream)

    @pytest.fixture
    def open_(self, request):
        return open_mock(request, "pptx.text.fonts")


class Describe_TableFactory(object):
    def it_constructs_the_appropriate_table_object(self, fixture):
        tag, stream_, offset, length, TableClass_, TableClass = fixture
        table = _TableFactory(tag, stream_, offset, length)
        TableClass_.assert_called_once_with(tag, stream_, offset, length)
        assert isinstance(table, TableClass)

    # fixtures ---------------------------------------------

    @pytest.fixture(params=["name", "head", "foob"])
    def fixture(self, request, stream_):
        tag = request.param
        offset, length = 42, 21
        TableClass, target = {
            "name": (_NameTable, "pptx.text.fonts._NameTable"),
            "head": (_HeadTable, "pptx.text.fonts._HeadTable"),
            "foob": (_BaseTable, "pptx.text.fonts._BaseTable"),
        }[tag]
        TableClass_ = class_mock(request, target)
        return tag, stream_, offset, length, TableClass_, TableClass

    # fixture components -----------------------------------

    @pytest.fixture
    def stream_(self, request):
        return instance_mock(request, _Stream)


class Describe_HeadTable(object):
    def it_knows_whether_the_font_is_bold(self, bold_fixture):
        head_table, expected_value = bold_fixture
        assert head_table.is_bold is expected_value

    def it_knows_whether_the_font_is_italic(self, italic_fixture):
        head_table, expected_value = italic_fixture
        assert head_table.is_italic is expected_value

    def it_reads_its_macStyle_field_to_help(self, macStyle_fixture):
        head_table, expected_value = macStyle_fixture
        assert head_table._macStyle == expected_value

    # fixtures ---------------------------------------------

    @pytest.fixture(params=[(0, False), (1, True)])
    def bold_fixture(self, request, _macStyle_):
        macStyle, expected_value = request.param
        _macStyle_.return_value = macStyle
        head_table = _HeadTable(None, None, None, None)
        return head_table, expected_value

    @pytest.fixture(params=[(5, False), (7, True)])
    def italic_fixture(self, request, _macStyle_):
        macStyle, expected_value = request.param
        _macStyle_.return_value = macStyle
        head_table = _HeadTable(None, None, None, None)
        return head_table, expected_value

    @pytest.fixture
    def macStyle_fixture(self):
        bytes_ = b"xxxxyyyy....................................\xF0\xBA........"
        stream = _Stream(BytesIO(bytes_))
        offset, length = 0, len(bytes_)
        head_table = _HeadTable(None, stream, offset, length)
        expected_value = 61626
        return head_table, expected_value

    # fixture components -----------------------------------

    @pytest.fixture
    def _macStyle_(self, request):
        return property_mock(request, _HeadTable, "_macStyle")


class Describe_NameTable(object):
    def it_knows_the_font_family_name(self, family_fixture):
        name_table, expected_value = family_fixture
        family_name = name_table.family_name
        assert family_name == expected_value

    def it_provides_access_to_its_names_to_help_props(self, names_fixture):
        name_table, names_dict = names_fixture
        names = name_table._names
        name_table._iter_names.assert_called_once_with()
        assert names == names_dict

    def it_iterates_over_its_names_to_help_read_names(self, iter_fixture):
        name_table, expected_calls, expected_names = iter_fixture
        names = list(name_table._iter_names())
        assert name_table._read_name.call_args_list == expected_calls
        assert names == expected_names

    def it_reads_the_table_header_to_help_read_names(self, header_fixture):
        names_table, expected_value = header_fixture
        header = names_table._table_header
        assert header == expected_value

    def it_buffers_the_table_bytes_to_help_read_names(self, bytes_fixture):
        name_table, expected_value = bytes_fixture
        table_bytes = name_table._table_bytes
        name_table._stream.read.assert_called_once_with(
            name_table._offset, name_table._length
        )
        assert table_bytes == expected_value

    def it_reads_a_name_to_help_read_names(self, read_fixture):
        name_table, bufr, idx, strs_offset, platform_id = read_fixture[:5]
        encoding_id, name_str_offset, length = read_fixture[5:8]
        expected_value = read_fixture[8]

        name = name_table._read_name(bufr, idx, strs_offset)

        name_table._name_header.assert_called_once_with(bufr, idx)
        name_table._read_name_text.assert_called_once_with(
            bufr, platform_id, encoding_id, strs_offset, name_str_offset, length
        )
        assert name == expected_value

    def it_reads_a_name_header_to_help_read_names(self, name_hdr_fixture):
        name_table, bufr, idx, expected_value = name_hdr_fixture
        header = name_table._name_header(bufr, idx)
        assert header == expected_value

    def it_reads_name_text_to_help_read_names(self, name_text_fixture):
        name_table, bufr, platform_id, encoding_id = name_text_fixture[:4]
        strings_offset, name_str_offset, length = name_text_fixture[4:7]
        raw_name, name_ = name_text_fixture[7:]

        name = name_table._read_name_text(
            bufr, platform_id, encoding_id, strings_offset, name_str_offset, length
        )

        name_table._raw_name_string.assert_called_once_with(
            bufr, strings_offset, name_str_offset, length
        )
        name_table._decode_name.assert_called_once_with(
            raw_name, platform_id, encoding_id
        )
        assert name is name_

    def it_reads_name_bytes_to_help_read_names(self, raw_fixture):
        name_table, bufr, strings_offset, str_offset = raw_fixture[:4]
        length, expected_bytes = raw_fixture[4:]
        bytes_ = name_table._raw_name_string(bufr, strings_offset, str_offset, length)
        assert bytes_ == expected_bytes

    def it_decodes_a_raw_name_to_help_read_names(self, decode_fixture):
        name_table, raw_name, platform_id, encoding_id, expected_value = decode_fixture
        name = name_table._decode_name(raw_name, platform_id, encoding_id)
        assert name == expected_value

    # fixtures ---------------------------------------------

    @pytest.fixture
    def bytes_fixture(self, stream_):
        name_table = _NameTable(None, stream_, 42, 360)
        bytes_ = "\x00\x01\x02\x03\x04\x05"
        stream_.read.return_value = bytes_
        expected_value = bytes_
        return name_table, expected_value

    @pytest.fixture(
        params=[
            (1, 0, b"Foob\x8Ar", "Foobär"),
            (1, 1, b"Foobar", None),
            (0, 9, "Foobär".encode("utf-16-be"), "Foobär"),
            (3, 6, "Foobär".encode("utf-16-be"), "Foobär"),
            (2, 0, "Foobar", None),
        ]
    )
    def decode_fixture(self, request):
        platform_id, encoding_id, raw_name, expected_value = request.param
        name_table = _NameTable(None, None, None, None)
        return name_table, raw_name, platform_id, encoding_id, expected_value

    @pytest.fixture(
        params=[
            ({(0, 1): "Foobar", (1, 1): "Barfoo"}, "Foobar"),
            ({(1, 1): "Barfoo", (3, 1): "Farbaz"}, "Barfoo"),
            ({(3, 1): "Farbaz", (6, 2): "BazFoo"}, "Farbaz"),
            ({(9, 1): "Foobar", (6, 1): "Barfoo"}, None),
        ]
    )
    def family_fixture(self, request, _names_):
        names, expected_value = request.param
        name_table = _NameTable(None, None, None, None)
        _names_.return_value = names
        return name_table, expected_value

    @pytest.fixture
    def header_fixture(self, _table_bytes_):
        name_table = _NameTable(None, None, None, None)
        _table_bytes_.return_value = b"\x00\x00\x00\x02\x00\x2A"
        expected_value = (0, 2, 42)
        return name_table, expected_value

    @pytest.fixture
    def iter_fixture(self, _table_header_, _table_bytes_, _read_name):
        name_table = _NameTable(None, None, None, None)
        _table_header_.return_value = (0, 3, 42)
        _table_bytes_.return_value = "xXx"
        _read_name.side_effect = [(0, 1, "Foobar"), (3, 1, "Barfoo"), (9, 9, None)]
        expected_calls = [call("xXx", 0, 42), call("xXx", 1, 42), call("xXx", 2, 42)]
        expected_names = [((0, 1), "Foobar"), ((3, 1), "Barfoo")]
        return name_table, expected_calls, expected_names

    @pytest.fixture
    def name_hdr_fixture(self):
        name_table = _NameTable(None, None, None, None)
        bufr = (
            b"123456"
            b"123456789012"
            b"\x00\x00"
            b"\x00\x01"
            b"\x00\x02"
            b"\x00\x03"
            b"\x00\x04"
            b"\x00\x05"
        )
        idx = 1
        expected_value = (0, 1, 2, 3, 4, 5)
        return name_table, bufr, idx, expected_value

    @pytest.fixture
    def names_fixture(self, _iter_names_):
        name_table = _NameTable(None, None, None, None)
        _iter_names_.return_value = iter([((0, 1), "Foobar"), ((3, 1), "Barfoo")])
        names_dict = {(0, 1): "Foobar", (3, 1): "Barfoo"}
        return name_table, names_dict

    @pytest.fixture
    def name_text_fixture(self, _raw_name_string_, _decode_name_):
        name_table = _NameTable(None, None, None, None)
        bufr, platform_id, encoding_id, strings_offset = "xXx", 6, 7, 8
        name_str_offset, length, raw_name = 9, 10, "Foobar"
        _raw_name_string_.return_value = raw_name
        name_ = _decode_name_.return_value
        return (
            name_table,
            bufr,
            platform_id,
            encoding_id,
            strings_offset,
            name_str_offset,
            length,
            raw_name,
            name_,
        )

    @pytest.fixture
    def raw_fixture(self):
        name_table = _NameTable(None, None, None, None)
        bufr = b"xXxFoobarxXx"
        strings_offset, str_offset, length = 1, 2, 6
        expected_bytes = b"Foobar"
        return (name_table, bufr, strings_offset, str_offset, length, expected_bytes)

    @pytest.fixture
    def read_fixture(self, _name_header, _read_name_text):
        name_table = _NameTable(None, None, None, None)
        bufr, idx, strs_offset, platform_id, name_id = "buffer", 3, 47, 0, 1
        encoding_id, name_str_offset, length, name = 7, 36, 12, "Arial"
        _name_header.return_value = (
            platform_id,
            encoding_id,
            666,
            name_id,
            length,
            name_str_offset,
        )
        _read_name_text.return_value = name
        expected_value = (platform_id, name_id, name)
        return (
            name_table,
            bufr,
            idx,
            strs_offset,
            platform_id,
            encoding_id,
            name_str_offset,
            length,
            expected_value,
        )

    # fixture components -----------------------------------

    @pytest.fixture
    def _decode_name_(self, request):
        return method_mock(request, _NameTable, "_decode_name")

    @pytest.fixture
    def _iter_names_(self, request):
        return method_mock(request, _NameTable, "_iter_names")

    @pytest.fixture
    def _name_header(self, request):
        return method_mock(request, _NameTable, "_name_header")

    @pytest.fixture
    def _raw_name_string_(self, request):
        return method_mock(request, _NameTable, "_raw_name_string")

    @pytest.fixture
    def _read_name(self, request):
        return method_mock(request, _NameTable, "_read_name")

    @pytest.fixture
    def _read_name_text(self, request):
        return method_mock(request, _NameTable, "_read_name_text")

    @pytest.fixture
    def stream_(self, request):
        return instance_mock(request, _Stream)

    @pytest.fixture
    def _table_bytes_(self, request):
        return property_mock(request, _NameTable, "_table_bytes")

    @pytest.fixture
    def _table_header_(self, request):
        return property_mock(request, _NameTable, "_table_header")

    @pytest.fixture
    def _names_(self, request):
        return property_mock(request, _NameTable, "_names")
