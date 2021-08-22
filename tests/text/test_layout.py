# encoding: utf-8

"""Unit-test suite for `pptx.text.layout` module."""

import pytest

from pptx.text.layout import _BinarySearchTree, _Line, _LineSource, TextFitter

from ..unitutil.mock import (
    ANY,
    call,
    class_mock,
    function_mock,
    initializer_mock,
    instance_mock,
    method_mock,
    property_mock,
)


class DescribeTextFitter(object):
    """Unit-test suite for `pptx.text.layout.TextFitter` object."""

    def it_can_determine_the_best_fit_font_size(self, request, line_source_):
        _LineSource_ = class_mock(
            request, "pptx.text.layout._LineSource", return_value=line_source_
        )
        _init_ = initializer_mock(request, TextFitter)
        _best_fit_font_size_ = method_mock(
            request, TextFitter, "_best_fit_font_size", return_value=36
        )
        extents, max_size = (19, 20), 42

        font_size = TextFitter.best_fit_font_size(
            "Foobar", extents, max_size, "foobar.ttf"
        )

        _LineSource_.assert_called_once_with("Foobar")
        _init_.assert_called_once_with(line_source_, extents, "foobar.ttf")
        _best_fit_font_size_.assert_called_once_with(ANY, max_size)
        assert font_size == 36

    def it_finds_best_fit_font_size_to_help_best_fit(self, _best_fit_fixture):
        text_fitter, max_size, _BinarySearchTree_ = _best_fit_fixture[:3]
        sizes_, predicate_, font_size_ = _best_fit_fixture[3:]

        font_size = text_fitter._best_fit_font_size(max_size)

        _BinarySearchTree_.from_ordered_sequence.assert_called_once_with(
            range(1, max_size + 1)
        )
        sizes_.find_max.assert_called_once_with(predicate_)
        assert font_size is font_size_

    @pytest.mark.parametrize(
        "extents, point_size, text_lines, expected_value",
        (
            ((66, 99), 6, ("foo", "bar"), False),
            ((66, 100), 6, ("foo", "bar"), True),
            ((66, 101), 6, ("foo", "bar"), True),
        ),
    )
    def it_provides_a_fits_inside_predicate_fn(
        self,
        request,
        line_source_,
        _rendered_size_,
        extents,
        point_size,
        text_lines,
        expected_value,
    ):
        _wrap_lines_ = method_mock(
            request, TextFitter, "_wrap_lines", return_value=text_lines
        )
        _rendered_size_.return_value = (None, 50)
        text_fitter = TextFitter(line_source_, extents, "foobar.ttf")

        predicate = text_fitter._fits_inside_predicate
        result = predicate(point_size)

        _wrap_lines_.assert_called_once_with(text_fitter, line_source_, point_size)
        _rendered_size_.assert_called_once_with(
            "Ty", point_size, text_fitter._font_file
        )
        assert result is expected_value

    def it_provides_a_fits_in_width_predicate_fn(self, fits_cx_pred_fixture):
        text_fitter, point_size, line = fits_cx_pred_fixture[:3]
        _rendered_size_, expected_value = fits_cx_pred_fixture[3:]

        predicate = text_fitter._fits_in_width_predicate(point_size)
        result = predicate(line)

        _rendered_size_.assert_called_once_with(
            line.text, point_size, text_fitter._font_file
        )
        assert result is expected_value

    def it_wraps_lines_to_help_best_fit(self, request):
        line_source, remainder = _LineSource("foo bar"), _LineSource("bar")
        _break_line_ = method_mock(
            request,
            TextFitter,
            "_break_line",
            side_effect=[("foo", remainder), ("bar", _LineSource(""))],
        )
        text_fitter = TextFitter(None, (None, None), None)

        text_fitter._wrap_lines(line_source, 21)

        assert _break_line_.call_args_list == [
            call(text_fitter, line_source, 21),
            call(text_fitter, remainder, 21),
        ]

    def it_breaks_off_a_line_to_help_wrap(
        self, request, line_source_, _BinarySearchTree_
    ):
        bst_ = instance_mock(request, _BinarySearchTree)
        _fits_in_width_predicate_ = method_mock(
            request, TextFitter, "_fits_in_width_predicate"
        )
        _BinarySearchTree_.from_ordered_sequence.return_value = bst_
        predicate_ = _fits_in_width_predicate_.return_value
        max_value_ = bst_.find_max.return_value
        text_fitter = TextFitter(None, (None, None), None)

        value = text_fitter._break_line(line_source_, 21)

        _BinarySearchTree_.from_ordered_sequence.assert_called_once_with(line_source_)
        text_fitter._fits_in_width_predicate.assert_called_once_with(text_fitter, 21)
        bst_.find_max.assert_called_once_with(predicate_)
        assert value is max_value_

    # fixtures ---------------------------------------------

    @pytest.fixture
    def _best_fit_fixture(self, _BinarySearchTree_, _fits_inside_predicate_):
        text_fitter = TextFitter(None, (None, None), None)
        max_size = 42
        sizes_ = _BinarySearchTree_.from_ordered_sequence.return_value
        predicate_ = _fits_inside_predicate_.return_value
        font_size_ = sizes_.find_max.return_value
        return (
            text_fitter,
            max_size,
            _BinarySearchTree_,
            sizes_,
            predicate_,
            font_size_,
        )

    @pytest.fixture(params=[(49, True), (50, True), (51, False)])
    def fits_cx_pred_fixture(self, request, _rendered_size_):
        rendered_width, expected_value = request.param
        text_fitter = TextFitter(None, (50, None), "foobar.ttf")
        point_size, line = 12, _Line("foobar", None)
        _rendered_size_.return_value = (rendered_width, None)
        return (text_fitter, point_size, line, _rendered_size_, expected_value)

    # fixture components -----------------------------------

    @pytest.fixture
    def _BinarySearchTree_(self, request):
        return class_mock(request, "pptx.text.layout._BinarySearchTree")

    @pytest.fixture
    def _fits_inside_predicate_(self, request):
        return property_mock(request, TextFitter, "_fits_inside_predicate")

    @pytest.fixture
    def line_source_(self, request):
        return instance_mock(request, _LineSource)

    @pytest.fixture
    def _rendered_size_(self, request):
        return function_mock(request, "pptx.text.layout._rendered_size")


class Describe_BinarySearchTree(object):
    """Unit-test suite for `pptx.text.layout._BinarySearchTree` object."""

    def it_can_construct_from_an_ordered_sequence(self):
        bst = _BinarySearchTree.from_ordered_sequence(range(10))

        def in_order(node):
            """
            Traverse the tree depth first to produce a list of its values,
            in order.
            """
            result = []
            if node is None:
                return result
            result.extend(in_order(node._lesser))
            result.append(node.value)
            result.extend(in_order(node._greater))
            return result

        assert bst.value == 9
        assert bst._lesser.value == 4
        assert bst._greater is None
        assert in_order(bst) == list(range(10))

    def it_can_find_the_max_value_satisfying_a_predicate(self, max_fixture):
        bst, predicate, expected_value = max_fixture
        assert bst.find_max(predicate) == expected_value

    # fixtures ---------------------------------------------

    @pytest.fixture(
        params=[
            (range(10), lambda n: n < 6.5, 6),
            (range(10), lambda n: n > 9.9, None),
            (range(10), lambda n: n < 0.0, None),
        ]
    )
    def max_fixture(self, request):
        seq, predicate, expected_value = request.param
        bst = _BinarySearchTree.from_ordered_sequence(seq)
        return bst, predicate, expected_value


class Describe_LineSource(object):
    """Unit-test suite for `pptx.text.layout._LineSource` object."""

    def it_generates_text_remainder_pairs(self):
        line_source = _LineSource("foo bar baz")
        expected = (
            ("foo", _LineSource("bar baz")),
            ("foo bar", _LineSource("baz")),
            ("foo bar baz", _LineSource("")),
        )
        assert all((a == b) for a, b in zip(expected, line_source))


# produces different results on Linux, fails Travis-CI

# from pptx.text.layout import _rendered_size
# from ..unitutil.file import testfile
# class Describe_rendered_size(object):

#     def it_calculates_the_rendered_size_of_text_at_point_size(self, fixture):
#         text, point_size, font_file, expected_value = fixture
#         extents = _rendered_size(text, point_size, font_file)
#         assert extents == expected_value

#     # fixtures ---------------------------------------------

#     @pytest.fixture(params=[
#         ('Typical',     18, (673100, 254000)),
#         ('foo bar baz', 12, (698500, 165100)),
#     ])
#     def fixture(self, request):
#         text, point_size, expected_value = request.param
#         font_file = testfile('calibriz.ttf')
#         return text, point_size, font_file, expected_value
