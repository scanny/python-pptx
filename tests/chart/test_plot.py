# encoding: utf-8

"""
Test suite for pptx.chart.plot module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.plot import DataLabels, Plot

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class DescribePlot(object):

    def it_provides_access_to_the_data_labels(self, data_labels_fixture):
        plot, data_labels_, DataLabels_, dLbls = data_labels_fixture
        data_labels = plot.data_labels
        DataLabels_.assert_called_once_with(dLbls)
        assert data_labels is data_labels_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def data_labels_fixture(self, DataLabels_, data_labels_):
        barChart = element('c:barChart/c:dLbls')
        dLbls = barChart[0]
        plot = Plot(barChart)
        return plot, data_labels_, DataLabels_, dLbls

    # fixture components ---------------------------------------------

    @pytest.fixture
    def DataLabels_(self, request, data_labels_):
        return class_mock(
            request, 'pptx.chart.plot.DataLabels',
            return_value=data_labels_
        )

    @pytest.fixture
    def data_labels_(self, request):
        return instance_mock(request, DataLabels)


class DescribeDataLabels(object):

    def it_knows_its_number_format(self, number_format_get_fixture):
        data_labels, expected_value = number_format_get_fixture
        assert data_labels.number_format == expected_value

    def it_can_change_its_number_format(self, number_format_set_fixture):
        data_labels, new_value, expected_xml = number_format_set_fixture
        data_labels.number_format = new_value
        assert data_labels._element.xml == expected_xml

    def it_knows_whether_its_number_format_is_linked(
            self, number_format_is_linked_get_fixture):
        data_labels, expected_value = number_format_is_linked_get_fixture
        assert data_labels.number_format_is_linked is expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ('c:dLbls',                             'General'),
        ('c:dLbls/c:numFmt{formatCode=foobar}', 'foobar'),
    ])
    def number_format_get_fixture(self, request):
        dLbls_cxml, expected_value = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        return data_labels, expected_value

    @pytest.fixture(params=[
        ('c:dLbls', 'General', 'c:dLbls/c:numFmt{formatCode=General}'),
        ('c:dLbls/c:numFmt{formatCode=General}', '00.00',
         'c:dLbls/c:numFmt{formatCode=00.00}'),
    ])
    def number_format_set_fixture(self, request):
        dLbls_cxml, new_value, expected_dLbls_cxml = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        expected_xml = xml(expected_dLbls_cxml)
        return data_labels, new_value, expected_xml

    @pytest.fixture(params=[
        ('c:dLbls',                          True),
        ('c:dLbls/c:numFmt',                 True),
        ('c:dLbls/c:numFmt{sourceLinked=0}', False),
        ('c:dLbls/c:numFmt{sourceLinked=1}', True),
    ])
    def number_format_is_linked_get_fixture(self, request):
        dLbls_cxml, expected_value = request.param
        data_labels = DataLabels(element(dLbls_cxml))
        return data_labels, expected_value
