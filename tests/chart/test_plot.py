# encoding: utf-8

"""
Test suite for pptx.chart.plot module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.chart.plot import DataLabels, Plot

from ..unitutil.cxml import element
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
