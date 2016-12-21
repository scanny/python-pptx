.. _chart-data-api:

ChartData objects
=================

A ChartData object is used to specify the data depicted in a chart. It is
used when creating a new chart and when replacing the data for an existing
chart. Most charts are created using a |CategoryChartData| object, however
the data for XY and bubble chart types is different enough that those each
require a distinct chart data object.

.. autoclass:: pptx.chart.data.ChartData
   :members:
   :member-order: bysource
   :undoc-members:


.. autoclass:: pptx.chart.data.CategoryChartData
   :members:
   :member-order: bysource
   :inherited-members:
   :exclude-members:
       append, categories_ref, count, data_point_offset, index, series_index,
       series_name_ref, x_values_ref, xlsx_blob, xml_bytes, y_values_ref
   :undoc-members:

.. autoclass:: pptx.chart.data.Categories
   :members:
   :member-order: bysource
   :inherited-members:
   :exclude-members: count
   :undoc-members:

.. autoclass:: pptx.chart.data.Category
   :members:
   :member-order: bysource
   :inherited-members:
   :exclude-members: depth, idx, index, leaf_count
   :undoc-members:

.. autoclass:: pptx.chart.data.XyChartData
   :members:
   :member-order: bysource
   :inherited-members:
   :exclude-members:
       count, data_point_offset, index, series_index, series_name_ref,
       x_values_ref, xlsx_blob, xml_bytes, y_values_ref

.. autoclass:: pptx.chart.data.BubbleChartData
   :members:
   :member-order: bysource
   :inherited-members:
   :exclude-members:
       count, bubble_sizes_ref, data_point_offset, index, series_index,
       series_name_ref, x_values_ref, xlsx_blob, xml_bytes, y_values_ref

.. autoclass:: pptx.chart.data.XySeriesData
   :members:
   :member-order: bysource
   :inherited-members:
   :exclude-members:
       count, data_point_offset, name_ref, x_values_ref, y_values_ref

.. autoclass:: pptx.chart.data.BubbleSeriesData
   :members:
   :member-order: bysource
   :inherited-members:
   :exclude-members:
       count, bubble_sizes_ref, name_ref, x_values_ref, y_values_ref

