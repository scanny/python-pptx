Feature: Access a series
  In order to operate on an individual chart series
  As a developer using python-pptx
  I need sequence operations on SeriesCollection


  Scenario Outline: SeriesCollection.__len__()
    Given a series collection for a <container> having <count> series
     Then len(series_collection) is <count>

    Examples: series container types
      | container         | count |
      | single-plot chart |   3   |
      | multi-plot chart  |   5   |
      | plot              |   3   |


  Scenario Outline: SeriesCollection.__getitem__()
    Given a series collection for a <container> having <count> series
     Then series_collection[2] is a Series object

    Examples: series container types
      | container         | count |
      | single-plot chart |   3   |
      | multi-plot chart  |   5   |
      | plot              |   3   |


  Scenario Outline: SeriesCollection.__iter__()
    Given a series collection for a <container> having <count> series
     Then iterating series_collection produces <count> Series objects

    Examples: series container types
      | container         | count |
      | single-plot chart |   3   |
      | multi-plot chart  |   5   |
      | plot              |   3   |
