Feature: Access a point
  In order to operate on an individual data point
  As a developer using python-pptx
  I need sequence operations on xPoints objects


  Scenario Outline: len(Points)
    Given a <type> object containing 3 points
     Then len(points) is 3

    Examples: point container types
      | type         |
      | XyPoints     |
      | BubblePoints |


  @wip
  Scenario Outline: access point by index
    Given a <type> object containing 3 points
     Then points[2] is a Point object

    Examples: point container types
      | type         |
      | XyPoints     |
      | BubblePoints |


  @wip
  Scenario Outline: iterate points
    Given a <type> object containing 3 points
     Then iterating points produces 3 Point objects

    Examples: point container types
      | type         |
      | XyPoints     |
      | BubblePoints |
