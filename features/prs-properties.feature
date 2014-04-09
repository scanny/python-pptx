Feature: Get and set presentation properties
  In order to discover and change behaviors of an overall presentation
  As a developer using python-pptx
  I need a set of read/write properties on the presentation object

  Scenario: Get slide dimensions
    Given a presentation
     Then its slide width matches its known value
      And its slide height matches its known value

  Scenario:
    Given a presentation
     When I change the slide width and height
     Then the slide width matches the new value
      And the slide height matches the new value
