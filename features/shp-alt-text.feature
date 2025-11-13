Feature: Shape alternative text
  In order to provide accessibility features
  As a developer using python-pptx
  I need to get and set alternative text for shapes

  Scenario: Get alt_text from shape with no alternative text
    Given a shape with no alternative text
     Then the alt_text property should be None

  Scenario: Get alt_text from shape with existing alternative text
    Given a shape with alternative text "A red rectangle"
     Then the alt_text property should be "A red rectangle"

  Scenario: Set alt_text on a shape
    Given a shape with no alternative text
     When I set the alt_text property to "Accessibility description"
     Then the alt_text property should be "Accessibility description"

  Scenario: Update alt_text on a shape
    Given a shape with alternative text "Old description"
     When I set the alt_text property to "New description"
     Then the alt_text property should be "New description"

  Scenario: Clear alt_text from a shape
    Given a shape with alternative text "Some description"
     When I set the alt_text property to an empty string
     Then the alt_text property should be an empty string

  Scenario: Delete alt_text from a shape
    Given a shape with alternative text "Some description"
     When I delete the alt_text property
     Then the alt_text property should be None