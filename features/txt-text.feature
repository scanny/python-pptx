Feature: Access and mutate text in a shape
  In order to discover and change the text in a presentation
  As a developer using python-pptx
  I need a way to access and mutate the text in a shape

  Scenario: Access the text of a run
    Given a run containing text
     Then run.text is the text in the run

  Scenario: Change the text of a run
    Given a run
     When I assign a string to run.text
     Then run.text matches the assigned string

  Scenario: Access the text of a paragraph
    Given a paragraph containing text
     Then paragraph.text is the text in the paragraph

  Scenario: Change the text of a paragraph
    Given a paragraph
     When I assign a string to paragraph.text
     Then paragraph.text matches the assigned string

  Scenario: Access the text of a shape
    Given a text frame containing text
     Then text_frame.text is the text in the shape

  Scenario: Change the text of a shape
    Given a text frame
     When I assign a string to text_frame.text
     Then text_frame.text matches the assigned string
