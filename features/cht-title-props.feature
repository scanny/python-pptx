Feature: Get and set title properties
  In order to customize the formatting of a title on a chart
  As a developer using python-pptx
  I need a way to get and set title properties

  Scenario: Access the text frame of a title
    Given a title containing text
     Then title.text_frame contains the text in the title
