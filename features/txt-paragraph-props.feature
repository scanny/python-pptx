Feature: Change paragraph properties
  In order to change the formatting of text to my needs
  As a developer using python-pptx
  I need a set of read/write properties on paragraph objects


  Scenario: Set paragraph alignment
     Given a paragraph
      When I set the paragraph alignment to centered
       And I reload the presentation
      Then the paragraph is aligned centered


  Scenario: Set paragraph indentation
     Given a paragraph
      When I indent the paragraph
       And I reload the presentation
      Then the paragraph is indented to the second level
