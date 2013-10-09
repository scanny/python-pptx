Feature: Change properties of text in shapes
  In order to change the formatting of text to my needs
  As a developer using python-pptx
  I need to set the properties of text in a shape

  Scenario: Set paragraph alignment
     Given I have a reference to a paragraph
      When I set the paragraph alignment to centered
       And I save the presentation
      Then the paragraph is aligned centered

  Scenario: Set word wrap to True
	Given I have a reference to a textframe
	 When I set the textframe word wrap to True
	  And I save the presentation
	 Then the textframe word wrap is on

  Scenario: Set word wrap to False
	Given I have a reference to a textframe
	 When I set the textframe word wrap to False
	  And I save the presentation
	 Then the textframe word wrap is off

  Scenario: Set word wrap to empty
	Given I have a reference to a textframe
	 When I set the textframe word wrap to None
	  And I save the presentation
	 Then the textframe word wrap is empty
