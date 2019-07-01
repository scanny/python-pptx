Feature: Shape properties and methods
  In order to identify and adjust autoshapes
  As a developer using python-pptx
  I need properties and methods on Shape


  Scenario: Shape.adjustments setter
     Given a chevron shape
      When I assign 0.15 to shape.adjustments[0]
      Then shape.adjustments[0] is 0.15


  Scenario: Shape.line
     Given a Shape object as shape
      Then shape.line is a LineFormat object


  Scenario: Shape.text getter
     Given a Shape object having text as shape
      Then shape.text == "Fee Fi\vF\xf8\xf8 Fum\nI am a shape\vwith textium"


  Scenario: Shape.text setter
     Given a Shape object having text as shape
      When I assign shape.text = "F\xf8o\vBar\nBaz\x1b"
      Then shape.text == "F\xf8o\vBar\nBaz_x001B_"
