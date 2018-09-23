Feature: _Column properties and methods
  In order to discover and adjust the formatting of a table column
  As a developer using python-pptx
  I need properties and methods on _Column


  Scenario: _Column.width setter
     Given a _Column object as column
      When I assign column.width = Inches(1.5)
      Then column.width.inches == 1.5
