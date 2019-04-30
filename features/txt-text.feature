Feature: Access and mutate text in a shape
  In order to discover and change the text in a presentation
  As a developer using python-pptx
  I need a way to access and mutate the text in a shape


  Scenario: _Run.text getter
    Given a _Run object containing text as run
     Then run.text == " Foo Bar "


  Scenario Outline: _Run.text setter
    Given a run
     When I assign run.text = <value>
     Then run.text == <expected-value>

    Examples: _Paragraph assigned text replacement cases
      | value     | expected-value |
      | "abc"     | "abc"          |
      | "a\tb\nc" | "a\tb\nc"      |
