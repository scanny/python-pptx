Feature: Get and set click action properties
  In order to determine the click action of a shape or text run
  As a developer using python-pptx
  I need a set of properties on ActionSetting

  Scenario Outline: get action
     Given a shape having click action <action>
      Then click_action.action is <value>

    Examples: Click actions
      | action             | value             |
      | none               | NONE              |
      | first slide        | FIRST_SLIDE       |
      | last slide         | LAST_SLIDE        |
      | previous slide     | PREVIOUS_SLIDE    |
      | next slide         | NEXT_SLIDE        |
      | last slide viewed  | LAST_SLIDE_VIEWED |
      | named slide        | NAMED_SLIDE       |
      | end show           | END_SHOW          |
      | hyperlink          | HYPERLINK         |
      | other presentation | PLAY              |
      | open file          | OPEN_FILE         |
      | custom slide show  | NAMED_SLIDE_SHOW  |
      | OLE action         | OLE_VERB          |
      | run macro          | RUN_MACRO         |
      | run program        | RUN_PROGRAM       |
