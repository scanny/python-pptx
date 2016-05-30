Feature: Get and set click action properties
  In order to determine the click action of a shape or text run
  As a developer using python-pptx
  I need a set of properties on ActionSetting

  Scenario Outline: Get ActionSetting.action
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

  Scenario Outline: Get ActionSetting.hyperlink
     Given a shape having click action <action>
      Then click_action.hyperlink is a Hyperlink object

    Examples: Click actions
      | action             |
      | none               |
      | first slide        |
      | last slide         |
      | previous slide     |
      | next slide         |
      | last slide viewed  |
      | named slide        |
      | end show           |
      | hyperlink          |
      | other presentation |
      | open file          |
      | custom slide show  |
      | OLE action         |
      | run macro          |
      | run program        |

  Scenario Outline: Get ActionSetting.target slide
     Given a shape having click action <action>
      Then click_action.target_slide is slide <slide_idx>

    Examples: Click actions
      | action             | slide_idx |
      | none               |    None   |
      | first slide        |     0     |
      | last slide         |     4     |
      | previous slide     |     1     |
      | next slide         |     3     |
      | last slide viewed  |    None   |
      | named slide        |     2     |
      | end show           |    None   |
      | hyperlink          |    None   |
      | other presentation |    None   |
      | open file          |    None   |
      | custom slide show  |    None   |
      | OLE action         |    None   |
      | run macro          |    None   |
      | run program        |    None   |
