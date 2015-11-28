Feature: Get or set an external hyperlink on a shape or text run
  In order to connect a presentation to external resources
  As a developer using python-pptx
  I need CRUD operations on the Hyperlink object


  Scenario Outline: Get hyperlink URL
    Given a shape having click action <action>
     Then click_action.hyperlink.address is <value>

    Examples: Click actions
      | action             | value                        |
      | none               | None                         |
      | first slide        | None                         |
      | last slide viewed  | None                         |
      | named slide        | slide3.xml                   |
      | hyperlink          | http://yahoo.com             |
      | custom slide show  | None                         |
      | OLE action         | None                         |
      | run macro          | None                         |
      | run program        | /Applications/Calculator.app |


  Scenario Outline: Add hyperlink
    Given a shape having click action <action>
     When I assign <value> to click_action.hyperlink.address
     Then click_action.hyperlink.address is <value>

    Examples: Click actions
      | action      | value             |
      | none        | http://foobar.com |
      | first slide | http://barfoo.com |
      | named slide | http://barbaz.org |
      | hyperlink   | http://hooya.com  |
      | run program | http://booya.com  |


  Scenario Outline: Remove hyperlink
    Given a shape having click action <action>
     When I assign None to click_action.hyperlink.address
     Then click_action.hyperlink.address is None

    Examples: Click actions
      | action             |
      | none               |
      | first slide        |
      | last slide viewed  |
      | named slide        |
      | hyperlink          |
      | custom slide show  |
      | OLE action         |
      | run macro          |
      | run program        |

