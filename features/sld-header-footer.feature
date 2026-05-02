Feature: Header / footer / slide-number / date placeholder support
  In order to turn slide-number, header, footer, and date placeholders on or
  off and to add an auto-refresh slide-number field to a text frame
  As a developer using python-pptx
  I need a header_footer object on the slide master and slide layout and an
  add_field method on a paragraph


  Scenario Outline: SlideMaster.header_footer flag defaults to visible
    Given a SlideMaster with no p:hf element
     Then slide_master.header_footer.<flag> is True

    Examples: header_footer flags
      | flag                  |
      | slide_number_visible  |
      | header_visible        |
      | footer_visible        |
      | date_visible          |


  Scenario Outline: SlideMaster.header_footer flags can be turned off
    Given a SlideMaster with no p:hf element
     When I set slide_master.header_footer.<flag> to False
     Then slide_master.header_footer.<flag> is False
      And the p:sldMaster element has a p:hf child

    Examples: header_footer flags
      | flag                  |
      | slide_number_visible  |
      | header_visible        |
      | footer_visible        |
      | date_visible          |


  Scenario: SlideLayout.header_footer flag defaults to visible
    Given a SlideLayout with no p:hf element
     Then slide_layout.header_footer.footer_visible is True


  Scenario: Paragraph.add_field appends a slide-number field
    Given a paragraph
     When I call paragraph.add_field("slidenum", "#")
     Then paragraph has one a:fld child
      And the a:fld type is "slidenum"
      And the a:fld text is "#"
