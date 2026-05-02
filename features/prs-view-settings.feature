Feature: Read and write presentation open-settings (view_props / first_slide_num)
  In order to control how PowerPoint opens a presentation
  As a developer using python-pptx
  I need access to the view-properties part and the firstSlideNum attribute

  Scenario: read view_props.view_type from the default template
     Given I open a freshly-created presentation
      Then view_props.view_type reports PP_VIEW_TYPE.SLIDE_THUMBNAIL

  Scenario: write view settings and round-trip through save + reopen
     Given I open a freshly-created presentation
      When I set view_type to OUTLINE and slide_view_zoom to 1.5
       And I set show_comments to False and first_slide_num to 7
       And I save the presentation
      Then opening the saved presentation shows view_type OUTLINE
       And opening the saved presentation shows slide_view_zoom 1.5
       And opening the saved presentation shows show_comments False
       And opening the saved presentation shows first_slide_num 7
