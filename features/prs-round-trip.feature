Feature: Save/reload round-trip fidelity for fork additions
  In order to trust that writes actually survive a save+reopen cycle
  As a developer using python-pptx
  I need end-to-end scenarios that assign properties then reload the package


  # -- Shape-level round-trips (BaseShape additions, Wave 12/13 area) --

  Scenario: BaseShape.flip_horizontal survives save and reload
    Given a fresh presentation with a flipped autoshape
     When I round-trip the presentation
     Then the reloaded autoshape flip_horizontal is True
      And the reloaded autoshape flip_vertical is True


  Scenario: BaseShape.alt_text and title survive save and reload
    Given a fresh presentation with a shape carrying alt_text and title
     When I round-trip the presentation
     Then the reloaded shape.alt_text is "description of chart"
      And the reloaded shape.title is "Quarterly Revenue"


  Scenario: BaseShape.is_hidden survives save and reload
    Given a fresh presentation with a hidden shape
     When I round-trip the presentation
     Then the reloaded shape.is_hidden is True


  Scenario: SlideShapes.get_by_name locates a named shape after reload
    Given a fresh presentation with a shape renamed to "LogoBanner"
     When I round-trip the presentation
     Then slide.shapes.get_by_name("LogoBanner") is the renamed shape
      And slide.shapes.get_by_name("DoesNotExist") is None


  Scenario: SlideShapes.find_all_by_name finds every match after reload
    Given a fresh presentation with three shapes all named "Tag"
     When I round-trip the presentation
     Then slide.shapes.find_all_by_name("Tag") returns 3 shapes
      And slide.shapes.find_all_by_name("Nope") returns 0 shapes


  Scenario: SlideShapes.iter_leaf_shapes yields only non-group descendants after reload
    Given a fresh presentation with a group of two shapes plus a lone rectangle
     When I round-trip the presentation
     Then the reloaded slide.shapes.iter_leaf_shapes yields 3 leaves


  Scenario: BaseShape.zorder helpers reorder shapes and round-trip
    Given a fresh presentation with three rectangles in z-order A B C
     When I call bring_to_front on A
      And I round-trip the presentation
     Then the reloaded shapes z-order is B C A


  # -- Text-frame round-trips --

  Scenario: TextFrame.rotation survives save and reload
    Given a fresh presentation with a text_frame rotated 45 degrees
     When I round-trip the presentation
     Then the reloaded text_frame.rotation == 45.0


  Scenario: TextFrame.font_scale and line_space_reduction survive save and reload
    Given a fresh presentation with autofit font_scale 75% and line_space_reduction 10%
     When I round-trip the presentation
     Then the reloaded text_frame.font_scale == 75.0
      And the reloaded text_frame.line_space_reduction == 10.0


  Scenario: TextFrame.replace_text survives save and reload
    Given a fresh presentation with a text_frame containing "Hello world" runs
     When I call text_frame.replace_text("world", "Python")
      And I round-trip the presentation
     Then the reloaded text_frame.text == "Hello Python"


  # -- Paragraph and run round-trips --

  Scenario: _Paragraph.delete removes the paragraph and round-trips
    Given a fresh presentation with a textbox of three paragraphs
     When I call paragraph.delete on the middle paragraph
      And I round-trip the presentation
     Then the reloaded textbox has 2 paragraphs
      And the reloaded paragraph texts are "first" and "third"


  Scenario: _Paragraph.delete on the last paragraph leaves an empty paragraph
    Given a fresh presentation with a textbox of one paragraph containing "only"
     When I call paragraph.delete on that paragraph
      And I round-trip the presentation
     Then the reloaded textbox has 1 paragraph
      And the reloaded first paragraph has no runs


  Scenario: _Run.delete removes a run and round-trips
    Given a fresh presentation with a paragraph of three runs "A" "B" "C"
     When I call run.delete on the middle run
      And I round-trip the presentation
     Then the reloaded paragraph has 2 runs
      And the reloaded run texts are "A" and "C"


  Scenario: _Paragraph.write_rich authors mixed-format runs that round-trip
    Given a fresh presentation with an empty paragraph
     When I call paragraph.write_rich with plain bold and sized parts
      And I round-trip the presentation
     Then the reloaded paragraph has 3 runs with the expected formatting


  # -- Font round-trips: baseline / subscript / shadow / effect_format --

  Scenario: Font.baseline survives save and reload
    Given a fresh presentation with a run whose font.baseline is 30000
     When I round-trip the presentation
     Then the reloaded font.baseline is 30000
      And the reloaded font.superscript is True
      And the reloaded font.subscript is False


  Scenario: Font.subscript survives save and reload
    Given a fresh presentation with a run whose font.subscript is True
     When I round-trip the presentation
     Then the reloaded font.subscript is True
      And the reloaded font.baseline is -25000


  Scenario: Font.shadow properties survive save and reload
    Given a fresh presentation with a run carrying a text shadow
     When I round-trip the presentation
     Then the reloaded font.shadow.blur_radius is 50800
      And the reloaded font.shadow.distance is 38100
      And the reloaded font.shadow.direction is 45.0


  Scenario: Font.effect_format.glow survives save and reload
    Given a fresh presentation with a run carrying a glow effect
     When I round-trip the presentation
     Then the reloaded font.effect_format.glow.size is 63500


  Scenario: Font.strikethrough survives save and reload
    Given a fresh presentation with a run whose font.strikethrough is SINGLE_LINE
     When I round-trip the presentation
     Then the reloaded font.strikethrough is MSO_STRIKE.SINGLE_LINE


  Scenario: Font.name_ea and name_cs survive save and reload
    Given a fresh presentation with a run whose Latin EA and CS fonts are set
     When I round-trip the presentation
     Then the reloaded font.name is "Calibri"
      And the reloaded font.name_ea is "MS Gothic"
      And the reloaded font.name_cs is "Arial"


  # -- Slide-level round-trips --

  Scenario: Slide.is_hidden survives save and reload
    Given a fresh presentation with a hidden slide
     When I round-trip the presentation
     Then the reloaded slide.is_hidden is True


  Scenario: Slide.show_master_shapes round-trips
    Given a fresh presentation with show_master_shapes False on slide 0
     When I round-trip the presentation
     Then the reloaded slide.show_master_shapes is False


  Scenario: Slide.name survives save and reload
    Given a fresh presentation with slide 0 named "Title slide"
     When I round-trip the presentation
     Then the reloaded slide.name is "Title slide"


  # -- Chart round-trip coverage --

  Scenario: Chart.display_blanks_as survives save and reload
    Given a fresh presentation with a chart whose display_blanks_as is GAPS
     When I round-trip the presentation
     Then the reloaded chart.display_blanks_as is XL_DISPLAY_BLANKS_AS.GAPS


  Scenario: CategoryAxis.tick_label_skip survives save and reload
    Given a fresh presentation with a chart whose tick_label_skip is 3
     When I round-trip the presentation
     Then the reloaded category_axis.tick_label_skip == 3
      And the reloaded category_axis.tick_mark_skip == 2


  Scenario: TickLabels.rotation survives save and reload
    Given a fresh presentation with a chart whose tick-label rotation is -45
     When I round-trip the presentation
     Then the reloaded value_axis.tick_labels.rotation == 315.0


  Scenario: Series error bars survive save and reload
    Given a fresh presentation with a chart whose first series has error bars
     When I round-trip the presentation
     Then the reloaded first series has error bars


  # -- Table cell/style round-trips --

  Scenario: _Cell.border_left survives save and reload
    Given a fresh presentation with a cell border_left red 1.5pt dashed
     When I round-trip the presentation
     Then the reloaded cell.border_left.color.rgb is FF0000
      And the reloaded cell.border_left.width.pt == 1.5


  Scenario: Table.style_id survives save and reload
    Given a fresh presentation with a table whose style_id is set
     When I round-trip the presentation
     Then the reloaded table.style_id matches the expected GUID


  Scenario: Table.add_row and add_column survive save and reload
    Given a fresh presentation with a 2x2 table
     When I call table.add_row and table.add_column
      And I round-trip the presentation
     Then the reloaded table is 3 rows by 3 columns


  # -- ActionSetting / hover round-trips --

  Scenario: ActionSetting.screen_tip survives save and reload
    Given a fresh presentation with a shape carrying a hover screen_tip
     When I round-trip the presentation
     Then the reloaded hover_action.screen_tip is "Hover me"


  Scenario: hover_action.hyperlink.address survives save and reload
    Given a fresh presentation with a shape carrying a hover hyperlink
     When I round-trip the presentation
     Then the reloaded hover_action.hyperlink.address is "https://example.com/hover"


  # -- Custom document-part props on shapes --

  Scenario: shape.custom_props survives save and reload for multiple keys
    Given a fresh presentation with a shape carrying three custom_props
     When I round-trip the presentation
     Then the reloaded shape.custom_props has 3 entries
      And the reloaded shape.custom_props gets k1 v1 k2 v2 k3 v3


  # -- Color / fill round-trips --

  Scenario: ColorFormat.alpha survives save and reload
    Given a fresh presentation with a solid-fill shape whose alpha is 0.5
     When I round-trip the presentation
     Then the reloaded fill.fore_color.alpha == 0.5


  # -- Picture / connector additions --

  Scenario: SlideShapes.add_picture_link survives save and reload
    Given a fresh presentation with a picture_link to an external URL
     When I round-trip the presentation
     Then the reloaded picture has an external rel to the URL


  # -- More feature coverage: trendline / chart helpers --

  Scenario: Series trendline survives save and reload
    Given a fresh presentation with a series carrying a LINEAR trendline
     When I round-trip the presentation
     Then the reloaded series has 1 trendline
      And the reloaded trendline type is LINEAR
      And the reloaded trendline displays the equation


  Scenario: Trendline.delete() removes the trendline and round-trips
    Given a fresh presentation with a series carrying a LINEAR trendline
     When I call trendline.delete on the first trendline
      And I round-trip the presentation
     Then the reloaded series has 0 trendlines


  Scenario: Chart.replace_data_preserve_formulas keeps embedded formulas
    Given a fresh presentation with a chart and a tweaked embedded workbook
     When I call chart.replace_data_preserve_formulas with new category data
      And I round-trip the presentation
     Then the reloaded chart has the replacement category values


  Scenario: Chart add_plot combo round-trips a second plot
    Given a fresh presentation with a combo chart carrying a LINE plot
     When I round-trip the presentation
     Then the reloaded chart has 2 plots


  # -- Line-format end formatting round-trip --

  Scenario: LineFormat.end_arrow.type round-trips on a connector
    Given a fresh presentation with a connector whose end_arrow.type is TRIANGLE
     When I round-trip the presentation
     Then the reloaded connector end_arrow.type is MSO_LINE_END_TYPE.TRIANGLE


  # -- Paragraph bullet round-trip --

  Scenario: _Paragraph.bullet.character round-trips
    Given a fresh presentation with a paragraph having a bullet character
     When I round-trip the presentation
     Then the reloaded paragraph bullet character is "•"


  Scenario: _Paragraph.bullet.auto_number round-trips
    Given a fresh presentation with an auto-numbered paragraph
     When I round-trip the presentation
     Then the reloaded paragraph uses auto-numbering


  # -- Sections round-trip --

  Scenario: Section creation round-trips and groups slides
    Given a fresh presentation with two sections covering three slides
     When I round-trip the presentation
     Then the reloaded presentation has 2 sections
      And the reloaded section names are "Intro" and "Main"


  # -- Slide tags round-trip --

  Scenario: Slide.tags round-trip two values
    Given a fresh presentation with a slide carrying two custom tags
     When I round-trip the presentation
     Then the reloaded slide has 2 tags
      And the reloaded slide tag "project" is "atlas"
      And the reloaded slide tag "status" is "draft"


  # -- Slide_layout re-point round-trip with shape preservation --

  Scenario: Slide.slide_layout setter re-points and shapes survive
    Given a fresh presentation with a slide on layout 1 carrying a textbox
     When I re-point the slide to layout 5
      And I round-trip the presentation
     Then the reloaded slide's layout index is 5
      And the reloaded slide still has the textbox


  # -- LineFormat dash style round-trip --

  Scenario: LineFormat.dash_style round-trips on a shape
    Given a fresh presentation with a shape having a dashed line
     When I round-trip the presentation
     Then the reloaded shape.line.dash_style is MSO_LINE.DASH


  # -- Notes slide round-trip --

  Scenario: Notes slide text survives save and reload
    Given a fresh presentation with notes "Speaker reminder" on slide 0
     When I round-trip the presentation
     Then the reloaded slide.notes_slide.notes_text_frame.text is "Speaker reminder"


  # -- Font highlight_color round-trip --

  Scenario: Font.highlight_color RGB survives save and reload
    Given a fresh presentation with a run carrying a yellow highlight
     When I round-trip the presentation
     Then the reloaded font.highlight_color.rgb is FFFF00


  # -- Picture cropping round-trip --

  Scenario: Picture crop properties survive save and reload
    Given a fresh presentation with a cropped picture
     When I round-trip the presentation
     Then the reloaded picture crop_left is 0.25
      And the reloaded picture crop_right is 0.25


  # -- Table merge round-trip --

  Scenario: Table cell merge survives save and reload
    Given a fresh presentation with a 3x3 table merging the top row
     When I round-trip the presentation
     Then the reloaded cell(0, 0) is_merge_origin is True
      And the reloaded cell(0, 0) span_width is 3


  # -- Slide Duplicate round-trip --

  Scenario: Slides.duplicate followed by round-trip preserves both slides
    Given a fresh presentation with a slide carrying a textbox "original"
     When I call slides.duplicate on the first slide
      And I round-trip the presentation
     Then the reloaded presentation has 2 slides
      And both reloaded slides have a textbox with text "original"


  # -- Animation read-only round-trip (#975) --

  Scenario: Set an appear animation survives save and reload
    Given a fresh presentation with a shape carrying an APPEAR entrance animation
     When I round-trip the presentation
     Then the reloaded slide.has_animations is True
