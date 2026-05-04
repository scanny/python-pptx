Feature: Change appearance of font used to render text
  In order to fine-tune the appearance of text
  As a developer using python-pptx
  I need a set of properties on the font used to render text


  Scenario Outline: Get Font.bold
    Given a font with bold set <bold-state>
     Then font.bold is <expected-value>

    Examples: font.bold states
      | bold-state | expected-value |
      | on         | True           |
      | off        | False          |
      | to inherit | None           |


  Scenario Outline: Set Font.bold
    Given a font with bold set <initial-state>
     When I assign <new-value> to font.bold
     Then font.bold is <new-value>

    Examples: Expected results of changing font.bold setting
      | initial-state | new-value |
      | on            | True      |
      | off           | True      |
      | to inherit    | True      |
      | on            | False     |
      | off           | False     |
      | to inherit    | False     |
      | on            | None      |
      | off           | None      |
      | to inherit    | None      |


  Scenario Outline: Get Font.italic
    Given a font with italic set <italic-state>
     Then font.italic is <expected-value>

    Examples: font.italic states
      | italic-state | expected-value |
      | on           | True           |
      | off          | False          |
      | to inherit   | None           |


  Scenario Outline: Set Font.italic
    Given a font with italic set <initial-state>
     When I assign <new-value> to font.italic
     Then font.italic is <new-value>

    Examples: Expected results of changing font.italic setting
      | initial-state | new-value |
      | on            | True      |
      | off           | True      |
      | to inherit    | True      |
      | on            | False     |
      | off           | False     |
      | to inherit    | False     |
      | on            | None      |
      | off           | None      |
      | to inherit    | None      |


  Scenario Outline: Get Font.language_id
    Given a font having language id <lang-id-state>
     Then font.language_id is MSO_LANGUAGE_ID.<member>

    Examples: font.language_id states
      | lang-id-state          | member |
      | of no explicit setting | NONE   |
      | MSO_LANGUAGE_ID.POLISH | POLISH |


  Scenario Outline: Set Font.language_id
    Given a font having language id <initial-state>
     When I assign <new-value> to font.language_id
     Then font.language_id is MSO_LANGUAGE_ID.<member>

    Examples: font.language_id assignment state changes
      | initial-state          | new-value              | member |
      | of no explicit setting | MSO_LANGUAGE_ID.FRENCH | FRENCH |
      | MSO_LANGUAGE_ID.FRENCH | MSO_LANGUAGE_ID.POLISH | POLISH |
      | MSO_LANGUAGE_ID.POLISH | MSO_LANGUAGE_ID.NONE   | NONE   |
      | MSO_LANGUAGE_ID.FRENCH | None                   | NONE   |


  Scenario Outline: Get Font.underline
    Given a font with underline set <underline-state>
     Then font.underline is <expected-value>

    Examples: font.underline states
      | underline-state | expected-value |
      | on              | True           |
      | off             | False          |
      | to inherit      | None           |
      | to DOUBLE_LINE  | DOUBLE_LINE    |
      | to WAVY_LINE    | WAVY_LINE      |


  Scenario Outline: Set Font.underline
    Given a font with underline set <initial-state>
     When I assign <new-value> to font.underline
     Then font.underline is <expected-value>

    Examples: Expected results of changing font.underline setting
      | initial-state  | new-value   | expected-value |
      | on             | True        | True           |
      | off            | SINGLE_LINE | True           |
      | to inherit     | True        | True           |
      | to WAVY_LINE   | False       | False          |
      | to inherit     | NONE        | False          |
      | to DOUBLE_LINE | None        | None           |
      | off            | DOUBLE_LINE | DOUBLE_LINE    |
      | to WAVY_LINE   | DOUBLE_LINE | DOUBLE_LINE    |


  Scenario Outline: Get Font.strikethrough
    Given a font with strikethrough set <strike-state>
     Then font.strikethrough is <expected-value>

    Examples: font.strikethrough states
      | strike-state    | expected-value |
      | on              | True           |
      | off             | False          |
      | to inherit      | None           |
      | to DOUBLE_LINE  | DOUBLE_LINE    |


  Scenario Outline: Set Font.strikethrough
    Given a font with strikethrough set <initial-state>
     When I assign <new-value> to font.strikethrough
     Then font.strikethrough is <expected-value>

    Examples: Expected results of changing font.strikethrough setting
      | initial-state  | new-value   | expected-value |
      | on             | True        | True           |
      | off            | SINGLE_LINE | True           |
      | to inherit     | True        | True           |
      | to DOUBLE_LINE | False       | False          |
      | to inherit     | NONE        | False          |
      | to DOUBLE_LINE | None        | None           |
      | off            | DOUBLE_LINE | DOUBLE_LINE    |
      | to DOUBLE_LINE | SINGLE_LINE | True           |


  Scenario Outline: Get Font.size
    Given a font having size of <value>
     Then font.size is <reported-size>

    Examples: Font sizes
      | value             | reported-size |
      | no explicit value | None          |
      | 42pt              | 42.0 points   |


  # -- effective font metrics (issue #378) -------------------------

  Scenario Outline: Get Font.effective_size walking the inheritance chain
    Given a font having size of <value>
     Then font.effective_size is <reported-size>

    Examples: resolved sizes
      | value             | reported-size |
      | no explicit value | 32.0 points   |
      | 42pt              | 42.0 points   |


  Scenario Outline: Get Font.effective_bold walking the inheritance chain
    Given a font with bold set <bold-state>
     Then font.effective_bold is <expected-value>

    Examples: resolved bold states
      | bold-state | expected-value |
      | on         | True           |
      | off        | False          |
      | to inherit | None           |


  Scenario Outline: Get Font.effective_italic walking the inheritance chain
    Given a font with italic set <italic-state>
     Then font.effective_italic is <expected-value>

    Examples: resolved italic states
      | italic-state | expected-value |
      | on           | True           |
      | off          | False          |
      | to inherit   | None           |


  Scenario: Get Font.effective_name walking the inheritance chain
    Given a font having size of no explicit value
     Then font.effective_name is the theme minor-latin placeholder


  Scenario: Set Font.name
    Given a font
     When I assign a typeface name to the font
     Then the font name matches the typeface I set


  Scenario: Set Font.name_ea for East-Asian text
    Given a font
     When I assign an East-Asian typeface name to the font
     Then the East-Asian font name matches the typeface I set
      And the Latin font name is unchanged


  Scenario: Set Font.name_cs for complex-script text
    Given a font
     When I assign a complex-script typeface name to the font
     Then the complex-script font name matches the typeface I set
      And the Latin font name is unchanged


  Scenario: Set all three font slots independently
    Given a font
     When I assign Latin, East-Asian, and complex-script typeface names to the font
     Then font.name, font.name_ea, and font.name_cs each match the value I set


  Scenario: Clear Font.name_ea and Font.name_cs by assigning None
    Given a font with Latin, East-Asian, and complex-script typefaces set
     When I assign None to font.name_ea
      And I assign None to font.name_cs
     Then font.name_ea is None
      And font.name_cs is None
      And the Latin font name is unchanged


  Scenario: Set a text shadow on a run (#546)
    Given a font
     When I assign 50800 to font.shadow.blur_radius
      And I assign 38100 to font.shadow.distance
      And I assign 45.0 to font.shadow.direction
     Then font.shadow.blur_radius is 50800
      And font.shadow.distance is 38100
      And font.shadow.direction is 45.0
      And font.shadow.inherit is False


  Scenario: Clear a text shadow by restoring inheritance (#546)
    Given a font
     When I assign 50800 to font.shadow.blur_radius
      And I assign True to font.shadow.inherit
     Then font.shadow.inherit is True
      And font.shadow.blur_radius is None


  Scenario: Set a text highlight color by RGB (#675)
    Given a font
     When I set font.highlight_color.rgb to FFFF00
      And I save and reload the font's presentation
     Then font.highlight_color.type is RGB
      And font.highlight_color.rgb is FFFF00


  Scenario: Set a text highlight color by theme color (#675)
    Given a font
     When I set font.highlight_color.theme_color to MSO_THEME_COLOR.ACCENT_1
      And I save and reload the font's presentation
     Then font.highlight_color.type is SCHEME
      And font.highlight_color.theme_color is MSO_THEME_COLOR.ACCENT_1


  Scenario: Clear a text highlight color (#675)
    Given a font with a yellow RGB highlight color
     When I call font.clear_highlight_color()
      And I save and reload the font's presentation
     Then the run has no a:highlight element


  Scenario: Add hyperlink
    Given a text run
     When I set the hyperlink address
     Then run.text is a hyperlink


  Scenario: Add hyperlink in table cell
    Given a text run in a table cell
     When I set the hyperlink address
     Then run.text is a hyperlink


  Scenario: Remove hyperlink
    Given a text run having a hyperlink
     When I assign None to hyperlink.address
     Then run.text is not a hyperlink


  Scenario: Set run hyperlink to jump to another slide in the deck
    Given a text run
      And another slide in the deck as target_slide
     When I assign target_slide to run.hyperlink.target_slide
     Then run.hyperlink.target_slide is target_slide
      And run.hyperlink.address is None


  Scenario: Clear a run slide-jump hyperlink by assigning None
    Given a text run with a slide-jump hyperlink to another slide
     When I assign None to run.hyperlink.target_slide
     Then run.hyperlink.target_slide is None
      And run.hyperlink.address is None


  Scenario: Setting a slide-jump hyperlink replaces an existing URL hyperlink
    Given a text run having a hyperlink
      And another slide in the deck as target_slide
     When I assign target_slide to run.hyperlink.target_slide
     Then run.hyperlink.target_slide is target_slide
      And run.hyperlink.address is None


  # -- Font.copy_from (issue #566) ------------------------------------

  Scenario: Copy every explicit property from one font to another
    Given a source run with bold, italic, size, color, and a Latin typeface
      And a destination run with no explicit formatting
     When I call dst_run.font.copy_from(src_run.font)
     Then dst_run.font.bold equals src_run.font.bold
      And dst_run.font.italic equals src_run.font.italic
      And dst_run.font.size equals src_run.font.size
      And dst_run.font.name equals src_run.font.name
      And dst_run.font.color.rgb equals src_run.font.color.rgb


  Scenario: copy_from clears destination properties absent on the source
    Given a source run with no explicit formatting
      And a destination run with bold, italic, size, and a Latin typeface
     When I call dst_run.font.copy_from(src_run.font)
     Then dst_run.font.bold is None
      And dst_run.font.italic is None
      And dst_run.font.size is None
      And dst_run.font.name is None


  Scenario: copy_from returns self to support chaining
    Given a source run with bold, italic, size, color, and a Latin typeface
      And a destination run with no explicit formatting
     When I call dst_run.font.copy_from(src_run.font) capturing the return
     Then the return value is the destination font


  # -- baseline / subscript / superscript (issue #1045) -----------------

  Scenario: Font.baseline is None by default and round-trips explicit values
    Given a font
     Then font.baseline is None
     When I assign 30000 to font.baseline
      And I save and reload the font's presentation
     Then font.baseline is 30000
      And font.superscript is True
      And font.subscript is False


  Scenario: Font.subscript = True writes PowerPoint's default -25% baseline
    Given a font
     When I assign True to font.subscript
      And I save and reload the font's presentation
     Then font.baseline is -25000
      And font.subscript is True
      And font.superscript is False


  Scenario: Font.superscript = True writes PowerPoint's default +30% baseline
    Given a font
     When I assign True to font.superscript
      And I save and reload the font's presentation
     Then font.baseline is 30000
      And font.superscript is True
      And font.subscript is False


  Scenario: Assigning None to Font.subscript clears the baseline attribute
    Given a font
     When I assign True to font.subscript
      And I assign None to font.subscript
     Then font.baseline is None
      And font.subscript is None
      And font.superscript is None


  Scenario: Assigning False to Font.superscript pins the baseline to zero
    Given a font
     When I assign True to font.superscript
      And I assign False to font.superscript
     Then font.baseline is 0
      And font.subscript is False
      And font.superscript is False
