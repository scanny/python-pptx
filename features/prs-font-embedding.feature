Feature: Embed custom fonts in a presentation
  In order to preserve my document's appearance on machines that lack my
  branded typefaces
  As a developer using python-pptx
  I need to embed TrueType / OpenType font files in the presentation


  Scenario: Embed a regular-style font file
    Given a presentation
     When I embed a font file under the typeface "Pacifico"
     Then Presentation.embedded_fonts contains "Pacifico"
      And the presentation has a ppt/fonts/font1.fntdata part after save


  Scenario: Embed bold and regular styles under a single typeface
    Given a presentation
     When I embed a regular and a bold font file under the typeface "Pacifico"
     Then Presentation.embedded_fonts contains "Pacifico"
      And the p:embeddedFont entry has both p:regular and p:bold child elements
