Feature: Add a picture to a slide
  In order to produce a presentation that includes images
  As a developer using python-pptx
  I need a way to place an image on a slide

  Scenario Outline: Add a picture to a slide
     Given a slide
      When I add the image <filename> using shapes.add_picture()
       And I save the presentation
      Then a <ext> image part appears in the pptx file
       And the picture appears in the slide

    Examples: image files of supported types
      | filename         | ext  |
      | sonic.gif        | gif  |
      | python-icon.jpeg | jpg  |
      | monty-truth.png  | png  |
      | 72-dpi.tiff      | tiff |
      | CVS_LOGO.WMF     | wmf  |
      | pic.emf          | wmf  |
      | python.bmp       | bmp  |


  Scenario Outline: Add a picture stream to a slide
     Given a slide
      When I add the stream image <filename> using shapes.add_picture()
       And I save the presentation
      Then a <ext> image part appears in the pptx file
       And the picture appears in the slide

    Examples: image files of supported types
      | filename         | ext  |
      | sonic.gif        | gif  |
      | python-icon.jpeg | jpg  |
      | monty-truth.png  | png  |
      | 72-dpi.tiff      | tiff |
      | CVS_LOGO.WMF     | wmf  |
      | pic.emf          | wmf  |
      | python.bmp       | bmp  |
