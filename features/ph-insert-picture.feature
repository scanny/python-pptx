Feature: Insert a picture into a placeholder
  In order to add an image at a pre-defined location and size
  As a developer using python-pptx
  I need a way to insert an image into a placeholder


  Scenario Outline: Insert an image into a picture placeholder
     Given an unpopulated picture placeholder shape
      When I call placeholder.insert_picture('<filename>')
      Then the return value is a PlaceholderPicture object
       And the placeholder contains the image
       And the <sides> crop is <value>

    Examples: Images inserted into a picture placeholder
      | filename           | sides          | value   |
      | monty-truth.png    | top and bottom | 0.23715 |
      | python-powered.png | left and right | 0.23333 |
