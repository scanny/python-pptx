Feature: Picture properties and methods
  In order to manipulate Picture shapes
  As a developer using python-pptx
  I need a set of properties and methods on the Picture object


  Scenario Outline: Picture.crop_x() getters
    Given a Picture object with <crop-or-no> as picture
     Then picture.crop_left == <l-crop>
      And picture.crop_top == <t-crop>
      And picture.crop_right == <r-crop>
      And picture.crop_bottom == <b-crop>

    Examples: Cropping getter cases
      | crop-or-no  | l-crop | t-crop | r-crop | b-crop |
      | no cropping | 0.0    | 0.0    | 0.0    | 0.0    |
      | cropping    | .15479 | .25571 | .10463 | .25572 |


  Scenario Outline: Picture.crop_x() setters
    Given a Picture object with <crop-or-no> as picture
     When I assign <value> to picture.crop_<side>
     Then picture.crop_<side> == <value>

    Examples: Cropping setter cases
      | crop-or-no  | value | side   |
      | no cropping | .1    | left   |
      | no cropping | .2    | top    |
      | no cropping | .3    | right  |
      | no cropping | .4    | bottom |
      | cropping    | .5    | left   |
      | cropping    | .6    | top    |
      | cropping    | .7    | right  |
      | cropping    | .8    | bottom |
      | no cropping | 0.0   | left   |
      | no cropping | 0.0   | top    |
      | no cropping | 0.0   | right  |
      | no cropping | 0.0   | bottom |
      | cropping    | 0.0   | left   |
      | cropping    | 0.0   | top    |
      | cropping    | 0.0   | right  |
      | cropping    | 0.0   | bottom |


  Scenario: Picture.image
    Given a Picture object as picture
     Then picture.image is an Image object


  Scenario: Picture.line
    Given a Picture object as picture
     Then picture.line is a LineFormat object
