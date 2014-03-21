Feature: Access master placeholders
  In order to operate on the placeholder shapes of a slide master
  As a developer using python-pptx
  I need a placeholder collection on the slide master

  Scenario: Access placeholder collection of a slide master
     Given a slide master having two placeholders
      Then I can access the placeholder collection of the slide master
       And the length of the master placeholder collection is 2

  Scenario: Access placeholder in master placeholder collection
     Given a master placeholder collection
      Then I can iterate over the master placeholders
       And I can access a master placeholder by index
       And I can access a master placeholder by type
