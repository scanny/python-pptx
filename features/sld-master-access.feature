Feature: Access slide masters of a presentation
  In order to operate on a slide master in a presentation
  As a developer using python-pptx
  I need access to the slide master collection of a presentation

  Scenario: Access slide master collection
     Given a presentation having two slide masters
      Then I can access the slide master collection of the presentation
       And the length of the slide master collection is 2

  Scenario: Access slide master in slide master collection
     Given a slide master collection containing two masters
      Then I can iterate over the slide masters
       And I can access a slide master by index
