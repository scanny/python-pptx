Feature: Select a specific placeholder by idx
  In order to address a particular placeholder without iterating and matching
  As a developer using python-pptx
  I need an idx-based lookup on the slide's placeholder collection


  Scenario: Select a placeholder by idx with dictionary-style access
     Given a slide having placeholders at idx 0 and 10
      Then slide.placeholders[0] returns the title placeholder
       And slide.placeholders[10] returns the content placeholder


  Scenario: Select a placeholder by idx using .get()
     Given a slide having placeholders at idx 0 and 10
      Then slide.placeholders.get(10) returns the content placeholder
       And slide.placeholders.get(99) is None
       And a caller-supplied default is returned for an unknown idx


  Scenario: Dictionary-style access raises KeyError for unknown idx
     Given a slide having placeholders at idx 0 and 10
      Then slide.placeholders[99] raises KeyError
