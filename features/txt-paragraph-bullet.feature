Feature: Change paragraph bullet formatting
  In order to control bullet formatting in my presentations
  As a developer using python-pptx
  I need a read/write API on paragraph objects for bullet settings


  Scenario: _Paragraph.bullet proxy
     Given a _Paragraph object as paragraph
      Then paragraph.bullet.type is None


  Scenario: Set a character bullet on a paragraph
     Given a _Paragraph object as paragraph
      When I call paragraph.bullet.character("-")
      Then paragraph.bullet.type == "char"
       And paragraph.bullet.char == "-"


  Scenario: Set an auto-number bullet on a paragraph
     Given a _Paragraph object as paragraph
      When I call paragraph.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)
      Then paragraph.bullet.type == "autonum"
       And paragraph.bullet.number_scheme == PP_AUTO_NUMBER.ARABIC_PERIOD
       And paragraph.bullet.start_at == 1


  Scenario: Set an auto-number bullet with custom start-at
     Given a _Paragraph object as paragraph
      When I call paragraph.bullet.auto_number(PP_AUTO_NUMBER.ROMAN_UC_PERIOD, 3)
      Then paragraph.bullet.type == "autonum"
       And paragraph.bullet.number_scheme == PP_AUTO_NUMBER.ROMAN_UC_PERIOD
       And paragraph.bullet.start_at == 3


  Scenario: Explicitly suppress a bullet
     Given a _Paragraph object as paragraph
      When I call paragraph.bullet.none()
      Then paragraph.bullet.type == "none"


  Scenario: Replace an existing bullet setting
     Given a _Paragraph object as paragraph
      When I call paragraph.bullet.character("-")
       And I call paragraph.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)
      Then paragraph.bullet.type == "autonum"
       And paragraph.bullet.char is None


  Scenario: Clear a bullet setting back to inherited
     Given a _Paragraph object as paragraph
      When I call paragraph.bullet.character("-")
       And I call paragraph.bullet.clear()
      Then paragraph.bullet.type is None
