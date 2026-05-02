Feature: Round-trip a presentation
  In order to satisfy myself that python-pptx might work
  As a pptx developer
  I want to see it pass a basic sanity-check

  Scenario: Round-trip a basic presentation
     Given a clean working directory
      When I open a basic PowerPoint presentation
       And I save the presentation
      Then I see the pptx file in the working directory

  Scenario: Start presentation from package stream
     Given a clean working directory
      When I open a presentation contained in a stream
       And I save the presentation
      Then I see the pptx file in the working directory

  Scenario: Start presentation from package extracted into directory
     Given a clean working directory
      When I open a presentation extracted into a directory
       And I save the presentation
      Then I see the pptx file in the working directory

  Scenario: Save presentation to package stream
     Given a clean working directory
      When I open a basic PowerPoint presentation
       And I save the presentation to a stream
       And I save that stream to a file
      Then I see the pptx file in the working directory

  Scenario: Round-trip external relationships
     Given a presentation with external relationships
      When I save and reload the presentation
      Then the external relationships are still there
       And the package has the expected number of .rels parts

  Scenario: Load presentation with invalid image/jpg MIME-type
     Given a presentation with an image/jpg MIME-type
      Then I can access the JPEG image

  Scenario: Save presentation with a fixed zip date-time produces byte-identical output
     Given a clean working directory
      When I open a basic PowerPoint presentation
       And I save it twice with zip_date_time fixed to 2020-01-01
      Then both saved streams are byte-for-byte identical
       And every zip member carries the 2020-01-01 00:00:00 last-modified stamp
  Scenario: Open a PowerPoint template (.potx) file
     Given a clean working directory
      When I open a PowerPoint template file
      Then the presentation is loaded
       And I see the pptx file in the working directory after saving

  Scenario: Round-trip a password-protected presentation
     Given a clean working directory
      When I save a password-protected presentation
       And I open the password-protected presentation with the correct password
      Then I see the pptx file in the working directory
       And the saved .pptx starts with the OLE2 magic signature

  Scenario: Opening a password-protected presentation without a password fails
     Given a password-protected presentation on disk
      Then opening it without a password raises EncryptedPackageError

  Scenario: Opening a password-protected presentation with the wrong password fails
     Given a password-protected presentation on disk
      Then opening it with the wrong password raises EncryptedPackageError

  Scenario: Round-trip a password-protected presentation with fixed zip_date_time
     Given a clean working directory
      When I save a password-protected presentation with a fixed zip date-time
       And I open the password-protected presentation with the correct password
      Then I see the pptx file in the working directory
       And the saved .pptx starts with the OLE2 magic signature
       And every decrypted zip member carries the 2020-01-01 00:00:00 last-modified stamp

  Scenario: Strip the slides from an existing deck to use it as a blank template
     Given an existing presentation with three slides and two sections
      When I call prs.strip_slides()
       And I save and reload the presentation
      Then the reloaded presentation has no slides
       And the reloaded presentation has no sections
       And the reloaded presentation retains its slide masters and layouts
       And I can add a new slide to the reloaded presentation
