Feature: OMML (Office Math Markup Language) Support
  As a presentation developer
  I want to add mathematical equations to slides using OMML
  So that I can include complex mathematical formulas in PowerPoint presentations

  Scenario: Add basic math equation to slide
    Given a presentation with one slide
    When I add a math equation "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:r><m:t>x = 2</m:t></m:r></m:oMath>"
    Then the slide should contain one math shape
    And the math shape should contain the equation "x = 2"

  Scenario: Add fraction equation to slide
    Given a presentation with one slide
    When I add a math equation with fraction "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>"
    Then the slide should contain one math shape
    And the math shape should contain a fraction with numerator "x" and denominator "2"

  Scenario: Add superscript equation to slide
    Given a presentation with one slide
    When I add a math equation with superscript "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:r><m:t>x</m:t></m:r><m:sSup><m:e><m:r><m:t>2</m:t></m:r></m:e></m:sSup></m:oMath>"
    Then the slide should contain one math shape
    And the math shape should contain "x²"

  Scenario: Add radical equation to slide
    Given a presentation with one slide
    When I add a math equation with radical "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:rad><m:radPr/><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>"
    Then the slide should contain one math shape
    And the math shape should contain a square root of "x"

  Scenario: Add n-ary operator equation to slide
    Given a presentation with one slide
    When I add a math equation with summation "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:nary><m:naryPr><m:chr val='∑'/><m:limLoc val='undOvr'/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary></m:oMath>"
    Then the slide should contain one math shape
    And the math shape should contain a summation from i=1 to n of i

  Scenario: Position and size math equation
    Given a presentation with one slide
    When I add a math equation "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:r><m:t>E = mc²</m:t></m:r></m:oMath>"
    And I position the math shape at left 100000, top 100000
    And I size the math shape to width 200000, height 100000
    Then the math shape left should be 100000
    And the math shape top should be 100000
    And the math shape width should be 200000
    And the math shape height should be 100000

  Scenario: Add multiple math equations to slide
    Given a presentation with one slide
    When I add a first math equation "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:r><m:t>a² + b²</m:t></m:r></m:oMath>"
    And I add a second math equation "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:r><m:t>= c²</m:t></m:r></m:oMath>"
    Then the slide should contain two math shapes
    And the first math shape should contain "a² + b²"
    And the second math shape should contain "= c²"

  Scenario: Get OMML XML from existing math equation
    Given a presentation with one slide containing a math equation
    When I get the OMML XML from the math shape
    Then the OMML XML should be valid math markup
    And the OMML XML should contain the original equation content

  Scenario: Replace OMML content in existing math equation
    Given a presentation with one slide containing a math equation
    When I replace the OMML content with "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:r><m:t>y = mx + b</m:t></m:r></m:oMath>"
    Then the math shape should contain the equation "y = mx + b"

  Scenario: Math shape properties
    Given a presentation with one slide containing a math shape
    Then the math shape should have default properties
    And the math shape should support rotation
    And the math shape should support shadow effects
    And the math shape should be included in slide shapes collection

  Scenario: Save and reload presentation with math equations
    Given a presentation with math equations
    When I save the presentation to a file
    And I reload the presentation from the file
    Then the presentation should contain the same math equations
    And the math equations should have the same content
    And the math equations should have the same positions and sizes
