Feature: Round-trip fidelity via ooxml-validate
  Scenario: Round-tripping a real PowerPoint deck introduces no new validator issues
    Given a PowerPoint-authored pptx fixture
    When I round-trip it through python-pptx
    Then ooxml-validate reports zero new issues

  Scenario: The round-trip is visually identical under LibreOffice
    Given a PowerPoint-authored pptx fixture
    When I round-trip it through python-pptx
    Then the LibreOffice PDF render is visually identical
