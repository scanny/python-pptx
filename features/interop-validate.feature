# These scenarios are optional integration tests against loadfix/ooxml-validate
# (a cross-format validator wrapping Microsoft Open XML SDK + LibreOffice).
#
# They skip at runtime (not via @skip/@wip tag) when any of the following are
# absent:
#   1. The PowerPoint-authored fixture at /mnt/data/Temp/microsoft/sample.pptx
#      (a developer-local reference deck, not shipped in the repo).
#   2. The `ooxml-validate` Python package (installed from a sibling checkout;
#      see the `[dev]` extra in pyproject.toml).
#   3. The `ooxml-validate` underlying tooling (OpenXmlValidator / LibreOffice
#      binaries resolved by the ooxml-validate package itself).
#
# The conditional skips are implemented in features/steps/interop_validate.py.
# Expect "2 skipped" in a normal CI-equivalent run; a "0 skipped" result
# means the full environment is present and the round-trip passed validation.

Feature: Round-trip fidelity via ooxml-validate
  Scenario: Round-tripping a real PowerPoint deck introduces no new validator issues
    Given a PowerPoint-authored pptx fixture
    When I round-trip it through python-pptx
    Then ooxml-validate reports zero new issues

  Scenario: The round-trip is visually identical under LibreOffice
    Given a PowerPoint-authored pptx fixture
    When I round-trip it through python-pptx
    Then the LibreOffice PDF render is visually identical
