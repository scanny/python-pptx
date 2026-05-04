"""Feature-manifest conformance tests against the corpus.

Iterates every ``pptx/*.json`` manifest in
``loadfix/ooxml-reference-corpus`` (expected as a sibling checkout at
``../ooxml-reference-corpus/``), regenerates the fixture via its
committed ``scripts/gen_<name>.py``, and asserts the result passes
``ooxml_validate.conformance.run_feature``.

These tests are the live guard against drift between python-pptx and
the shared feature definitions. They auto-skip when either sibling
repo is absent, so local checkouts without the corpus still succeed.
"""

from __future__ import annotations

import importlib.util
import runpy
from pathlib import Path

import pytest

_REPO_ROOT = Path(__file__).resolve().parent.parent
_CORPUS_ROOT = _REPO_ROOT.parent / "ooxml-reference-corpus"
_FEATURES_DIR = _CORPUS_ROOT / "features" / "pptx"
_SCRIPTS_DIR = _CORPUS_ROOT / "scripts"
_FIXTURES_DIR = _CORPUS_ROOT / "fixtures" / "pptx"


def _ooxml_validate_available() -> bool:
    return importlib.util.find_spec("ooxml_validate") is not None


def _manifest_ids() -> list[str]:
    if not _FEATURES_DIR.is_dir():
        return []
    return sorted(p.stem for p in _FEATURES_DIR.glob("*.json"))


# Skip at collection time so absent-corpus checkouts just don't see the tests.
pytestmark = [
    pytest.mark.skipif(
        not _CORPUS_ROOT.is_dir(),
        reason=f"Corpus sibling checkout not found at {_CORPUS_ROOT}",
    ),
    pytest.mark.skipif(
        not _ooxml_validate_available(),
        reason="ooxml-validate not installed in this env",
    ),
]


class DescribeCorpusConformance:
    @pytest.mark.parametrize("feature_id", _manifest_ids() or ["<no-corpus>"])
    def it_passes_every_pptx_manifest(self, feature_id: str, tmp_path: Path):
        if feature_id == "<no-corpus>":
            pytest.skip("No manifests found in corpus")

        from ooxml_validate import load_manifest, run_feature

        manifest_path = _FEATURES_DIR / f"{feature_id}.json"
        manifest = load_manifest(manifest_path)

        # Re-run the committed generator to produce a fresh fixture
        # with the current python-pptx source. We can't reuse the
        # committed fixture: the whole point is to check THIS source
        # still produces passing output.
        gen_script = _SCRIPTS_DIR / manifest["generator"]["python"].split("/")[-1]
        if not gen_script.is_file():
            pytest.skip(f"Generator script missing: {gen_script}")

        # The generator writes directly to corpus fixtures/ on success;
        # ignore the filesystem-side effect, just rely on its write.
        # Generators end with ``raise SystemExit(main())``, which propagates
        # even on success — catch the SystemExit(0) and re-raise on failure.
        try:
            runpy.run_path(str(gen_script), run_name="__main__")
        except SystemExit as e:
            if e.code not in (0, None):
                raise

        fixture = _FIXTURES_DIR / f"{feature_id}.pptx"
        assert fixture.is_file(), f"Generator did not produce {fixture}"

        result = run_feature(
            manifest,
            library="python-pptx",
            fixture_path=fixture,
            tool_version=_current_version(),
        )
        assert result.status == "pass", [a.to_dict() for a in result.assertions]


def _current_version() -> str:
    try:
        from pptx import __version__

        return str(__version__)
    except Exception:
        return "unknown"
