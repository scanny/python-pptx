"""Feature-manifest conformance tests against the corpus.

Iterates every ``pptx/*.json`` manifest in
``loadfix/ooxml-reference-corpus`` (expected as a sibling checkout at
``../ooxml-reference-corpus/``). For each manifest the suite:

1. Expands it into one or more concrete test cases (``kind: literal``
   is a single case; ``kind: parameterised`` is the Cartesian product
   of its `parameters` axes).
2. Invokes the committed generator to produce the case's fixture.
3. Asserts ``ooxml_validate.conformance.run_feature`` returns `pass`.

These tests are the live guard against drift between python-pptx and
the shared feature definitions. They auto-skip when either sibling
repo is absent.
"""

from __future__ import annotations

import importlib.util
import os
import runpy
import shlex
import subprocess
import sys
from pathlib import Path

import pytest

_REPO_ROOT = Path(__file__).resolve().parent.parent
_CORPUS_ROOT = _REPO_ROOT.parent / "ooxml-reference-corpus"
_FEATURES_DIR = _CORPUS_ROOT / "features" / "pptx"
_SCRIPTS_DIR = _CORPUS_ROOT / "scripts"
_FIXTURES_DIR = _CORPUS_ROOT / "fixtures" / "pptx"


def _ooxml_validate_available() -> bool:
    return importlib.util.find_spec("ooxml_validate") is not None


def _collect_cases() -> list[tuple[str, str, int]]:
    """Return a list of (manifest_id, case_id, case_index) tuples.

    For literal manifests the case_id equals the manifest_id.
    Parameterised manifests expand into one tuple per concrete case.
    ``case_index`` is the index into the expanded list so the test
    body can re-expand without mismatch.
    """
    if not _FEATURES_DIR.is_dir() or not _ooxml_validate_available():
        return []
    from ooxml_validate import expand_manifest, load_manifest

    cases: list[tuple[str, str, int]] = []
    for manifest_path in sorted(_FEATURES_DIR.glob("*.json")):
        manifest = load_manifest(manifest_path)
        for idx, case in enumerate(expand_manifest(manifest)):
            cases.append((manifest_path.stem, case["id"], idx))
    return cases


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


_COLLECTED_CASES = _collect_cases()


class DescribeCorpusConformance:
    @pytest.mark.parametrize(
        ("manifest_stem", "case_id", "case_index"),
        _COLLECTED_CASES or [("<no-corpus>", "<no-corpus>", 0)],
        ids=[case_id for _, case_id, _ in _COLLECTED_CASES] or ["<no-corpus>"],
    )
    def it_passes_every_pptx_case(
        self,
        manifest_stem: str,
        case_id: str,
        case_index: int,
        tmp_path: Path,
    ):
        if manifest_stem == "<no-corpus>":
            pytest.skip("No manifests found in corpus")

        from ooxml_validate import expand_manifest, load_manifest, run_feature
        from ooxml_validate.conformance import _substitute

        manifest_path = _FEATURES_DIR / f"{manifest_stem}.json"
        manifest = load_manifest(manifest_path)
        case = expand_manifest(manifest)[case_index]
        assert case["id"] == case_id, "case index drifted since collection"

        gen_script = _SCRIPTS_DIR / manifest["generator"]["python"].split("/")[-1]
        if not gen_script.is_file():
            pytest.skip(f"Generator script missing: {gen_script}")

        if manifest.get("kind") == "parameterised":
            _invoke_parameterised_generator(manifest, case, gen_script)
        else:
            _invoke_literal_generator(gen_script)

        fmt, logical_name = case["fixtures"]["machine"].split("/", 1)
        fixture = _CORPUS_ROOT / "fixtures" / fmt / f"{logical_name}.pptx"
        assert fixture.is_file(), f"Generator did not produce {fixture}"

        result = run_feature(
            case,
            library="python-pptx",
            fixture_path=fixture,
            tool_version=_current_version(),
        )
        assert result.status == "pass", [a.to_dict() for a in result.assertions]


def _invoke_literal_generator(gen_script: Path) -> None:
    """Run the generator as ``python <script>``; swallow a clean exit."""
    try:
        runpy.run_path(str(gen_script), run_name="__main__")
    except SystemExit as e:
        if e.code not in (0, None):
            raise


def _invoke_parameterised_generator(
    manifest: dict, case: dict, gen_script: Path
) -> None:
    """Run the generator with the manifest's arg_template rendered for this case.

    Parameterised generators accept CLI args. We substitute
    ``{axis.field}`` placeholders in ``generator.arg_template`` using
    the case's ``_expansion.bindings`` to find the right parameter
    record, then shell out.
    """
    from ooxml_validate.conformance import _substitute

    arg_template = manifest["generator"].get("arg_template")
    if not arg_template:
        pytest.skip(f"Parameterised manifest missing generator.arg_template")

    bindings = case["_expansion"]["bindings"]
    parameter_records = {
        axis: next(p for p in manifest["parameters"][axis] if p["id"] == bindings[axis])
        for axis in bindings
    }
    args_str = _substitute(arg_template, parameter_records)

    # Inherit PYTHONPATH so the subprocess can import `pptx` from the
    # editable src/ checkout just like the pytest process does.
    env = os.environ.copy()
    pythonpath = env.get("PYTHONPATH", "")
    src_path = str(_REPO_ROOT / "src")
    if src_path not in pythonpath.split(os.pathsep):
        env["PYTHONPATH"] = (
            src_path if not pythonpath else f"{src_path}{os.pathsep}{pythonpath}"
        )
    subprocess.run(
        [sys.executable, str(gen_script), *shlex.split(args_str)],
        cwd=_CORPUS_ROOT,
        check=True,
        capture_output=True,
        env=env,
    )


def _current_version() -> str:
    try:
        from pptx import __version__

        return str(__version__)
    except Exception:
        return "unknown"
