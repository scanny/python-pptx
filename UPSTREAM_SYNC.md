# Upstream sync catalogue — python-pptx

Governed by ADR 004 in the
[ooxml-reference-corpus](https://github.com/loadfix/ooxml-reference-corpus/blob/master/docs/adr/004-upstream-sync.md)
repo. See that ADR for cadence, tooling, and disposition vocabulary.

## Tracking

- **Upstream project.** `scanny/python-pptx`
  (`https://github.com/scanny/python-pptx.git`). Effectively dormant
  since August 2024.
- **Fork baseline SHA.** `278b47b1dedd5b46ee84c286e77cdfb0bf4594be`
  (upstream `master` HEAD). Subject:
  `fix(enum): replace read-only enum values`. Dated 2024-08-07.
  This is the upstream `master` HEAD at fork time and, as of
  2026-05-05, it is still the upstream `master` HEAD.
- **Fork divergent commit count.** 785 (fork additions since
  `278b47b1`, as of 2026-05-05).
- **Upstream remote is pre-configured.** A default clone already has:
  ```bash
  git remote add upstream https://github.com/scanny/python-pptx.git
  ```
  To reproduce this sweep:
  ```bash
  git fetch upstream
  git log --no-merges --oneline 278b47b1..upstream/master
  ```

## Sync status (sweep: 2026-05-05)

As of this sweep `upstream/master` is **at the baseline SHA** — no
new commits have landed upstream since the fork began. There is
nothing to evaluate and nothing to pull.

| short-sha  | date       | subject                               | tier   | disposition | rationale |
|------------|------------|---------------------------------------|--------|-------------|-----------|
| *(none)*   | —          | —                                     | —      | —           | `upstream/master` == fork baseline SHA |

Upstream non-master branches observed (not pulled):

- `upstream/develop` — present but at or behind `upstream/master`;
  no new commits since baseline.
- `upstream/rc` — release candidate branch, no movement since baseline.
- `upstream/hotfix-iss36` — ancient (2013) stale branch.
  Disposition: `ignore`.
- `upstream/dependabot/pip/jinja2-3.1.4` — dependency-bump for the
  *docs build*, not the library runtime. Tier: compatibility.
  Disposition: `ignore` (fork docs build uses its own pins, see
  the "stale doc-build pins" project note).

## Consequence for loadfix/python-pptx

Because upstream is dormant, loadfix is effectively the primary
maintainer. Fork-only bug fixes should not expect to be mirrored
upstream. If `scanny/python-pptx` resumes activity, re-read ADR 004
and consider whether to offer selected loadfix commits back upstream;
that decision is out of scope for this catalogue.

## Next sweep due

**2026-08-03** (first Monday of August 2026).

## History

- 2026-05-05 — initial catalogue created by Wave 11-D. Baseline
  confirmed; zero upstream divergence.
