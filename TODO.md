# python-pptx — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Open

_None tracked yet._

## Done

- Issue #627 — author tables and charts (plus textbox / picture / auto-shape /
  connector) directly inside a `GroupShape` via `group.shapes.add_table(...)`
  etc. `add_table` promoted from `SlideShapes` onto the shared
  `_BaseGroupShapes` base so every subclass inherits it. Branch
  `feat/issue-627-group-shapes-authoring`.
