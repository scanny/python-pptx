
Slide-id manager
================

Foundation item *F6* — the collection of rules, helpers, and API surface that
keep every `p:sldId/@id` attribute in ``presentation.xml`` unique and stable
across insert, reorder, delete, and duplicate operations on the
|Slides| collection.

Unlike Foundations F1–F5, F6 did not ship as a single reusable class. The
headline features it was meant to enable — reorder (`#68`), delete (`#67`),
duplicate (`#132`), and cross-presentation append (`#1036`) — landed on
:class:`pptx.slide.Slides` as individual methods during Waves 1–3, each
reusing the pre-existing id allocator on :class:`CT_SlideIdList`. This
document consolidates the contract so the rules are easy to find when
future features (move-slide-to-other-presentation, slide-level undo, etc.)
need to interact with the same machinery.


The invariants
--------------

F6 guarantees the following about the `p:sldIdLst` child of
``presentation.xml``:

1. **Uniqueness.** Every child `p:sldId` element carries a distinct
   ``@id`` attribute in the range ``[256, 2_147_483_647]`` (the XSD range
   for ``ST_SlideId``). No two slides in a presentation ever share a slide
   id while both are present in the list.

2. **Stability under reorder.** Moving a slide (:meth:`Slides.move_slide`)
   changes its position in `p:sldIdLst` but never changes its ``@id``.
   External artifacts that reference the slide by id (for example, a
   hyperlink with `action="ppaction://hlinksldjump"`) keep resolving.

3. **Stability under delete.** Removing a slide (:meth:`Slides.delete`)
   drops its `p:sldId` entry and decrements the relationship-count on its
   r:id, but does not renumber any surviving slide. A deleted slide's
   id *may* be recycled by a later insert only after the last
   higher-numbered slide is also gone — see the allocator below.

4. **Freshness for inserts.** Every insertion path
   (:meth:`Slides.add_slide`, :meth:`Slides.duplicate`,
   :meth:`Slides.add_slide_from_external`) allocates a *new* id that is
   guaranteed not to collide with any id currently in use. The allocator
   is :meth:`pptx.oxml.presentation.CT_SlideIdList.add_sldId`.


Allocation algorithm
--------------------

All five mutating entry points on ``Slides`` ultimately call
``CT_SlideIdList.add_sldId(rId)``, which delegates to the
``CT_SlideIdList._next_id`` property. That single method is the one and
only place where a slide-id is chosen; there is no per-call-site
allocation.

The algorithm (in :file:`src/pptx/oxml/presentation.py`):

1. Read every existing ``p:sldId/@id`` as an ``int``.
2. Compute ``simple_next = max(255, *used_ids) + 1``. If this is at or
   below ``MAX_SLIDE_ID`` (2 147 483 647), return it.
3. Otherwise — the presentation contains a slide with the maximum legal
   id — scan upward from 256 and return the first integer that is not in
   the valid-used set. (Invalid out-of-range ids are ignored so a
   malformed file cannot make the allocator run off the end.)

Step 2 is the common case. It gives the "id equals highest-ever-used + 1"
behavior PowerPoint produces and means the id of a deleted slide is
*not* recycled while any higher-numbered slide survives — a property
downstream consumers depend on.

Step 3 is the fallback that fixed `#972` (``_next_id`` failing when the
max was already in use). The unit tests in
:file:`tests/oxml/test_presentation.py` exercise both branches.


API surface
-----------

All methods are on :class:`pptx.slide.Slides`. F6 does not expose any
allocator directly to user code; the contract is that any
slide-producing method returns a slide with a fresh, unique id, and any
reordering / removal method preserves the invariants above.

.. list-table::
   :header-rows: 1
   :widths: 24 10 10 56

   * - Method
     - allocates id?
     - reorders?
     - F6 role
   * - :meth:`add_slide(slide_layout)` — `#35`
     - yes
     - appends
     - Baseline entry point; pre-dates F6 and uses the allocator
       directly. Every subsequent feature reuses this allocator.
   * - :meth:`move_slide(slide, new_idx)` — `#68`
     - no
     - yes
     - Reorders a `p:sldId` without touching its `@id`. Delegates to
       ``_reposition_sldId``.
   * - :meth:`delete(slide)` — `#67`
     - no
     - no
     - Removes the `p:sldId` entry and drops the presentation-to-slide
       relationship. Deferred garbage-collection of uniquely-referenced
       child parts happens at ``save()`` time.
   * - :meth:`duplicate(slide, index=None)` — `#132`
     - yes (one)
     - optional
     - Calls the allocator (step 1), then optionally reorders via
       ``_reposition_sldId``. Guarantees the duplicate carries a fresh
       id distinct from both the source and every other slide.
   * - :meth:`add_slide_from_external(source_slide, slide_layout)` —
       `#1036`
     - yes
     - appends
     - Basic cross-presentation clone path. The id is allocated against
       *this* presentation's `p:sldIdLst`, never the source's, which
       means two presentations that both happen to use id 256 merge
       without collision.


The shared ``_reposition_sldId`` helper
---------------------------------------

:meth:`move_slide` and :meth:`duplicate(index=...)` both need the same
primitive: *move a `p:sldId` element that is already a child of
`p:sldIdLst` to a new zero-based position, matching Python list
semantics for negative and out-of-range indices.* The shared logic lives
in ``Slides._reposition_sldId``:

.. code-block:: python

    def _reposition_sldId(self, sldId: CT_SlideId, new_idx: int) -> None:
        n = len(self._sldIdLst)
        target_idx = max(0, new_idx + n) if new_idx < 0 else min(new_idx, n - 1)
        if self._sldIdLst.sldId_lst.index(sldId) == target_idx:
            return
        self._sldIdLst.remove(sldId)
        siblings = self._sldIdLst.sldId_lst
        if target_idx >= len(siblings):
            self._sldIdLst.append(sldId)
        else:
            siblings[target_idx].addprevious(sldId)

Key invariants:

* ``_reposition_sldId`` **does not allocate or modify `@id`**. Allocation
  is exclusively ``CT_SlideIdList.add_sldId``'s job; the two concerns
  are deliberately kept apart.

* Index normalization uses ``n = len(self._sldIdLst)`` where ``n``
  already counts the element being moved (because the caller in both
  call sites holds a live reference to a `p:sldId` that is already a
  child). This gives the same clamp / wrap semantics ``Slides.move_slide``
  shipped with in `#68`, which ``Slides.duplicate`` adopted in `#132`
  to avoid divergent behavior across the two APIs.


Interactions with other foundations
-----------------------------------

* **F1 (cross-part relationship cloner).** ``Slides.duplicate`` and
  ``Slides.add_slide_from_external`` are the two F6 entry points that
  rely on F1 for the rId-and-parts bookkeeping on the new slide part.
  F1 assigns rIds; F6 assigns slide-ids. Neither touches the other's
  identifier space.

* **F5 (embedded-workbook handler).** Not directly relevant to F6 — F5
  sits below F1 and never reads `p:sldIdLst`. Noted here because F5 is
  commonly triggered as a side-effect of F6 operations that move charts.

* **presentation.xml rels.** Every `p:sldId` carries an `r:id` that
  must resolve to a live relationship on the presentation part. F6
  operations preserve this invariant by always routing rId allocation
  through ``PresentationPart.add_slide`` / ``duplicate_slide`` /
  ``add_slide_from_external`` *before* adding the `p:sldId`, and by
  always removing the `p:sldId` *before* ``drop_rel`` in the delete
  path (so ``drop_rel``'s reference-count check sees zero).


Regression test
---------------

:file:`tests/test_slide.py` contains a
``DescribeSlides.it_allocates_a_unique_sldId_after_delete_move_and_insert``
test that exercises a move → delete → duplicate(index=0) →
add_slide_from_external sequence and asserts the resulting `@id`
values are all distinct and that the deleted slide's id is not
recycled while a higher-numbered slide survives. The existing
:file:`tests/oxml/test_presentation.py`::DescribeCT_SlideIdList tests
cover the allocator in isolation — including the `#972`
max-slide-id-in-use fallback.


Why no dedicated class?
-----------------------

During Wave 0 planning, F6 was scoped as a potential `SlideIdManager`
class collaborating with :class:`Slides`. Once `CT_SlideIdList.add_sldId`
was fixed to handle the max-id edge case (`#972`), the allocator itself
became a stable single-method primitive, and every F6 feature could
consume it directly with no additional abstraction. The only shared
logic on the ``Slides`` side turned out to be positional (sibling
reorder under Python list semantics), which is captured in the
``_reposition_sldId`` private helper.

Introducing a dedicated manager class now would add a layer without
eliminating a call site. If a future feature needs a more elaborate
policy (for example: id recycling with a free-list, or transactional
multi-slide moves), that would be the time to extract one. Until then,
F6 is the invariants and the documentation — not a class.


Delivering PRs
--------------

* `#68`  — ``Slides.move_slide(slide, new_idx)``
* `#67`  — ``Slides.delete(slide)``
* `#132` — ``Slides.duplicate(slide, index=None)``
* `#1036` — ``Slides.add_slide_from_external(source_slide, slide_layout)``
