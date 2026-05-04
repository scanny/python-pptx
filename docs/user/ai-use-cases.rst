.. _ai-use-cases:

AI use cases
============

A recurring question (`issue #950
<https://github.com/scanny/python-pptx/issues/950>`_) is whether |pp| will
"include Generative AI" — i.e. whether it will grow an in-library interface
to an LLM that authors slides, rewrites decks, or generates imagery on
behalf of the caller. The short answer is **no**. |pp| is an OOXML
emitter / reader: it reads and writes the bytes of a ``.pptx`` package and
has no notion of models, prompts, API keys, or network calls. AI
orchestration — prompt construction, model selection, authentication,
rate limiting, cost accounting, retries, and safety filtering — is
firmly the caller's responsibility and belongs in caller code.

That said, |pp| is an excellent *tool* for AI-driven slide workflows.
Nearly every real-world "LLM that builds PowerPoint" pipeline uses
|pp| as the final stage — the LLM produces a structured description of
what the deck should contain, and |pp| turns that description into a
round-trip-clean ``.pptx`` file. The sections below sketch the patterns
that work well.


Patterns that compose well with |pp|
------------------------------------

* **Bulk slide authoring from LLM-generated JSON.** Prompt the model to
  emit a structured document (JSON, YAML, or a pydantic model) describing
  slides, titles, bullets, table rows, and chart series. Validate the
  structure in your own code. Then walk it and call |pp| to materialise
  the deck. Letting the model emit raw Python that calls |pp| directly
  works in a pinch but is harder to validate and sandbox; the
  structured-data approach is usually more robust.

* **Content translation / localisation.** Load an existing deck, iterate
  its text runs, translate each run via an LLM or translation API, and
  write the translated string back. This round-trips cleanly because
  |pp| preserves every non-text aspect of the deck (layout, theme,
  master, images, chart data) unchanged.

* **Accessibility alt-text generation.** Walk every picture on every
  slide, send the image bytes to a vision-capable model, and write the
  returned description to :attr:`Shape.alt_text`. |pp|'s accessibility
  metadata support (``alt_text`` description and short ``title`` on any
  shape) is the matching write surface.

* **Data-to-slide pipelines.** Pull data from a database, analytics
  warehouse, or spreadsheet; have an LLM decide which cuts of the data
  matter for a given audience and suggest a chart type per cut; then
  emit the charts with |pp|'s chart API. The LLM handles editorial
  judgement (what to show); |pp| handles the mechanical OOXML.

* **Deck summarisation / search indexing.** Read a deck with |pp|,
  extract text from slides and notes, and feed it to an LLM for
  summarisation or to an embedding model for semantic search. This is
  a read-only pattern and stays entirely on the |pp| ``pptx.Presentation``
  API plus whatever model client you use.


Out of scope
------------

The following are deliberately **not** |pp| concerns:

* **Rendering the generated deck.** Converting ``.pptx`` to PDF, PNG,
  MP4, or HTML is a *rendering* step, handled by a separate tool. A
  common complete pipeline is
  ``LLM -> python-pptx -> libreoffice --convert-to pdf``; the
  ``.pptx`` → rendered-output half is covered in detail at
  :ref:`rendering-to-pdf-video-or-image-formats`.

* **Model calls, API keys, authentication, and billing.** These are
  application-level concerns that vary per provider (OpenAI, Anthropic,
  Azure, Bedrock, Vertex, local Ollama / llama.cpp, …) and per
  deployment. |pp| has no opinion about which SDK you use or how you
  manage credentials.

* **Prompt templates and agent frameworks.** Libraries such as
  LangChain, LlamaIndex, DSPy, and the vendor SDKs exist precisely to
  solve the prompt / orchestration problem; |pp| is complementary to
  them, not a competitor.

* **Safety, moderation, and provenance.** If the generated deck will
  be shown to end users, the caller is responsible for any content
  filtering, provenance watermarking, or human-in-the-loop review
  required by policy.


Example: build a deck from an LLM-produced dict
-----------------------------------------------

The code below assumes a prior step — not shown — that prompts an LLM
to return a JSON document shaped like ``deck_spec``. Keep the model
call, error handling, retries, and schema validation in that upstream
step; the |pp| portion is just a deterministic walk over the
validated structure.

.. code-block:: python

    from pptx import Presentation
    from pptx.util import Inches

    # Hypothetical output from an LLM call, already parsed + validated.
    deck_spec = {
        "title": "Q3 Product Review",
        "slides": [
            {"title": "Highlights", "bullets": [
                "Revenue up 18% QoQ",
                "Two new enterprise logos",
                "Churn at an all-time low",
            ]},
            {"title": "Risks", "bullets": [
                "Dependency on a single supplier",
                "Tightening regulatory environment",
            ]},
        ],
    }

    prs = Presentation()
    title_layout = prs.slide_layouts[0]
    bullet_layout = prs.slide_layouts[1]

    # Cover slide
    cover = prs.slides.add_slide(title_layout)
    cover.shapes.title.text = deck_spec["title"]

    # Bullet slides
    for slide_spec in deck_spec["slides"]:
        slide = prs.slides.add_slide(bullet_layout)
        slide.shapes.title.text = slide_spec["title"]
        tf = slide.shapes.placeholders[1].text_frame
        tf.text = slide_spec["bullets"][0]
        for bullet in slide_spec["bullets"][1:]:
            tf.add_paragraph().text = bullet

    prs.save("q3-review.pptx")

Everything above the ``Presentation()`` call is application logic —
including whatever LLM SDK produced ``deck_spec`` — and is out of scope
for |pp|. Everything from ``Presentation()`` down is a plain,
deterministic OOXML emit using the same public API documented
throughout the rest of this guide.
