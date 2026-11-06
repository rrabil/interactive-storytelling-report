# An Evaluation of Software Tools for Interactive Storytelling

A 2026 web edition of my 2012 technical report comparing four tools for writing branching game stories: articy:draft, Chat Mapper, the Dragon Age Toolset, and Excel.

**Live site:** https://rrabil.github.io/interactive-storytelling-report/

Built with [Quarto](https://quarto.org/). The page is `index.qmd`, which includes one file per report section from `sections/`. The Editorial theme is in `editorial.scss`.

## Build and publish

```
quarto render                 # builds the site into _site/
quarto publish gh-pages       # renders and pushes the site to the gh-pages branch
```

## About the source files

The report sections in `sections/` were generated from the original Word document by `_source/convert.py`, which also applies the 2026 edits. The Word file and intermediate files stay local, so the script runs only on my machine. The splash, Results at a glance, and Making Of sections are hand-written.
