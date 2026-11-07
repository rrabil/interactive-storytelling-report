# An Evaluation of Software Tools for Interactive Storytelling

A 2026 web edition of my 2012 technical report comparing four tools for writing branching game stories: articy:draft, Chat Mapper, the Dragon Age Toolset, and Excel.

**Live site:** https://storytelling.richardrabil.com/

Built with [Quarto](https://quarto.org/). The page is `index.qmd`, which includes one file per report section from `sections/`. The Editorial theme is in `editorial.scss`.

## Build and publish

```
quarto render                 # builds the site into _site/
quarto publish gh-pages       # renders and pushes the site to the gh-pages branch
```

## Accessibility check

After each publish, run [Pa11y](https://pa11y.org/) against the live site. It checks the page as a browser sees it, after scripts run, against WCAG 2 AA:

```
npx pa11y https://storytelling.richardrabil.com/
```

Pa11y flags images with missing or empty alt text but can't judge whether the text is any good. The alt text for every report figure lives in the `MEDIA` table in `_source/convert.py`, written by looking at each image.

## About the source files

The report sections in `sections/` were generated from the original Word document by `_source/convert.py`, which also applies the 2026 edits and gives the images in `media/` their readable names and alt text. The Word file and intermediate files stay local, so the script runs only on my machine. The splash, Results at a glance, and Making Of sections are hand-written.
