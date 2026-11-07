"""Convert the 2012 report (Pandoc Markdown + source docx) into Quarto section files.

Pipeline:
  1. pandoc Rabil-Final-Report-v11_1.docx -t markdown --wrap=none --extract-media=. -o _source/raw.md
  2. python _source/convert.py            (this script; writes sections/*.qmd and
                                            renames media/imageN to the names in MEDIA)

Rules (from the spec): the original text stays as it was. Only mechanical fixes are
made here: broken figure/table/section references, table rebuilding, icon-to-text
symbols, list nesting, and formatting. Every edit must match exactly once, so a
change in the source makes this script fail loudly instead of silently.
"""

import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "_source"
OUT = ROOT / "sections"
DOCX = SRC / "Rabil-Final-Report-v11_1.docx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# Capability-matrix icons (decision 6): image7 green check, image8 yellow warning, image9 red X.
SYMBOLS = {"image7.png": "✓ Yes", "image8.png": "◐ Partial", "image9.png": "✗ No"}

# ---------------------------------------------------------------- docx access

with zipfile.ZipFile(DOCX) as z:
    doc_xml = z.read("word/document.xml").decode("utf-8")
    rels_xml = z.read("word/_rels/document.xml.rels").decode("utf-8")

RID_TO_IMAGE = {
    m.group(1): m.group(2)
    for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="media/([^"]+)"', rels_xml)
}
ROOT_EL = ET.fromstring(doc_xml)
BODY = ROOT_EL.find(W + "body")


def para_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def top_tables(el):
    for ch in el:
        if ch.tag == W + "tbl":
            yield ch
        elif ch.tag in (W + "sdt", W + "sdtContent"):
            yield from top_tables(ch)


def run_is(run, tag):
    el = run.find(W + "rPr/" + W + tag)
    return el is not None and el.get(W + "val") not in ("0", "false")


def styled(txt, bold, ital):
    """Wrap text in Markdown emphasis, keeping surrounding spaces outside the markers."""
    core = txt.strip()
    if not core or not (bold or ital):
        return txt
    mark = "***" if bold and ital else "**" if bold else "*"
    lead, trail = txt[: len(txt) - len(txt.lstrip())], txt[len(txt.rstrip()):]
    return f"{lead}{mark}{core}{mark}{trail}"


def cell_text(tc):
    """Text of a table cell with run-level bold/italic; icon images become text symbols."""
    parts = []
    for p in tc.findall(W + "p"):
        spans = []  # [text, bold, ital], adjacent runs with the same style merged
        for run in p.iter(W + "r"):
            blip = run.find(".//" + A + "blip")
            if blip is not None:
                spans.append([SYMBOLS[RID_TO_IMAGE[blip.get(R + "embed")]], False, False])
                continue
            txt = "".join(t.text or "" for t in run.findall(W + "t"))
            if not txt:
                continue
            style = [run_is(run, "b"), run_is(run, "i")] if txt.strip() else None
            if spans and (style is None or spans[-1][1:] == style):
                spans[-1][0] += txt
            else:
                spans.append([txt] + (style or [False, False]))
        s = " ".join("".join(styled(*sp) for sp in spans).split())
        if s:
            parts.append(s)
    return " ".join(parts).replace("|", "\\|")


def table_rows(tbl):
    return [[cell_text(tc) for tc in tr.findall(W + "tc")] for tr in tbl.findall(W + "tr")]


def unbold(s):
    return s[2:-2] if s.startswith("**") and s.endswith("**") else s


def pipe_table(rows, header=None, skip=0, bold_short_rows=False, widths=None):
    """Build a pipe table. header=None uses the first remaining row; header=[] makes
    an empty (hidden) header. Rows shorter than the header keep their first cell
    in column 1 and push the remaining cells to the right-hand columns."""
    rows = rows[skip:]
    if header is None:
        header, rows = [unbold(c) for c in rows[0]], rows[1:]
    n = len(header) if header else max(len(r) for r in rows)
    body = []
    for r in rows:
        if len(r) < n:
            first = r[0]
            if bold_short_rows and first and not first.startswith("**"):
                first = f"**{first}**"
            r = [first] + [""] * (n - len(r)) + r[1:]
        body.append(r)
    # Tool-name headers may break after the colon (ARTICY: / DRAFT) so narrow columns stay narrow.
    hdr = [c.replace("articy:draft", "articy:<wbr>draft") for c in header] if header else [" "] * n
    out = ["| " + " | ".join(hdr) + " |",
           "|" + "|".join("-" * w for w in (widths or col_widths(header, body, n))) + "|"]
    out += ["| " + " | ".join(r) + " |" for r in body]
    return "\n".join(out)


def col_widths(header, body, n):
    """Relative column widths, which Pandoc reads from the separator dashes. Each
    column gets the larger of its longest word and its median cell length, so
    symbol and number columns stay narrow and prose columns get the room."""
    widths = []
    for j in range(n):
        cells = [r[j] for r in body if j < len(r) and r[j].strip()]
        head = header[j].replace(":", ": ") if header else ""
        longest = max((len(w) for c in cells + [head] for w in c.split()), default=3)
        lens = sorted(len(c) for c in cells) or [3]
        widths.append(max(3, longest, lens[len(lens) // 2]))
    return widths


TABLES = list(top_tables(BODY))
assert len(TABLES) == 21, len(TABLES)
T = [table_rows(t) for t in TABLES]


def slug(s):
    s = re.sub(r"[^\w\s.-]", "", s.lower())
    s = re.sub(r"\s+", "-", s.strip())
    return re.sub(r"^[^a-z]+", "", s)


def build_table(i):
    rows = T[i]
    if i == 0:  # survey weightings example, Collaboration
        title = "**Collaboration**\n\n*(7 = “Critical” and 1 = “Irrelevant”)*\n\n"
        return title + pipe_table(rows, skip=1)
    if i in (1, 3):  # scale for subjective scores; icon legend (no header row in original)
        return pipe_table(rows, header=[])
    if i == 4:  # capability matrix, one table per capability category (stage 2 review)
        hdr = ["Capability", "articy:draft", "Chat Mapper", "DA Toolset", "Excel"]
        groups = []
        for r in rows[2:]:
            if len(r) < len(hdr):  # category row
                groups.append((unbold(r[0]), []))
            else:
                groups[-1][1].append(r)
        # Shared column widths so the seven tables line up as the reader scrolls.
        widths = col_widths(hdr, [r for _, body in groups for r in body], len(hdr))
        # Explicit ids: Appendix B has headings with the same names.
        return "\n\n".join(
            f"### {name} {{#capabilities-{slug(name)}}}\n\n" + pipe_table(body, header=hdr, widths=widths)
            for name, body in groups)
    if i == 5:  # final scores
        legend = rows[0][0].removeprefix("Capabilities").strip()
        hdr = ["Capabilities", "articy:draft", "Chat Mapper", "DA Toolset", "Excel", "Avrg Wgt"]
        return legend + "\n\n" + pipe_table(rows, header=hdr, skip=2, bold_short_rows=True)
    return pipe_table(rows)  # Appendix B score tables, Appendix C survey tables


# ---------------------------------------------------------------- figure captions

def figure_captions():
    best = {}
    for p in ROOT_EL.iter(W + "p"):
        m = re.match(r"\s*Figure (\d+)\. (.*)", para_text(p), re.S)
        if m:
            n, cap = int(m.group(1)), m.group(2).strip()
            if n not in best or len(cap) < len(best[n]):
                best[n] = cap
    return best


CAPTIONS = figure_captions()
assert sorted(CAPTIONS) == list(range(1, 28)), sorted(CAPTIONS)

# (figure number, image, Quarto id, text of the paragraph the figure follows)
FIGURES = [
    (1, "image2.jpeg", "fig-skyrim", "shows an example of branching dialogue from *The Elder Scrolls V: Skyrim*."),
    (2, "image3.png", "fig-roles", 'Participants in the "Other" category added the following designations'),
    (3, "image4.png", "fig-tool-usage", "Twelve of the 33 participants had never written branching dialogue"),
    (4, "image5.png", "fig-collab-rankings", "depicts the rankings that participants gave to factors listed under the Collaboration category"),
    (5, "image6.png", "fig-price", "people would be willing to pay extra money to use articy:draft"),
    (6, "image10.png", "fig-articy-flow", "The results of my evaluation show that articy:draft is the most promising"),
    (7, "image11.png", "fig-chatmapper-sim", "depending on one's budget, technical requirements"),
    (8, "image12.png", "fig-area-editor", "Although Excel received the lowest score, it has notable strengths"),
    (9, "image13.png", "fig-version-control", "The DA Toolset lets you check files in and out of a database"),
    (10, "image14.png", "fig-articy-storyboard", "for plotting the sequences of events in your game in a storyboard fashion"),
    (11, "image15.png", "fig-excel-outline", "Users can create manual outlines in Excel"),
    (12, "image16.png", "fig-dat-tree", "The DA Toolset has a Conversation Editor that displays dialogue in nodes"),
    (13, "image17.png", "fig-chatmapper-tree", "default display of the branching dialogue is vertical"),
    (14, "image18.png", "fig-articy-pins", "in which you can manually write a condition"),
    (15, "image19.png", "fig-dat-plots", "tab that allows users to set conditions for any line of dialogue"),
    (16, "image20.png", "fig-chatmapper-scripting", "Chat Mapper handles this ability through its Scripting Editor"),
    (17, "image21.png", "fig-chatmapper-simulator", "Chat Mapper comes with a full-fledged Conversation Simulator"),
    (18, "image22.png", "fig-articy-entity", "articy:draft has an Entity Editor for filling out character profiles"),
    (19, "image23.png", "fig-excel-character", "Excel does not come with predefined character sheets, but such sheets can be easily created"),
    (20, "image24.png", "fig-chatmapper-settings", "Project Settings tool. (See Figure 20.)"),
    (21, "image25.png", "fig-dat-area", "The DA Toolset user interface is not intuitive"),
    (22, "image26.png", "fig-dat-script", "For basic writing tasks (such as developing branches"),
    (23, "image27.png", "fig-articy-split", "articy:draft has a wide range of flexible viewing options"),
    (24, "image28.png", "fig-chatmapper-layout", "Chat Mapper lets you configure and drag around each tabbed window"),
    (25, "image29.png", "fig-chatmapper-help", "Chat Mapper has an in-application help system that covers the basics"),
    (26, "image30.png", "fig-dat-wiki", "The DA Toolset has a range of video tutorials created by fans"),
    (27, "image31.png", "fig-excel-help", "Excel has a low learning curve, especially if you use other Microsoft Office"),
]
FIG_ID = {n: fid for n, _, fid, _ in FIGURES}

# Readable file names and alt text for every image Pandoc extracts as media/imageN.
# The number stays as a prefix so the mapping back to the Word file is obvious.
# Alt text describes what is in the image (the caption sits beside it), and charts
# carry their numbers. It replaces the Word file's own alt text, which was thin and,
# for image12, just a leftover file name. No double quotes: it lands in fig-alt="...".
MEDIA = {
    "image2.jpeg": ("02-skyrim-dialogue-choices.jpeg",
        "Scene from Skyrim. Jarl Balgruuf sits on a wooden throne in a stone hall beside a standing advisor. A menu next to him lists the player's two dialogue choices, one about having a great view of the execution, one about the dragon destroying Helgen."),
    "image3.png": ("03-survey-q1-roles.png",
        "Bar chart of survey participants' roles. Game developer or designer: 11. Game writer or editor: 9. Narrative designer: 3. Aspiring game writer: 3. Other: 3. Content writer: 2. Game writing student: 2. Copywriter and game writing hobbyist: 0."),
    "image4.png": ("04-survey-q4-dialogue-tools.png",
        "Bar chart of tools participants have used to write branching dialogue. Have not written branching dialogue: 12. Microsoft Word or other word processor: 10. Excel or other spreadsheet: 9. articy:draft beta: 4. Chat Mapper: 4. A proprietary toolset: 4. Aurora Toolset: 1. Scrivener: 1. Celtx, ChoiceScript, MediaWiki and Wide Ruled: 0."),
    "image5.png": ("05-survey-collaboration-rankings.png",
        "Grouped bar chart of how participants ranked five collaboration and story-design features, from 1 (critical) to 7 (irrelevant). Creating and editing a story outline drew the most critical votes, 12. Version history drew 8, linking content to an outline 7, a storyboard view 6, and web-based collaboration 1, with most of its votes spread across ranks 2 to 5."),
    "image6.png": ("06-survey-q17-price-range.png",
        "Bar chart of what participants consider an affordable single-user license. $0 to $50: 4. $50 to $200: 7. $200 to $400: 3. $400 to $600: 2. $600 to $800: 0. $800 to $1,000: 1."),
    "image10.png": ("10-articy-flow-and-dialog-editors.png",
        "articy:draft window split into two panels. The top panel shows the story flow as a row of colored boxes joined by arrows. The bottom panel shows the Prison Talk conversation nested in one of those boxes, as branching dialogue nodes for Brynn Harrowmage and Cush ap Rogar. A project tree runs down the left."),
    "image11.png": ("11-chatmapper-conversation-simulator.png",
        "Two overlapping Chat Mapper simulator windows with callout arrows. In one, the author picks from a list of dialogue options under a portrait of Brynn Harrowmage. In the other, Cush ap Rogar's portrait appears with his reply as on-screen text. Callouts say authors can simulate the dialogue options they wrote, and preview how their dialogue looks and sounds when delivered by an NPC."),
    "image12.png": ("12-dat-area-editor-overview.png",
        "Dragon Age Toolset Area editor. A 3D view of a swamp village at dusk, with placeholder character figures, a flag and a camera marker, sits between a tree of area objects on the left and a palette of resource folders on the right."),
    "image13.png": ("13-dat-version-control.png",
        "Dragon Age Toolset with the Resource History panel open. It lists two check-ins of the same file with date and user, and a red circle highlights a right-click menu with View Resource, Restore Resource and Restore to New Resource."),
    "image14.png": ("14-articy-storyboard-flow-view.png",
        "articy:draft Flow view. About a dozen colored story boxes, each holding a short text summary, are connected left to right by arrows that branch and rejoin. A project tree on the left lists the flow fragments."),
    "image15.png": ("15-excel-story-outline-hyperlinks.png",
        "Excel worksheet titled Story Outline listing the game's parts in column A. Parts 1 and 2 and their sub-steps are blue hyperlinks; Parts 3 to 7 are plain text."),
    "image16.png": ("16-dat-dialogue-tree.png",
        "Dragon Age Toolset Conversation Editor. A nested tree of dialogue lines, with the quest giver's lines in red and the player's replies in blue, sits above a Plots and Scripting tab. A palette of resource folders is on the right."),
    "image17.png": ("17-chatmapper-dialogue-tree.png",
        "Chat Mapper Prison Chatter conversation laid out as a top-down tree of dialogue cards, each with a portrait, speaker names and the line of dialogue. A minimap is at top right and a Properties panel for the selected line below it."),
    "image18.png": ("18-articy-inbound-pin-condition.png",
        "Close-up of an articy:draft dialogue node for Brynn Harrowmage asking: The Creeps? Who, or what, are the Creeps? An inbound pin on its left edge holds the condition IfAskedAboutPrison."),
    "image19.png": ("19-dat-plots-and-scripting-tab.png",
        "Close-up of the Dragon Age Toolset Plots and Scripting tab. The Condition section is set to plot clear_the_hut, is true, flag QUEST_ACCEPTED, so the line shows only after the quest is accepted."),
    "image20.png": ("20-chatmapper-scripting-editor.png",
        "Close-up of Chat Mapper with a dialogue card for Cush ap Rogar above the Script Editor, which holds one Lua line that sets the Creep actor's status to LearnedAbout."),
    "image21.png": ("21-chatmapper-simulator-tracking.png",
        "Chat Mapper Conversation Simulator. Brynn Harrowmage's portrait and four dialogue options are in the center. Empty variable, status and relationship tracker panels run down the left, and a list of possible dialogue choices with their conditions is on the right."),
    "image22.png": ("22-articy-character-entity.png",
        "articy:draft entity sheet for Brynn Harrowmage with portrait, a paragraph of backstory, and fields for type, age, sex, species, profession, birthplace, personality and inner conflict. The project tree on the left lists NPCs and locations."),
    "image23.png": ("23-excel-character-sheet.png",
        "Excel worksheet used as a character sheet. Each row is an NPC, here Tamrius Sextus and Ji Ling, with columns for name, NPC type, a picture, race, class, alignment, background, age, internal conflict and motivation."),
    "image24.png": ("24-chatmapper-custom-asset-fields.png",
        "Chat Mapper Project Settings, Custom Asset Fields tab for Actors. A table of required fields (Name, Pictures, Description, IsPlayer) sits above a table of custom fields such as Age, Gender, Faction, Class and Hometown, each with a type, default value, hint, and Show and Export checkboxes."),
    "image25.png": ("25-dat-area-editor-panels.png",
        "Dragon Age Toolset Area editor with five panels: an object tree, a 3D swamp scene with placeholder figures, a resource palette, an Object Inspector of properties such as Variable 2da and NameRequiresReTranslation, and a log window."),
    "image26.png": ("26-dat-script-editor.png",
        "Dragon Age Toolset script editor showing a short C-style script that sets a plot flag when a team of monsters is destroyed, with a long alphabetical list of script functions beside it and a resource palette on the right."),
    "image27.png": ("27-articy-split-screen.png",
        "articy:draft in a four-panel layout: the Prison Talk dialogue flow at top left, Brynn Harrowmage's character entity at top right, a floor plan of Prison Darkhenge with labeled cells and rooms at bottom left, and an assets tree at bottom right."),
    "image28.png": ("28-chatmapper-panel-docking.png",
        "Chat Mapper window mid-rearrangement. A panel is being dragged and a translucent blue highlight with docking handles shows where it will land, over a conversation tree flanked by asset lists and a Properties panel."),
    "image29.png": ("29-chatmapper-context-help.png",
        "Chat Mapper with three small blue question-mark help buttons circled in red: on the Dialogue Text toolbar, the Script Editor and the Conditions panel."),
    "image30.png": ("30-dat-wiki-plots-tutorial.png",
        "Firefox showing the Dragon Age Toolset wiki page Keeping track of quests via plots. Explanatory text, a yellow tip box and a list of three flags (QUEST_ACCEPTED, MONSTERS_SLAIN, REWARD_RECEIVED) sit above an embedded screenshot of the plot editor."),
    "image31.png": ("31-excel-help-export-xml.png",
        "An Excel dialogue script with actor names and lines in two columns, partly covered by an Excel Help window on exporting XML data from mapped cells, with numbered steps and a picture of the XML ribbon group."),
    "image32.png": ("32-survey-q2-industry-experience.png",
        "Pie chart: asked whether they work or have worked in the video game industry creating or editing story content, 21 participants said yes and 12 said no."),
    "image33.png": ("33-survey-q3-branching-dialogue.png",
        "Pie chart: asked whether they have written conversation trees or branching dialogue, 20 participants said yes and 11 said no."),
}
assert all('"' not in alt for _, alt in MEDIA.values())

IMG_RE = re.compile(r'!\[([^\]]*)\]\(\./media/(image\d+\.\w+)\)\{width="([\d.]+)in" height="[\d.]+in"\}')
TEXT_WIDTH_IN = 7.0


def attrs(alt, width_in, fid=None):
    alt = " ".join(alt.replace("\\'", "'").replace('\\"', "'").split())
    a = ([f"#{fid}"] if fid else []) + [f'fig-alt="{alt}"']
    pct = round(float(width_in) / TEXT_WIDTH_IN * 100)
    if pct < 90:
        a.append(f'width="{pct}%"')
    return "{" + " ".join(a) + "}"


# ---------------------------------------------------------------- load raw

raw_text = (SRC / "raw.md").read_text(encoding="utf-8")
raw = raw_text.split("\n")

# Alt text and width for every image, collected before tables (and the stray
# caption lines inside them) are replaced. First occurrence wins.
IMAGE_INFO = {}
for m in IMG_RE.finditer(raw_text):
    IMAGE_INFO.setdefault(m.group(2), (m.group(1), m.group(3)))

# Table blocks in raw.md (1-based, inclusive), in document order. Two blocks also
# swallow a stray ": caption" line Pandoc emitted for a floating figure.
BLOCKS = [(215, 325), (335, 347), (430, 446), (452, 458), (466, 556), (562, 670),
          (692, 730), (736, 754), (894, 907), (966, 979), (1040, 1049), (1086, 1103),
          (1169, 1184), (1241, 1254), (1304, 1319), (1411, 1431), (1459, 1479),
          (1499, 1515), (1527, 1549), (1561, 1583), (1601, 1621)]
KEEP_PANDOC = {2, 6, 7}  # weightings summary; strengths/weaknesses and cost grids (cells hold lists)
for i, (a, b) in reversed(list(enumerate(BLOCKS))):
    first = raw[a - 1]
    assert first.startswith(("<table", "  ---", "+---")), (i, first[:40])
    if i in KEEP_PANDOC:
        continue
    raw[a - 1:b] = [build_table(i)]

text = "\n".join(raw)


def sub_once(old, new):
    global text
    n = text.count(old)
    assert n == 1, f"expected 1 match, found {n}: {old[:70]!r}"
    text = text.replace(old, new)


def sub_all(old, new, count):
    global text
    n = text.count(old)
    assert n == count, f"expected {count} matches, found {n}: {old[:70]!r}"
    text = text.replace(old, new)


# ---------------------------------------------------------------- front matter

# Title, author, date go in index.qmd front matter; the Word TOC is replaced by
# Quarto's floating TOC; About the Author moves to Making Of (PDF text, per spec).
start = text.index("# Abstract")
toc_map_src = text[text.index("# Table of Contents"):text.index("# Introduction")]
text = text[start:]
cut_a, cut_b = text.index("# About the Author"), text.index("# Introduction")
text = text[:cut_a] + text[cut_b:]

# Section numbers and anchors from the original Word TOC.
TOC = {}
for num, title, anchor in re.findall(r"\[([\d.]+) (.+?) \[\d+\]\(#[^)]+\)\]\(#([^)]+)\)", toc_map_src):
    TOC[anchor] = num
assert len(TOC) == 49, len(TOC)

# ---------------------------------------------------------------- mechanical text fixes

# Figure numbers Pandoc lost (Word SEQ/REF fields).
sub_once("\\`1 shows an example", "Figure 1 shows an example")
sub_once("\nFigure . An example of basic branching dialogue from The Elder Scrolls V: Skyrim, published by Bethesda Softworks in November, 2011.\n", "\n")
sub_once("(See .) In programming terms", "(See Figure 22.) In programming terms")

# Section references Pandoc lost; numbers and titles confirmed against the published PDF.
sub_once('please see Section , "."', 'please see Section 8, "Appendix A: Rationale for Tool Selection."')
sub_once('In Section , "[]{.mark}Conclusion,"', 'In Section 6, "Conclusion,"')
sub_once('As I discuss in Section , ","', 'As I discuss in Section 5.3, "Which Tool Should I Get? Summary of Recommendations,"')
sub_once('(See Section , ".")', '(See Section 4.1.4, "Participant Perspectives on Cost.")')
sub_once('(See Section , "Participant Experience and Tools"', '(See Section 4.1.2, "Participant Experience and Tools"')
sub_once("included in the table in Section , but", "included in the table in Section 4.2, but")

# Word highlighting and stray bold line breaks.
text = text.replace("[]{.mark}", "")
text = re.sub(r" ?\*\*\\\n\*\*", "", text)

# Heading formatting.
sub_once("# **Methodology**", "# Methodology")

# List nesting Pandoc flattened (checked against the PDF).
for term in ("*Conversation editor*", "*Node*", "*Branch*", "*Link*"):
    sub_once(f"\n-   {term}:", f"\n    -   {term}:")
for tool in ("articy:draft (www", "Chat Mapper (www", "ChoiceScript (", "Wide Ruled (http",
             "Aurora Toolset (", "Dragon Age Toolset (http", "The Garden of Eden", "Celtx (",
             "Microsoft Excel (www", "MediaWiki (", "Microsoft Word (www", "Scrivener ("):
    sub_once(f"\n-   {tool}", f"\n    -   {tool}")
sub_once("\n```{=html}\n<!-- -->\n```\n1.  Calculate", "\n2.  Calculate")
for k, step in enumerate(("Assign a subjective", "Add up the total", "Multiply each", "Add up the scores"), start=3):
    sub_once(f"{k - 1}.  {step}", f"{k}.  {step}")

# Capability legend paragraph: icons to text symbols.
legend = re.search(r"\n(!\[reCheckIcon[^\n]*\*\*Legend\*\*[^\n]*)\n", text).group(1)
sub_once(legend, '**Legend**: ✓ Yes = Exists; ◐ Partial = Exists with a workaround or is somehow "less than ideal"; ✗ No = Does not exist')

# ---------------------------------------------------------------- figures

# Appendix C charts have no captions: plain images, alt text only.
appc = text.index("# Appendix C")
head, tail = text[:appc], text[appc:]
tail = re.sub(r"(#### )(!\[[^\]]*\]\([^)]*\)\{[^}]*\})(Question 2 Results)", r"\1\3\n\n\2", tail)
tail = re.sub(r"(!\[[^\]]*\]\(\./media/image4\.png\)\{[^}]*\})(In the \"Other\")", r"\1\n\n\2", tail)
tail = IMG_RE.sub(lambda m: f"![](media/{m.group(2)}){attrs(m.group(1), m.group(3))}", tail)
text = head

# Numbered figures: drop Pandoc's inline anchor, place the figure after its paragraph.
text = IMG_RE.sub("", text)
lines = text.split("\n")
for n, img, fid, anchor in FIGURES:
    hits = [i for i, l in enumerate(lines) if anchor in l]
    assert len(hits) == 1, (n, anchor, len(hits))
    i = hits[0]
    alt, width = IMAGE_INFO[img]
    indent = "    " if lines[i].startswith("-   ") else ""
    lines[i] += f"\n\n{indent}![{CAPTIONS[n]}](media/{img}){attrs(alt, width, fid)}"
text = "\n".join(lines) + tail

# In-text figure references become Quarto cross-references.
sub_once("(See Figure 5 below.)", "(See Figure 8 below.)<!--NOTE:fig8-->")
text = re.sub(r"(?<![!\[])\bFigure (\d+)\b(?![^\[]*\]\(media/)", lambda m: "@" + FIG_ID[int(m.group(1))], text)

# ---------------------------------------------------------------- headings

def heading(m):
    level, title = m.group(1), m.group(2).strip()
    level = "###" if len(level) >= 4 else level  # Appendix B/C sub-items were Word heading 4/5
    sid = slug(title)
    if sid in TOC:
        return f"{level} {TOC[sid]} {title} {{#{sid}}}"
    return f"{level} {title}"


text = re.sub(r"^(#{1,5}) (.+)$", heading, text, flags=re.M)
# The Bibliography keeps its original number (7) and place, after the Conclusion
# (index.qmd includes it there), so the appendices' numbers 8 to 10 follow on.

# "Section N" references link to their headings.
NUM_TO_ID = {v: k for k, v in TOC.items()}


def section_link(m):
    num = m.group(1).rstrip(".")
    return f"[Section {num}](#{NUM_TO_ID[num]})" + m.group(1)[len(num):] if num in NUM_TO_ID else m.group(0)


text = re.sub(r"(?<!\[)\bSection (\d+(?:\.\d+)*\.?)", section_link, text)

# ---------------------------------------------------------------- 2026 corrections

# Errors in the original (decision 5). Their editor's notes live in Making Of,
# "Changes in the 2026 edition", not in the report body (stage 2 review, round 2).
sub_once("<!--NOTE:fig8-->", "")
sub_once("Character Design): | 46 |", "Character Design): | 41.4 |")
sub_once("Skorupski, J., & Mat, M. (2009)", "Skorupski, J., & Mateas, M. (2009)")
sub_once('"Game Scriptwriting Software" (2009), for example', '"Game Scriptwriting Software" (2010), for example')
sub_once("each tool's strengths and weaknesses l.", "each tool's strengths and weaknesses.")

# ---------------------------------------------------------------- 2026 editorial edits
# Richard's stage 2 review, round 1. Each is listed in Making Of, "Changes in the 2026 edition".

sub_once("The method of evaluation involved four components: (1) selection of tools to evaluate, (2) a review of the literature on software tools and non-linear storytelling, (3) a web-based survey to determine the priorities of game writers and designers when evaluating software, and (4) hours of firsthand testing and scoring of each tool's capabilities.",
         "The method of evaluation involved four components:\n\n"
         "-   A selection of tools to evaluate.\n"
         "-   A review of the literature on software tools and non-linear storytelling.\n"
         "-   A web-based survey to determine the priorities of game writers and designers when evaluating software.\n"
         "-   Hours of firsthand testing and scoring of each tool's capabilities.")
sub_once("divided into nine major sections. *[Section 2](#background)* includes",
         "divided into nine major sections.\n\n-   *[Section 2](#background)* includes")
for k in (3, 4, 5, 6):
    sub_once(f". *[Section {k}](#", f".\n-   *[Section {k}](#")
sub_once("might be conducted, while *Sections 7- 10* contain", "might be conducted.\n-   *Sections 7- 10* contain")
sub_once("To deepen my understanding of the needs of game writers and designers, to gather evidence for which criteria to include in my evaluation, and to inform the development of the online survey,",
         "To deepen my understanding of the needs of game writers and designers, gather evidence for which criteria to include in my evaluation, and inform the development of the online survey,")
sub_once("Gamasutra ([www.gamasutra.com](http://www.gamasutra.com))", "Gamasutra ([www.gamasutra.com](https://www.gamasutra.com))")
CATEGORIES = ["Collaboration and Story Design", "Branching Dialogue", "Character Design", "Locations",
              "Export and Integration Pipeline", "Usability and User Interface", "Cost, Learning and Support"]
sub_once('one of the following categories: "Collaboration and Story Design," "Branching Dialogue," "Character Design," "Locations," "Export and Integration Pipeline," "Usability and User Interface," and "Cost, Learning and Support." I derived',
         "one of the following categories:\n\n" + "\n".join(f"-   {c}" for c in CATEGORIES) + "\n\nI derived")
sub_once("four sources of information: (a) past research", "four sources of information: (1) past research")


def numbered(items):
    return "\n\n" + "\n".join(f"{k}.  {it}" for k, it in enumerate(items, 1)) + "\n\n"


# Round 3: inline (a)-(c) series in 3.1 become numbered lists.
sub_once("Tools were grouped under three major categories: (a) tools designed specifically for the needs of an interactive medium; (b) proprietary game engines or toolsets that include \"dialogue editors\" or \"conversation editors\"; and (c) tools designed for traditional, linear storytelling or for more generic content authoring purposes.",
         "Tools were grouped under three major categories:" + numbered([
             "Tools designed specifically for the needs of an interactive medium.",
             "Proprietary game engines or toolsets that include \"dialogue editors\" or \"conversation editors.\"",
             "Tools designed for traditional, linear storytelling or for more generic content authoring purposes."]).rstrip("\n"))
sub_once("Once I created the list, I applied three criteria to narrow it down: (a) the relative degree of each tool's applicability to the unique demands of nonlinear storytelling; (b) the relative degree of each tool's prominence among game writers; and (c) the cost of each tool, as well as whether it could be fairly and accurately evaluated in the given period of time. ",
         "Once I created the list, I applied three criteria to narrow it down:" + numbered([
             "The relative degree of each tool's applicability to the unique demands of nonlinear storytelling.",
             "The relative degree of each tool's prominence among game writers.",
             "The cost of each tool, as well as whether it could be fairly and accurately evaluated in the given period of time."]))

# Round 3: the 3.4.3 worked example as a column of scores with a total (.score-sum).
SCORES = [("Collaboration Score", "94.5"), ("Branching Dialogue Score", "64.68"),
          ("Character Design Score", "41.4"), ("Assets, Locations, and Integration Score", "88.4"),
          ("Usability and User Interface Score", "102.6"), ("Cost Learning and Support Score", "62.28"),
          ("Other Factors Suggested by Survey Respondents Score", "10")]
sub_once(" + ".join(f"{name} ({n})" for name, n in SCORES) + " = **463.86**",
         "::: {.score-sum}" + numbered([f"{name}: [{n}]{{.pts}}" for name, n in SCORES])
         + "[Total:]{.total-label} [463.86]{.pts}\n:::")

# Conclusion: the opening paragraph (a restatement of the problem) is cut, so the
# section starts with the findings (Richard, 2026-11-06).
text, n = re.subn(r"\nThe interactive, non-linear characteristics of stories in video games pose[^\n]*\n\n"
                  r"To fill this gap, this report presented", "\nThis report presented", text)
assert n == 1, n

# Pandoc swallowed the closing parenthesis into the UMLet link (found by Lychee).
sub_all("<http://www.umlet.com/)>", "<http://www.umlet.com/>)", 2)

# Round 3: cross-references say "see", not "please see".
sub_all("please see ", "see ", 5)
sub_all("Please see ", "See ", 9)
sub_all("**Important Note**:", "**Note**:", 2)
sub_all("**Important Note:**", "**Note:**", 4)

# 5.2 pricing sources: titled links instead of bare URLs.
NEVIGO = "http://www.nevigo.com/index.php?id=56"
sub_once("-   Order page on Nevigo's Web site:\\", f"-   [Order page on Nevigo's Web site]({NEVIGO})")
sub_once(f"    <{NEVIGO}>", "")
for label in ["Pricing page on Chat Mapper's Web site",
              "Pricing page on Amazon for PC copy of Dragon Age: Origins",
              "Download page on the BioWare Social Network",
              "Pricing page on Microsoft Office Web site"]:
    m = re.search(re.escape(label) + r"( \(look under used copies\))?: <([^>]+)>", text)
    assert m, label
    text = text.replace(m.group(0), f"[{label}]({m.group(2)})" + (m.group(1) or ""))


def fix_grid_tables(src):
    """Re-pad Pandoc grid tables whose cells outgrew their columns after the
    substitutions above (the 5.1 row holding the Section 5.2 link broke this way).
    Columns come out equal width, since these grids compare four tools side by side."""
    out, lines, i = [], src.split("\n"), 0
    while i < len(lines):
        if not lines[i].startswith("+-"):
            out.append(lines[i])
            i += 1
            continue
        j = i
        while j < len(lines) and lines[j].startswith(("+", "|")):
            j += 1
        block = lines[i:j]
        n = block[0].count("+") - 1
        rows = [l[1:-1].split("|") for l in block if l.startswith("|")]
        assert all(len(r) == n for r in rows), block[0][:40]
        w = max(len(c.rstrip()) for r in rows for c in r) + 1
        for l in block:
            if l.startswith("+"):
                out.append("+" + "+".join(["=" if "=" in l else "-"][0] * w for _ in range(n)) + "+")
            else:
                out.append("|" + "|".join(c.rstrip().ljust(w) for c in l[1:-1].split("|")) + "|")
        i = j
    return "\n".join(out)


text = fix_grid_tables(text)

# Readable image names and alt text (MEDIA, above). Every image reference must be in
# the table, and every table entry must be used, so a new or dropped figure fails here.
seen = set()


def rename_image(m):
    old, attrs_ = m.group(1), m.group(2)
    new, alt = MEDIA[old]
    seen.add(old)
    attrs_, n = re.subn(r'fig-alt="[^"]*"', f'fig-alt="{alt}"', attrs_)
    assert n == 1, old
    return f"](media/{new}){{{attrs_}}}"


text = re.sub(r"\]\(media/(image\d+\.\w+)\)\{([^}]*)\}", rename_image, text)
assert seen == set(MEDIA), set(MEDIA) ^ seen

# Pandoc extracts media/imageN; rename to match. Skips files already renamed.
MEDIA_DIR = ROOT / "media"
for old, (new, _) in MEDIA.items():
    if (MEDIA_DIR / old).exists():
        (MEDIA_DIR / old).replace(MEDIA_DIR / new)
    assert (MEDIA_DIR / new).exists(), new

# "Note:" paragraphs get their own styled block (.report-note in editorial.scss).
# The original mixes "**Note**:" and "**Note:**"; both become "**Note:**".
text, n_notes = re.subn(r"(?m)^\*\*Note(?:\*\*:|:\*\*) (.+)$", r"::: {.report-note}\n**Note:** \1\n:::", text)
assert n_notes == 15, n_notes

# ---------------------------------------------------------------- split into files

def section(start_heading, end_heading=None):
    a = text.index(start_heading)
    b = text.index(end_heading) if end_heading else len(text)
    return text[a:b].strip() + "\n"


files = {
    "abstract.qmd": section("# Abstract", "# 1 Introduction"),
    "introduction.qmd": section("# 1 Introduction", "# 2 Background"),
    "background.qmd": section("# 2 Background", "# 3 Methodology"),
    "methodology.qmd": section("# 3 Methodology", "# 4 Results"),
    "results.qmd": section("# 4 Results", "# 5 Discussion"),
    "discussion.qmd": section("# 5 Discussion", "# 6 Conclusion"),
    "conclusion.qmd": section("# 6 Conclusion", "# 7 Bibliography"),
    # Request for Feedback, which followed the Bibliography, is dropped (stage 2 review, round 4).
    "references.qmd": section("# 7 Bibliography", "# Request for Feedback"),
    "appendix-a.qmd": section("# 8 Appendix A", "# 9 Appendix B"),
    "appendix-b.qmd": section("# 9 Appendix B", "# 10 Appendix C"),
    "appendix-c.qmd": section("# 10 Appendix C"),
}


# Appendix C (decision 8): each survey screen folds into a collapsed callout.
appc_parts = re.split(r"(?m)^(?=## 10\.\d)", files["appendix-c.qmd"])
files["appendix-c.qmd"] = appc_parts[0] + "".join(
    f'::: {{.callout-note collapse="true" appearance="simple" icon="false"}}\n{p.strip()}\n:::\n\n'
    for p in appc_parts[1:]
)

# Footnotes live at the end of raw.md; Pandoc resolves them per include, so move them
# into the file that uses them.
notes = re.findall(r"(?m)^\[\^\d+\]: .*$", files["appendix-c.qmd"])
for n in notes:
    files["appendix-c.qmd"] = files["appendix-c.qmd"].replace(n, "")
files["appendix-c.qmd"] = files["appendix-c.qmd"].rstrip() + "\n"
files["background.qmd"] += "\n" + "\n\n".join(notes) + "\n"

OUT.mkdir(exist_ok=True)
for name, body in files.items():
    (OUT / name).write_text(body, encoding="utf-8")
    print(f"wrote sections/{name}  ({body.count(chr(10))} lines)")

leftover = re.findall(r"Section , |Figure \.|\.mark|\]\(\./media", text)
assert not leftover, leftover
