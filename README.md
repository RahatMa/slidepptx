# QBank Slides

Turn an Excel sheet of MCQs and problems into an **editable PowerPoint** deck (56 × 31.5 in, Kalpurush + Hind Siliguri, 72 pt, 1.5 line spacing). Runs entirely in the browser — nothing is uploaded anywhere.

## Put it on GitHub Pages

1. Create a new repository on GitHub (public, or private with a paid plan).
2. Upload **everything in this folder** (keep the folder structure: `index.html`, `js/`, `template/`, `examples/`, `fonts/`, `.nojekyll`).
3. Repository → **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main`, folder `/ (root)` → Save.
4. After a minute the site is live at `https://<your-user>.github.io/<repo-name>/`.

To update later, just upload changed files; Pages redeploys automatically.

## Using it

1. **Download the Excel template** from the page (or `template/question_template.xlsx`). Delete the example rows and add your own.
2. **Drop the Excel file** on the page, together with any question images (or one `.zip` of images).
3. Check the preview and settings, then **Download PowerPoint**.

### Columns

| Column | Required | Notes |
|---|---|---|
| ExamID | no | Filter / split files by exam |
| QNo | no | Shown in the header badge (`MCQ 07`, `Problem 07`) |
| Question | **yes** | Alt+Enter = new line. Math in `\( ... \)` |
| ImageURL | no | A file name you drop in with the Excel (`mcq24.png`), or a direct https link |
| OptionA–OptionD | MCQ | Leave all empty for a written Problem |
| Answer | MCQ | `A–D`, `ক–ঘ` or `1–4`. For Problems: optional final answer |
| Solution | no | Becomes a Solution slide + speaker notes |
| Type | no | `MCQ` / `Problem` — inferred when blank |
| Layout | no | `auto` / `side` / `grid` / `stack` for that row |

Column names are matched loosely (`Option A`, `option_a`, `Explanation`, `Ans` all work), and order doesn't matter. Sheets named *Instructions* are ignored; every other sheet with a Question column is read.

### Math

Write LaTeX inside `\( ... \)` (or `\[ ... \]` for an equation on its own line). Each formula becomes a **native PowerPoint equation** (the same object as *Insert ▸ Equation*): stacked fractions, vector arrows, radicals, sub/superscripts — click it in PowerPoint and edit it with the Equation tab. Examples: `\(\vec{F} = \frac{d\vec{p}}{dt}\)`, `\(MLT^{-2}\)`, `\(F_{\mathrm{net}}\)`, `\(\sqrt{2gh}\)`, `\(60^\circ\)`, `\(\mu\)`, `\Rightarrow`.

Equations use **Cambria Math** (built into Windows and Mac Office). Each equation box also carries a plain-text copy, which LibreOffice, Google Slides and Keynote show instead (they can't display PowerPoint equations). If you need the plain-text version everywhere, untick *Formulas as PowerPoint equations* in the app.

### Slides produced

- **MCQ:** question → answer (correct option green) → solution
- **Problem:** question → solution
- Each type can be switched off. Long solutions continue on extra slides.
- Text starts at 72 pt and only shrinks (4 pt steps, min 44 pt by default) if it would overflow; the page lists anything that still needs a look.

### Header & footer (slide master)

All slides use one layout called **Question Slide**. In PowerPoint open **View → Slide Master**, add the logo/header/footer there, and every slide updates. Alternatively upload a full-slide transparent PNG in *Header, footer and labels* before generating. The per-slide number badge can be turned off if your master provides one. The content area (default 4.9 in – 28.8 in from the top) is adjustable there too.

## Photos or PDFs of questions

Use the prompt in the app (*Have photos or PDFs instead?*) or [`CLAUDE_PROMPT.md`](CLAUDE_PROMPT.md). Claude returns the filled `.xlsx` plus a `.zip` of cropped figures — drop both onto the page.

## Fonts

Install **Kalpurush** and **Hind Siliguri** on the computer that presents the deck (fonts are not embedded). For the most accurate preview, have them installed on the computer that generates the deck too, or add `fonts/Kalpurush.ttf` to the repo.

## Files

```
index.html             page + styles
js/latex.js            LaTeX parser + plain-text fallback
js/omml.js             LaTeX → native PowerPoint equations (OMML) + preview HTML
js/layout.js           positions, font fitting, slide types (settings in DEFAULTS)
js/pptx-writer.js      writes the .pptx (PptxGenJS)
js/preview.js          in-browser preview
js/app.js              file handling and UI
js/claude-prompt.js    prompt for the photos/PDF route (keep CLAUDE_PROMPT.md in sync)
template/              Excel template with examples + instructions sheet
examples/images/       sample figure used by the example
```

Libraries (loaded from CDNs, pinned): SheetJS 0.20.3, PptxGenJS 4.0.1, JSZip 3.10.1.

Known limits: pictures embedded inside Excel cells can't be read (use file names); many image hosts block direct downloads from other sites (save the image and add it as a file).
