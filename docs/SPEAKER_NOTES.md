# Markdown Speaker Notes Format Specification

`latex2pptx` allows you to attach speaker notes to PowerPoint slides directly from Markdown (`.md`) files. This document explains the supported formats, syntax, and CLI usage.

---

## Quick Usage

You can supply notes via the `-n` or `--notes` option:

```bash
# Convert LaTeX presentation and attach notes
latex2pptx presentation.tex -n notes.md -o presentation.pptx

# Convert compiled PDF and attach notes
latex2pptx presentation.pdf -n notes.md -o presentation.pptx
```

### Auto-Discovery
If you name your notes file `<input_stem>.notes.md` in the same directory as your presentation (e.g. `clase1.notes.md` for `clase1.tex`), `latex2pptx` automatically detects and attaches the notes without needing the `-n` flag!

---

## Supported Formats

`latex2pptx` supports three intuitive formats:

### 1. Explicit Slide Headings (Recommended)

Specify slide numbers directly in markdown headings (levels `#`, `##`, or `###`). Both English and Spanish keywords are supported:

```markdown
# Slide 1: Welcome & Introductions
Welcome everyone to today's workshop.
- Briefly introduce yourself
- Remind everyone that the session is recorded

## Slide 2: Agenda
Walk through the 3 main sections of the talk:
1. Current State
2. Practical Use Cases
3. Live Hands-on Demo

### Slide 3: Background
Context on why this topic matters today.
```

#### Supported Heading Keywords:
- **English**: `Slide`, `Slides`, `Page`, `Pages`, or plain numbers (e.g. `# 1. Overview`)
- **Spanish**: `Diapositiva`, `Diapositivas`, `Diapo`, `Diapos`, `Página`, `Páginas`, `Pág`, `Págs`

---

### 2. Slide Number Ranges (For Overlays)

LaTeX Beamer presentations often use overlays (`\only<1>`, `\only<2-3>`, `\pause`), which `latex2pptx` converts into sequential PowerPoint slides. You can map notes across multiple overlay slides using range syntax:

```markdown
# Slide 1: Welcome
Initial welcome note.

## Slide 2-4: Three Pillars of Modern AI
This note will be automatically attached to Slide 2, Slide 3, and Slide 4!
- Point 1: Data
- Point 2: Models
- Point 3: Evaluation

## Diapositivas 5 a 8: Casos de uso
Attached to slides 5, 6, 7, and 8.
```

Range delimiters supported: `-`, `to`, or `a` (e.g., `Slide 4-6`, `Slides 7 to 9`, `Diapositivas 10 a 12`).

---

### 3. HTML Comment Directives

If you prefer to keep slide numbers out of visible markdown text, use HTML comment directives:

```markdown
<!-- slide: 1 -->
Notes for slide 1.

<!-- slide: 2-4: Overview -->
Notes for slides 2 through 4.

<!-- slide: 5 -->
Notes for slide 5.
```

---

### 4. Sequential Delimiters (`---`, Marp / Slidev Style)

If you already have notes structured with horizontal rules (`---`, `***`, or `___`), `latex2pptx` sequentially assigns each section to Slide 1, Slide 2, Slide 3, etc.:

```markdown
Welcome notes for the opening slide.
- Highlight goal of session

---

Notes for the second slide.
- Explain problem statement

---

Notes for the third slide.
- Introduce solution
```

---

## Rich Content & Formatting

PowerPoint speaker notes text frames support multiple paragraphs and bullet points:
- Separate paragraphs with blank lines.
- Use bullet characters (`- `, `* `, `• `) or numbered lists (`1. `, `2. `) for speaking prompts.
- Subtitles in headings (e.g. `# Slide 1: Welcome`) are preserved at the top of the slide's notes frame.
