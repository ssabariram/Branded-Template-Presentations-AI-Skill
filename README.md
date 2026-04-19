# Branded Template Presentations — AI Skill

A reference implementation for generating **professional, brand-compliant PowerPoint presentations** from any corporate template using an AI agent — with full slide transitions and element animations, engineered from scratch without any paid library.

> **Companion blog post:** [The Weekend I Declared War on PowerPoint XML]([https://hai2.ai](https://medium.com/@sachinsabariram/the-weekend-i-declared-war-on-powerpoint-xml-and-what-was-built-when-the-dust-settled-e99cb6515a1f)) — the full story of why this exists and what it took to build it.

---

## What This Is

Most AI presentation tools fall into one of two camps:

- **From-scratch generators** (like `pptxgenjs`): creatively powerful, visually impressive, but can't follow an existing corporate template.
- **Template editors** (unpack → edit XML → repack): token-expensive, slow, and prone to corrupted output on anything beyond trivial edits.

This project takes a third path: a **content injection model** where the agent only produces a structured JSON content plan, and a deterministic Python injector handles all the template-aware rendering. The result is fast (seconds, not minutes), consistent (no visual QA loop), and works on models that don't support image input.

---

## How It Works

```
Agent → content_plan.json → injector script → branded .pptx
```

1. The agent writes a structured JSON file describing the slides (types, text, charts, animations, transitions).
2. The injector validates the plan, loads your template, and generates each slide by cloning the correct layout — never editing the template itself.
3. The output `.pptx` inherits all brand formatting (fonts, colours, backgrounds, logos, placeholder positions) from the template automatically.

Animations and transitions are injected via raw OOXML (`lxml`) since `python-pptx` has no public API for either. The XML structure matches what PowerPoint itself generates — no repair dialogs, no corrupted files.

---

## Repository Structure

```
.
├── SKILL.md                        # Agent skill definition — the agent reads this
├── assets/
│   └── <your_template>.pptx        # Your corporate PowerPoint template (add yours here)
├── scripts/
│   ├── <company_name>_injector.py  # Main orchestrator: reads JSON plan, generates .pptx
│   ├── animations.py               # Element animation injection (OOXML p:timing)
│   └── transitions.py              # Slide transition injection (OOXML p:transition)
└── README.md
```

---

## Adapting to Your Template

This is a reference implementation — it's designed to be adapted, not used as-is. Here's what you need to change:

### 1. Add your template file

Place your corporate `.pptx` template at `assets/<your_company>_template.pptx`.

### 2. Map your layouts

Open your template in PowerPoint and inspect its slide layouts (View → Slide Master). Note the index and name of each layout you want to support. Update the `LAYOUT_CATALOG` dictionary in the injector with your layouts:

```python
LAYOUT_CATALOG = {
    "COVER": {
        "layout_index": 0,          # index in prs.slide_layouts
        "layout_name": "Title Slide",
        "placeholders": {
            "title":    {"idx": 0, "max_chars_per_line": 25, "max_lines": 2},
            "subtitle": {"idx": 1, "max_chars_per_line": 40, "max_lines": 2},
        }
    },
    "CONTENT": { ... },
    # add as many as your template supports
}
```

The `max_chars_per_line` and `max_lines` values come from the physical size of the placeholder box at your template's font sizes — measure them by filling a placeholder with test text and seeing where it clips.

### 3. Update brand colours

Find the `CHART_COLORS` list in the injector and set your brand's hex values:

```python
CHART_COLORS = [
    RGBColor(0x00, 0x80, 0xBA),   # Primary blue
    RGBColor(0xFF, 0x99, 0x01),   # Accent orange
    # ...
]
```

### 4. Extract your icons (optional)

If your template has an embedded icon library, they live in `ppt/media/` inside the `.pptx` ZIP. Map their filenames to semantic names in `ICON_CATALOG` and the injector will be able to place them by name.

### 5. Update SKILL.md

Replace the `<company-name>` placeholders in `SKILL.md` with your company name and adjust the slide type descriptions, character limits, and examples to match your actual layouts.

---

## Running the Injector

```bash
pip install python-pptx lxml

python scripts/<company_name>_injector.py content_plan.json \
  --template assets/<your_template>.pptx \
  --output output/presentation.pptx
```

The injector validates the full content plan before writing any slides and prints warnings for anything that exceeds character limits or uses unknown slide types — so the agent can self-correct before you get a bad output.

---

## Content Plan Format

```json
{
  "slides": [
    {
      "type": "COVER",
      "title": "Q2 Business Review",
      "subtitle": "April 2026",
      "transition": {"type": "fade", "speed": "slow"},
      "animations": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",    "duration_ms": 800},
        {"shape": "subtitle", "effect": "fade",   "trigger": "after_prev", "delay_ms": 200, "duration_ms": 600}
      ]
    },
    {
      "type": "CONTENT",
      "title": "Key Findings",
      "bullets": [
        "Finding one with supporting detail",
        {"text": "Finding two", "sub_bullets": ["Sub-point A", "Sub-point B"]},
        "Finding three"
      ],
      "animations": [
        {"shape": "title", "effect": "fade", "trigger": "on_load",  "duration_ms": 400},
        {"shape": "body",  "effect": "wipe", "trigger": "on_click", "dir": "from_left",
         "duration_ms": 400, "text_build": "by_bullet"}
      ],
      "speaker_notes": "Pause here — let the audience read finding two before moving on."
    },
    {
      "type": "TWO_COLUMN",
      "title": "Build vs Buy",
      "left_title": "Build In-House",
      "left_bullets": ["Full control", "Higher upfront cost", "Slower time-to-value"],
      "right_title": "Buy / Partner",
      "right_bullets": ["Faster delivery", "Ongoing licensing", "Vendor dependency"]
    },
    {
      "type": "KPI",
      "title": "Program Highlights",
      "kpis": [
        {"value": "$4.2M", "label": "Annual Savings"},
        {"value": "99.9%", "label": "Uptime"},
        {"value": "3.2×",  "label": "ROI"}
      ]
    },
    {
      "type": "CHART",
      "title": "Revenue by Quarter",
      "chart": {
        "type": "column",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
          {"name": "2025", "values": [1.2, 1.4, 1.8, 2.1]},
          {"name": "2026", "values": [1.5, 1.9, 2.3, 2.8]}
        ]
      }
    }
  ]
}
```

### Supported Slide Types (adapt to your template)

The reference implementation ships with 17 slide types. Your catalog will depend on your template's layouts:

| Type | Content |
|------|---------|
| `COVER` | Title + subtitle, opening slide |
| `COVER_ALT` | Alternate cover layout |
| `COVER_FULL` | Bold full-background cover |
| `CHAPTER` | Chapter/section divider with subtitle |
| `SECTION_BLUE` | Section break, single title |
| `SECTION_GREY` | Section break variant |
| `CONTENT` | Title + bullet list |
| `CONTENT_SIDEBAR` | Title + bullets, wider body area |
| `TWO_COLUMN` | Title + two side-by-side columns |
| `QUOTE` | Large centred statement or stat |
| `CLOSING` | Fixed closing/thank-you slide |
| `CHART` | Title + full-width chart |
| `TABLE` | Title + styled data table |
| `CONTENT_CHART` | Split: bullets left, chart right |
| `CONTENT_TABLE` | Split: bullets left, table right |
| `KPI` | Title + 2–4 metric cards |
| `TIMELINE` | Title + horizontal milestone timeline |

---

## Animations

Element animations are injected as a `<p:timing>` block via raw OOXML — `python-pptx` has no public API for animations. The XML structure matches what PowerPoint's own animation engine generates (verified by diffing against real PowerPoint output).

**Available effects:**

| Class | Effects |
|-------|---------|
| Entrance | `fade`, `fly_in`, `float_in`, `wipe`, `zoom`, `dissolve`, `split`, `appear`, and more |
| Exit | `fade_out`, `fly_out`, `wipe_out`, `zoom_out`, `disappear` |
| Emphasis | `pulse`, `grow_shrink`, `spin`, `bold_flash` |

**Triggers:** `on_load` · `on_click` · `after_prev` · `with_prev`

**By-bullet builds:** Set `"text_build": "by_bullet"` on any body shape to reveal one bullet per click.

---

## Transitions

Slide transitions are injected as `<p:transition>` elements. PowerPoint 2010+ transitions are wrapped in `<mc:AlternateContent>` for backward compatibility with older renderers — all degrade gracefully to a plain fade.

**38 transitions supported across two namespace levels:**

- **ECMA-376 core (19):** `fade`, `cut`, `push`, `cover`, `pull`, `wipe`, `dissolve`, `split`, `zoom`, `wheel`, `blinds`, `checker`, `circle`, `diamond`, `plus`, `wedge`, `comb`, `randomBar`, `strips`
- **PowerPoint 2010+ (19):** `conveyor`, `doors`, `ferris`, `flip`, `flythrough`, `gallery`, `glitter`, `honeycomb`, `pan`, `prism`, `reveal`, `ripple`, `shred`, `switch`, `vortex`, `warp`, `window`, `flash`, `wheelReverse`

**Speed:** `slow` (~1500ms) · `med` (~800ms) · `fast` (~400ms)

---

## Using as an Agent Skill

The `SKILL.md` file is the agent's instruction set. Load it into your agent framework as a skill and the agent will:

1. Read the user's presentation request
2. Produce a `content_plan.json` following the schema
3. Run the injector script
4. Deliver the `.pptx`

The skill uses progressive disclosure — the agent doesn't need to load knowledge about animations or chart formatting until it needs them, keeping context lean.

---

## Dependencies

```bash
pip install python-pptx lxml
```

| Package | Purpose |
|---------|---------|
| `python-pptx` | Slide creation, chart/table/image injection, placeholder management |
| `lxml` | Raw XML manipulation for transitions and animations |

Python 3.8+ required.

---

## Background

This implementation grew out of a frustrating experience with the standard AI presentation approaches — template editing burns tokens and produces corruption; from-scratch generation can't follow brand guidelines. The full story, including the detailed journey of implementing animations against the ECMA-376 spec, is in the [companion blog post](https://hai2.ai/the-weekend-i-declared-war-on-powerpoint-xml).

The engineering notes in `docs/` (if you've cloned from the full repo) cover the architecture in depth, the specific XML bugs encountered implementing animations, and the comparison between the three approaches in detail.

---

## License

MIT
