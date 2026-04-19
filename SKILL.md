---
name: company-template-presentations
description: "Generate professional PowerPoint presentations using the <company-name> corporate template. Produces a structured JSON content plan that is injected into the template via python-pptx, preserving all brand formatting, backgrounds, fonts, and layout integrity. Supports text slides, charts, tables, KPI cards, timelines, icons, and speaker notes. Use when users request <company-name>-branded slide decks, presentations, or pitch decks."
version: "3.0"
---

# <company-name> Template Presentation Skill

## Overview

This skill generates <company-name>-branded PowerPoint presentations by:
1. **You** create a structured JSON content plan based on the user's request
2. **The injector script** reads the plan and produces a `.pptx` using the official <company-name> template

All formatting (fonts, colors, backgrounds, logos, layout positioning) is inherited from the template. You only need to provide the text content within strict character limits.

**v3.0 adds:** Charts (8 types), tables, KPI metric cards, timelines, 95 template-embedded icons, and speaker notes — all fully backward compatible with existing content plans.

## Prerequisites

Before generating a presentation:

1. **Install python-pptx**: `pip install python-pptx`
2. **Verify template exists**: The template file must be at `assets/<company-name>_template.pptx` relative to this skill's directory
3. **Verify injector exists**: The script must be at `scripts/<company-name>_injector.py` relative to this skill's directory

## How It Works

### Step 1: Create a Content Plan (JSON)

Based on the user's request, create a JSON file with a `slides` array. Each slide has a `type` and the text fields appropriate for that type.

### Step 2: Run the Injector

```bash
python scripts/<company-name>_injector.py content_plan.json \
  --template assets/<company-name>_template.pptx \
  --output output/presentation.pptx
```

The script will:
- Validate content against character limits (prints warnings if exceeded)
- Validate chart/table/KPI/timeline specifications
- Load the <company-name> template
- Remove all template slides (keeping layouts)
- Inject new slides from the content plan
- Save the output `.pptx`

### Step 3: Deliver

Copy the output `.pptx` to the user's outputs directory.

---

## Slide Type Catalog <Change as per your template>

There are **17 slide types** available. Each has strict content limits that MUST be respected for visual quality.

### COVER

**Purpose**: Opening/cover slide for the presentation.
**Visual**: Dark blue background with city skyline photo on right, <company-name> logo top-left.

| Field | Max Chars/Line | Max Lines | Notes |
|-------|---------------|-----------|-------|
| `title` | 20 | 2 | Large white bold text. Use `\n` for line breaks. |
| `subtitle` | 35 | 2 | Smaller white text below an orange dash. |

```json
{
  "type": "COVER",
  "title": "Cloud Migration\nStrategy 2026",
  "subtitle": "<company-name> Digital Solutions"
}
```

### COVER_ALT

**Purpose**: Alternate cover with people/conference photo. More professional look.
**Visual**: Blue background with people/conference photo on right.

| Field | Max Chars/Line | Max Lines | Notes |
|-------|---------------|-----------|-------|
| `title` | 20 | 3 | Allows one extra line vs COVER. |
| `subtitle` | 45 | 2 | Wider subtitle area. |

```json
{
  "type": "COVER_ALT",
  "title": "Digital Innovation\nTransforming\nEnterprise",
  "subtitle": "A Comprehensive Strategy for Modern Business"
}
```

### COVER_FULL

**Purpose**: Bold statement cover with maximum visual impact.
**Visual**: Full blue background, very large centered title, small brand icons top-right.

| Field | Max Chars/Line | Max Lines | Notes |
|-------|---------------|-----------|-------|
| `title` | 30 | 2 | Centered, large and bold. |
| `subtitle` | 45 | 2 | Below title. |

```json
{
  "type": "COVER_FULL",
  "title": "The Future of Work Is Now",
  "subtitle": "<company-name>'s Vision for Enterprise Technology"
}
```

### CHAPTER

**Purpose**: Chapter divider. Introduces a new major topic.
**Visual**: White background, city photo in top-right corner, title on left, orange dash + subtitle below.

| Field | Max Chars/Line | Max Lines | Notes |
|-------|---------------|-----------|-------|
| `title` | 18 | 2 | Keep short: 1-2 words per line like "Executive\nSummary". |
| `subtitle` | 35 | 1 | Brief caption. Single line only. |

```json
{
  "type": "CHAPTER",
  "title": "Executive\nSummary",
  "subtitle": "Strategic overview and goals"
}
```

### SECTION_BLUE

**Purpose**: Section divider within a chapter.
**Visual**: Blue horizontal band across middle of slide, title left, orange icon placeholder right.

| Field | Max Chars/Line | Max Lines | Notes |
|-------|---------------|-----------|-------|
| `title` | 17 | 2 | White text on blue. Keep short: 2-4 words. |

```json
{
  "type": "SECTION_BLUE",
  "title": "Migration\nApproach"
}
```

### SECTION_GREY

**Purpose**: Alternate section divider for visual variety.
**Visual**: Grey horizontal band across middle of slide (same layout as SECTION_BLUE but grey).

| Field | Max Chars/Line | Max Lines | Notes |
|-------|---------------|-----------|-------|
| `title` | 17 | 2 | Same constraints as SECTION_BLUE. |

```json
{
  "type": "SECTION_GREY",
  "title": "Current State\nAssessment"
}
```

### CONTENT

**Purpose**: PRIMARY content slide. Use for bullet points, text, lists. This is the workhorse.
**Visual**: White background, blue title bar at top with orange chevron, large content area below, footer with <company-name> logo.

| Field | Max Chars/Line | Max Lines | Max Bullets | Notes |
|-------|---------------|-----------|-------------|-------|
| `title` | 30 | 1 | - | Single line only! |
| `bullets` | 55 | 7 | 6 | Multi-level supported. |
| `body_text` | 55 | 7 | - | Alternative to bullets. Paragraph text. |
| `chart` | - | - | - | Optional. If provided WITHOUT bullets, replaces body with a full-width chart. |

**Bullet format** — each item can be a string or an object with sub-bullets:
```json
{
  "type": "CONTENT",
  "title": "Why Migrate to the Cloud?",
  "bullets": [
    "Reduce infrastructure costs by 40%",
    "Improve scalability and agility",
    {
      "text": "Enable digital transformation",
      "sub_bullets": [
        "AI/ML workloads",
        "Real-time analytics"
      ]
    },
    "Accelerate time to market"
  ]
}
```

**Chart overlay** (use instead of bullets for a full-width chart):
```json
{
  "type": "CONTENT",
  "title": "Revenue Growth",
  "chart": {
    "type": "line",
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [{"name": "Revenue", "values": [120, 145, 160, 195]}]
  }
}
```

**Guidelines for CONTENT slides:**
- 5-6 bullets max for readability (fewer is better)
- Each bullet should fit on one line (under 55 chars)
- Sub-bullets count toward the 7-line visual limit
- A bullet with 2 sub-bullets uses 3 lines of space
- When using `body_text`, do NOT also provide `bullets`
- When using `chart`, do NOT also provide `bullets` (use CONTENT_CHART for side-by-side)

### CONTENT_SIDEBAR

**Purpose**: Text-heavy alternate content layout with more space.
**Visual**: White background, vertical orange line on left as sidebar accent. Content takes nearly full slide width.

| Field | Max Chars/Line | Max Lines | Max Bullets | Notes |
|-------|---------------|-----------|-------------|-------|
| `title` | 15 | 1 | - | Rotated vertically along left sidebar. Keep very short. |
| `bullets` | 55 | 10 | 8 | Larger content area than CONTENT. |
| `body_text` | 55 | 10 | - | Alternative to bullets. |

```json
{
  "type": "CONTENT_SIDEBAR",
  "title": "Deep Analysis",
  "bullets": [
    "Infrastructure assessment reveals 73% legacy systems",
    "Security audit identified 12 critical vulnerabilities",
    {
      "text": "Cost analysis by department",
      "sub_bullets": [
        "Engineering: $2.4M annual",
        "Operations: $1.8M annual"
      ]
    },
    "Compliance gaps in three frameworks",
    "Integration debt across 47 services"
  ]
}
```

### TWO_COLUMN

**Purpose**: Side-by-side comparisons, pros/cons, before/after.
**Visual**: White background, blue title bar at top, two equal content areas side by side.

| Field | Max Chars/Line | Max Lines | Max Bullets | Notes |
|-------|---------------|-----------|-------------|-------|
| `title` | 30 | 1 | - | Single line only. |
| `left_bullets` | 25 | 7 | 5 | Left column (slightly narrower). |
| `right_bullets` | 28 | 7 | 5 | Right column. |
| `left_chart` | - | - | - | Optional. Replaces left_bullets with a chart. |
| `right_chart` | - | - | - | Optional. Replaces right_bullets with a chart. |

```json
{
  "type": "TWO_COLUMN",
  "title": "Current vs. Target State",
  "left_bullets": [
    "On-premise data centers",
    "Manual provisioning",
    "3-month release cycles",
    "Siloed teams"
  ],
  "right_bullets": [
    "Multi-cloud architecture",
    "Infrastructure as Code",
    "Weekly deployments",
    "Cross-functional DevOps"
  ]
}
```

**TWO_COLUMN with chart in one column:**
```json
{
  "type": "TWO_COLUMN",
  "title": "Cost Breakdown",
  "left_bullets": [
    "Infrastructure: 40%",
    "Personnel: 35%",
    "Licensing: 15%",
    "Other: 10%"
  ],
  "right_chart": {
    "type": "pie",
    "categories": ["Infra", "Personnel", "License", "Other"],
    "series": [{"name": "Cost", "values": [40, 35, 15, 10]}]
  }
}
```

**Guidelines for TWO_COLUMN:**
- Keep bullets shorter than CONTENT (25-28 chars vs 55)
- Match bullet counts across columns when possible
- 4-5 bullets per column is the sweet spot
- Do not provide both `left_chart` and `left_bullets` for the same column

### QUOTE

**Purpose**: Key quotes, statistics, or highlight statements. Maximum visual impact.
**Visual**: Blue background with decorative shapes, large centered text area.

| Field | Max Chars/Line | Max Lines | Notes |
|-------|---------------|-----------|-------|
| `title` | 30 | 3 | The title IS the content. No body text. Use `\n` for line breaks. |

```json
{
  "type": "QUOTE",
  "title": "93% of enterprises now\nhave a multi-cloud\nstrategy"
}
```

**Guidelines for QUOTE:**
- Use for impactful one-liners, key statistics, or memorable quotes
- Break long quotes across 2-3 lines using `\n`
- No attribution field — include source in the text if needed

### CLOSING

**Purpose**: Final "Thank You" slide with contact info.
**Visual**: City panorama photo top half, "Thank You" text, contact info area, <company-name> logo.

**No fields.** This slide is fully pre-built from the template.

```json
{
  "type": "CLOSING"
}
```

---

### CHART

**Purpose**: Full-width data visualization replacing the body area.
**Visual**: White background, blue title bar at top, chart occupies the entire body area.

| Field | Required | Notes |
|-------|----------|-------|
| `title` | Y | 30 chars max, single line. |
| `chart` | Y | Chart specification object. See [Chart Specification](#chart-specification). |

```json
{
  "type": "CHART",
  "title": "Quarterly Revenue Growth",
  "chart": {
    "type": "column",
    "categories": ["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025"],
    "series": [
      {"name": "Revenue", "values": [4200, 4800, 5100, 5900]},
      {"name": "Target", "values": [4000, 4500, 5000, 5500]}
    ],
    "title": "Revenue vs Target ($K)",
    "show_legend": true,
    "show_data_labels": false,
    "show_gridlines": true
  },
  "speaker_notes": "Revenue exceeded targets by 5-10% each quarter."
}
```

### TABLE

**Purpose**: Full-width styled data table replacing the body area.
**Visual**: White background, blue title bar at top, branded table with <company-name>-blue header and alternating row stripes.

| Field | Required | Notes |
|-------|----------|-------|
| `title` | Y | 30 chars max, single line. |
| `table` | Y | Table specification object. See [Table Specification](#table-specification). |

```json
{
  "type": "TABLE",
  "title": "Migration Timeline",
  "table": {
    "headers": ["Phase", "Duration", "Key Deliverable", "Status"],
    "rows": [
      ["Assessment", "4 weeks", "Current state report", "Complete"],
      ["Planning", "6 weeks", "Migration roadmap", "In Progress"],
      ["Migration", "12 weeks", "Workload migration", "Planned"],
      ["Optimization", "8 weeks", "Performance tuning", "Planned"]
    ]
  }
}
```

### CONTENT_CHART

**Purpose**: Split layout with bullet points on the left and a chart on the right.
**Visual**: White background, blue title bar, left half has bullets, right half has chart.

| Field | Required | Notes |
|-------|----------|-------|
| `title` | Y | 30 chars max, single line. |
| `bullets` | Y | Left side, 40 chars/line max, 6 bullets max. |
| `chart` | Y | Right side chart. See [Chart Specification](#chart-specification). |

```json
{
  "type": "CONTENT_CHART",
  "title": "Cloud Adoption Trends",
  "bullets": [
    "78% of enterprises use cloud",
    "Multi-cloud is the norm",
    "Security is top concern",
    "Cost optimization is #2 priority"
  ],
  "chart": {
    "type": "doughnut",
    "categories": ["AWS", "Azure", "GCP", "Other"],
    "series": [{"name": "Market Share", "values": [32, 23, 11, 34]}]
  }
}
```

### CONTENT_TABLE

**Purpose**: Split layout with bullet points on the left and a table on the right.
**Visual**: White background, blue title bar, left half has bullets, right half has table.

| Field | Required | Notes |
|-------|----------|-------|
| `title` | Y | 30 chars max, single line. |
| `bullets` | Y | Left side, 40 chars/line max, 6 bullets max. |
| `table` | Y | Right side table. See [Table Specification](#table-specification). |

```json
{
  "type": "CONTENT_TABLE",
  "title": "Team Structure",
  "bullets": [
    "Cross-functional DevOps model",
    "Dedicated security liaison",
    "Shared services platform team",
    "Business unit representatives"
  ],
  "table": {
    "headers": ["Role", "Count"],
    "rows": [
      ["Cloud Architects", "4"],
      ["DevOps Engineers", "8"],
      ["Security", "3"],
      ["Project Managers", "2"]
    ]
  }
}
```

### KPI

**Purpose**: 2-4 large metric cards for key performance indicators.
**Visual**: White background, blue title bar, colored rounded-rectangle cards evenly distributed across the body area. Each card shows a large value and a label in white text.

| Field | Required | Notes |
|-------|----------|-------|
| `title` | Y | 30 chars max, single line. |
| `kpis` | Y | Array of 2-4 KPI objects. See below. |

Each KPI object:
| Field | Required | Limits | Notes |
|-------|----------|--------|-------|
| `value` | Y | 8 chars max | The metric value, e.g., "$4.2M", "99.9%", "47" |
| `label` | Y | 20 chars max | Description, e.g., "Annual Revenue", "Uptime" |
| `color` | N | Hex color | Default: <company-name> Blue `#0080BA` |

```json
{
  "type": "KPI",
  "title": "Key Performance Indicators",
  "kpis": [
    {"value": "$4.2M", "label": "Annual Savings", "color": "#0080BA"},
    {"value": "99.99%", "label": "Uptime SLA", "color": "#4CAF50"},
    {"value": "47%", "label": "Faster Deploys", "color": "#FF9901"},
    {"value": "Zero", "label": "Critical Incidents", "color": "#00447A"}
  ]
}
```

**Guidelines for KPI:**
- 2-4 cards recommended (3 is ideal for visual balance)
- Keep values short (numbers, percentages, currency)
- Keep labels concise — they should be scannable
- Use contrasting brand colors for visual variety

### TIMELINE

**Purpose**: Horizontal timeline showing project milestones or phases.
**Visual**: White background, blue title bar, horizontal grey line with colored circle nodes, labels above each node, descriptions below.

| Field | Required | Notes |
|-------|----------|-------|
| `title` | Y | 30 chars max, single line. |
| `milestones` | Y | Array of 2-6 milestone objects. See below. |

Each milestone object:
| Field | Required | Notes |
|-------|----------|-------|
| `label` | Y | Phase or milestone name (14pt bold, colored) |
| `description` | N | Details below the node (11pt grey) |
| `color` | N | Hex color for the node and label. Default: `#0080BA` |

```json
{
  "type": "TIMELINE",
  "title": "Migration Roadmap",
  "milestones": [
    {"label": "Discovery", "description": "Assess current state", "color": "#0080BA"},
    {"label": "Planning", "description": "Design target arch", "color": "#FF9901"},
    {"label": "Migration", "description": "Move workloads", "color": "#00447A"},
    {"label": "Optimize", "description": "Tune performance", "color": "#4CAF50"}
  ]
}
```

**Guidelines for TIMELINE:**
- 3-5 milestones is the visual sweet spot
- Keep labels to 1-2 words
- Keep descriptions under 20 characters for clean layout
- Use different colors to distinguish phases

---

## Chart Specification

Charts are available on CHART, CONTENT_CHART, CONTENT (without bullets), and TWO_COLUMN (as left_chart/right_chart) slides.

### Chart Object Fields

| Field | Required | Default | Notes |
|-------|----------|---------|-------|
| `type` | Y | - | Chart type string. See table below. |
| `categories` | Y | - | Array of category labels (x-axis). |
| `series` | Y | - | Array of series objects: `{"name": "str", "values": [num]}`. |
| `title` | N | none | Chart title displayed above the chart area. |
| `show_legend` | N | `true` | Show/hide the legend. |
| `show_data_labels` | N | `false` | Show values on each data point. |
| `show_gridlines` | N | `true` | Show horizontal gridlines. |

### Supported Chart Types

| Type | Value | Best For |
|------|-------|----------|
| Clustered Column | `"column"` | Comparing values across categories |
| Clustered Bar | `"bar"` | Horizontal comparison (long category names) |
| Line with Markers | `"line"` | Trends over time |
| Pie | `"pie"` | Parts of a whole (max 8 categories) |
| Doughnut | `"doughnut"` | Parts of a whole, modern look (max 8 categories) |
| Stacked Column | `"stacked_column"` | Composition across categories |
| Stacked Bar | `"stacked_bar"` | Horizontal composition |
| Area | `"area"` | Volume/trend emphasis |

### Chart Styling

All charts are automatically styled with <company-name> brand colors:
- **Series colors cycle**: Blue (#0080BA), Orange (#FF9901), Dark Blue (#00447A), Green (#4CAF50), Light Blue (#5BC0DE), Grey (#989998)
- **Pie/Doughnut**: Colors applied per data point; data labels show category + percentage
- **Fonts**: Helvetica Neue — 14pt bold for chart title, 10pt for axes and labels
- **Gridlines**: Light grey (#D0D0D0) when enabled

### Chart Examples

**Column chart with multiple series:**
```json
{
  "type": "column",
  "categories": ["North", "South", "East", "West"],
  "series": [
    {"name": "2024", "values": [120, 95, 140, 110]},
    {"name": "2025", "values": [145, 110, 165, 130]}
  ],
  "title": "Revenue by Region ($K)",
  "show_legend": true,
  "show_data_labels": true
}
```

**Pie chart:**
```json
{
  "type": "pie",
  "categories": ["Cloud", "On-Prem", "Hybrid"],
  "series": [{"name": "Infrastructure", "values": [55, 25, 20]}]
}
```

**Line chart (trend):**
```json
{
  "type": "line",
  "categories": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
  "series": [
    {"name": "Incidents", "values": [12, 9, 7, 5, 3, 2]},
    {"name": "Target", "values": [10, 8, 6, 5, 4, 3]}
  ],
  "show_gridlines": true
}
```

---

## Table Specification

Tables are available on TABLE and CONTENT_TABLE slides.

### Table Object Fields

| Field | Required | Default | Notes |
|-------|----------|---------|-------|
| `headers` | Y | - | Array of column header strings. |
| `rows` | Y | - | Array of row arrays. Each row is an array of cell strings. |
| `header_color` | N | `"#0080BA"` | Hex color for header row background. |
| `stripe_color` | N | `"#F0F6FA"` | Hex color for alternating row stripes. |
| `font_size` | N | `12` | Font size in points. |
| `col_widths` | N | even | Array of column widths in inches. Must match header count. |

### Table Styling

- **Header row**: Colored background (default <company-name> Blue) with white bold text
- **Data rows**: Alternating stripe (light blue-grey `#F0F6FA`) and white
- **Font**: Helvetica Neue
- **Cell padding**: Comfortable margins for readability

### Table Limits

| Limit | Recommended | Hard Max |
|-------|------------|----------|
| Columns | 4-5 | 6 |
| Rows | 5-8 | 10 |

**Each row array must have the same number of cells as there are headers.** The validator will warn on mismatches.

---

## Icons

The template includes **95 professional brand icons** (SVG + PNG pairs) that can be placed on any slide type. Icons are referenced by name and positioned using presets or custom coordinates.

### Icon Usage

Add an `"icons"` array to any slide:

```json
{
  "type": "CONTENT",
  "title": "Cloud Strategy",
  "bullets": ["..."],
  "icons": [
    {"name": "cloud", "position": "top_right"},
    {"name": "cybersecurity", "position": "top_left", "size": 0.8}
  ]
}
```

### Icon Positions

| Preset | Location | Default Size |
|--------|----------|-------------|
| `"top_right"` | Top-right corner (11.5", 0.2") | 1.0" |
| `"top_left"` | Top-left corner (0.3", 0.2") | 1.0" |
| `"before_title"` | Left of title area (0.3", 1.55") | 0.8" |

**Custom position:**
```json
{"name": "cloud", "position": {"x": 10.0, "y": 3.0}, "size": 1.5}
```

### Icon Catalog (95 icons)

**Solution Category Icons:**
`managed_services_wheel`, `cyber_security`, `digital`, `modern_networking`, `modern_platforms`, `total_experience`, `cost_optimization`

**Core Business Icons:**
`cloud`, `infrastructure_modernization`, `workforce_transformation`, `cybersecurity`, `lifecycle_services`, `digital_transformation`, `application_stack`, `application_development`, `services`, `devops`, `datacenter`, `networking`, `spotlight`, `wireless`, `collaboration`, `download`, `systems`, `managed_services`, `multi_cloud`, `location`, `managers`, `folder`, `firewall`, `employees`, `email`, `computer`, `global`, `data_analytics`

**Technology & Security Icons:**
`vsoc`, `modern_device_management`, `dvi`, `data`, `data_security`, `software_automation`, `computer_link`, `computer_chip`, `incident_response`, `cloud_management`, `upload`, `handshake`, `consulting_services`, `messaging`, `iot_ot`, `threat_detection`, `security_analytics`, `threat_hunting`, `threat_intelligence`, `laptop`, `apps_infrastructure`, `integrate_public_clouds`, `devops_cicd`, `transform_networking`

**Workspace & Infrastructure Icons:**
`modern_workplace`, `empower_digital_workspaces`, `lifecycle_management`, `software_defined_infrastructure`, `digital_workspace`, `servers`, `desktop_transformation`, `managed_services_alt`

**Data Center & Storage Icons:**
`all_flash_storage`, `resource`, `full_stack_solutions`, `protect`, `hyper_converged`, `virtual_desktop`, `access`, `modernize_data_center`, `process`, `infrastructure`, `data_alt`, `app`

**Financial & Billing Icons:**
`predictable_expense`, `variable_opex`, `utility_billing`, `metering_agreement`

**Security & Compliance Icons:**
`identity`, `scorpion_secops`, `digital_solution`, `zero_trust_protect`, `zero_trust_visibility`, `zero_trust_architecture`, `zero_trust_policy`, `zero_trust_monitor`, `mobile_device`, `ppm_services`, `secure`, `unified`

---

## Element Animations

Every slide supports an optional `"animations"` array that controls how individual shapes animate **within** the slide (after it has appeared). If omitted, the injector applies smart defaults for the slide type. Provide `"animations": []` (empty array) to explicitly suppress all animations.

Animations execute in the **order listed**. The injector writes a full `<p:timing>` block that PowerPoint / LibreOffice Impress reads natively.

### Animation Object Fields

| Field | Required | Default | Notes |
|-------|----------|---------|-------|
| `shape` | Y | — | Which shape to animate. See [Shape Targeting](#shape-targeting). |
| `effect` | Y | — | Animation effect name from the catalog. |
| `trigger` | N | `"on_click"` | When to start: `"on_click"` \| `"after_prev"` \| `"with_prev"` \| `"on_load"` |
| `dir` | N | — | Direction qualifier. Valid values depend on the effect. |
| `duration_ms` | N | `500` | Animation duration in milliseconds (100–5000). |
| `delay_ms` | N | `0` | Delay after the trigger fires (milliseconds). |
| `text_build` | N | — | `"by_bullet"` animates each bullet paragraph individually. `"all_at_once"` (default) animates the whole text box. Only applies to body/text shapes. |

```json
{
  "type": "CONTENT",
  "title": "Our Cloud Strategy",
  "bullets": ["Reduce costs", "Improve agility", "Enhance security"],
  "animations": [
    {"shape": "title", "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
    {"shape": "body",  "effect": "wipe",   "trigger": "on_click", "dir": "from_left",
     "duration_ms": 400, "text_build": "by_bullet"}
  ]
}
```

---

### Shape Targeting

The `shape` field accepts:

| Value | Targets |
|-------|---------|
| `"title"` | The title placeholder (placeholder idx=0) |
| `"body"` | The body/content placeholder (placeholder idx=1) — also left column on TWO_COLUMN |
| `"subtitle"` | The subtitle placeholder on cover slides (same as `"body"`) |
| `"right"` | Right column placeholder on TWO_COLUMN slides (placeholder idx=2) |
| `<integer>` | A specific shape by its PowerPoint shape ID |

**Note on free shapes**: KPI cards, timeline nodes, chart objects, and table objects are injected as free shapes (not placeholders) — they cannot be reliably targeted by role name. Only `"title"` is safe for CHART, TABLE, KPI, and TIMELINE slides. To animate free shapes, use a numeric shape ID.

**TWO_COLUMN note**: Both `"body"` (left column) and `"right"` (right column) must be animated separately. The smart default animates both — `body` wipes from the left, `right` wipes from the right — so both columns reveal symmetrically on successive clicks. Always include both when customizing TWO_COLUMN animations.

---

### Animation Effect Catalog

#### Entrance Effects (element appears)

| Effect | Dir values | Best For | Notes |
|--------|-----------|----------|-------|
| `appear` | — | Rarely used | Instant appear, no animation. Use only when zero motion is intentional. |
| `fade` | — | Titles, subtitles, charts, tables | Smooth cross-fade. The safest universal choice — never distracting. |
| `fly_in` | `from_bottom` `from_top` `from_left` `from_right` | Body text, subtitles | Element flies in from an edge. Energetic — use for emphasis. |
| `float_in` | `from_bottom` `from_top` | Subtitles, supporting text | Gentle drift. Softer than fly_in; good for cover/closing. |
| `wipe` | `from_bottom` `from_top` `from_left` `from_right` | Body text, two-column content | Edge wipe reveal. Feels purposeful; reinforces reading direction. |
| `zoom` | `in` `out` `in_slide_center` | Cover titles, KPI statements | Zoom entrance. High impact — use for bold moments only. |
| `dissolve` | — | Quote slides, atmospheric reveals | Pixel dissolve. Feels organic; good for pause slides. |
| `split` | `horizontal_in` `vertical_in` `horizontal_out` `vertical_out` | Two-column, comparison slides | Barn-door split. Visual metaphor for duality/comparison. |
| `bounce` | — | Avoid in corporate decks | Informal — use only in internal/casual presentations. |
| `blinds` | `horizontal` `vertical` | Avoid overuse | Venetian blinds — feels dated; use only for a deliberate retro look. |

*(Other iris effects — `box`, `circle`, `diamond`, `plus`, `wedge`, `wheel`, `strips`, `checker`, `random_bars`, `swivel` — are available but rarely appropriate in professional decks. Omit unless a specific design rationale exists.)*

#### Exit Effects (element disappears)

| Effect | Dir values | Notes |
|--------|-----------|-------|
| `disappear` | — | Instant disappear |
| `fade_out` | — | Cross-fade out |
| `fly_out` | `to_bottom` `to_top` `to_left` `to_right` | Flies off-screen |
| `wipe_out` | `to_bottom` `to_top` `to_left` `to_right` | Wipe out |
| `zoom_out` | — | Zoom out exit |

Exit effects are rarely needed in slide-by-slide presentations. Only use them if a shape needs to disappear before the next click (e.g. swapping a placeholder for a chart reveal).

#### Emphasis Effects (element draws attention, stays visible)

| Effect | Dir values | Best For | Notes |
|--------|-----------|----------|-------|
| `pulse` | — | Icons, KPI values | Quick scale-up pulse — good for calling out a single number. |
| `grow_shrink` | — | Icons, accent shapes | Larger grow/shrink — more dramatic than pulse. |
| `spin` | `clockwise` `counter_clockwise` | Icons only | Full rotation — NEVER use on text. |
| `bold_flash` | — | Key terms | Bold flash on text — use sparingly (max once per deck). |

---

### Smart Defaults (applied when `"animations"` is omitted)

These defaults are carefully calibrated for professional presentations. Override only when you have a specific reason.

| Slide Type | Default Animation Sequence |
|---|---|
| COVER | title: `fade` on_load (800ms) → subtitle: `fade` after_prev +200ms (600ms) |
| COVER_ALT | title: `fly_in from_bottom` on_load (700ms) → subtitle: `fade` after_prev +100ms (600ms) |
| COVER_FULL | title: `zoom in` on_load (800ms) → subtitle: `fade` after_prev +200ms (600ms) |
| CHAPTER | title: `fly_in from_left` on_load (600ms) → subtitle: `fade` after_prev +100ms (500ms) |
| SECTION_BLUE | title: `wipe from_left` on_load (500ms) |
| SECTION_GREY | title: `wipe from_left` on_load (500ms) |
| CONTENT | title: `fade` on_load (400ms) → body: `wipe from_left` on_click by_bullet (400ms each) |
| CONTENT_SIDEBAR | title: `fade` on_load (400ms) → body: `fly_in from_left` on_click by_bullet (400ms each) |
| TWO_COLUMN | title: `fade` on_load (400ms) → body/left: `wipe from_left` on_click (400ms) → right: `wipe from_right` on_click (400ms) |
| QUOTE | title: `dissolve` on_load (900ms) |
| CLOSING | title: `fade` on_load (800ms) |
| CHART | title: `fade` on_load (400ms) |
| TABLE | title: `fade` on_load (400ms) |
| CONTENT_CHART | title: `fade` on_load (400ms) → body: `fly_in from_left` on_click (400ms) |
| CONTENT_TABLE | title: `fade` on_load (400ms) → body: `fly_in from_left` on_click (400ms) |
| KPI | title: `fade` on_load (400ms) |
| TIMELINE | title: `fade` on_load (400ms) |

---

### Animation Design Principles

These are rules, not suggestions. Follow them exactly unless the user gives a specific override reason.

#### Rule 1: Title always animates first, always on_load

The title must appear automatically when the slide opens — never on a click. The audience needs context before content. Never put body/bullet animations before the title, and never use `on_click` for a title.

```
✓ title: fade on_load → body: wipe on_click
✗ body: wipe on_click → title: fade on_click   (body first — wrong)
✗ title: fade on_click                          (title click-gated — wrong)
```

#### Rule 2: Use `on_click` for body content, `after_prev` for flowing auto-sequences

- **`on_load`** — title, section heading, cover subtitle (things that set context instantly)
- **`on_click`** — bullets, right columns, any content the presenter wants to pace
- **`after_prev` + small `delay_ms`** — subtitle that flows after the title automatically (keep delays 100–300 ms); use for at most 2 elements in a sequence, never more than 3

```
✓ title: fade on_load → subtitle: fade after_prev delay=200ms   (smooth auto flow)
✗ title: fade on_load → body: wipe after_prev → right: wipe after_prev → chart: zoom after_prev
  (too many auto-chained elements — robotic, presenter loses control)
```

#### Rule 3: Choose the effect based on the shape's role, not variety

Don't mix effects randomly to "make it interesting." Variety reads as noise. Instead:

| Shape Role | Use This Effect | Why |
|---|---|---|
| Slide title | `fade` | Safe, universal, never distracts |
| Section/chapter title | `fly_in from_left` or `wipe from_left` | Reinforces forward momentum |
| Cover/closing title | `fade` or `zoom in` | Cinematic weight for structural slides |
| Quote/statement | `dissolve` | Organic feel matches the pause-and-reflect moment |
| Body bullets | `wipe from_left` | Mirrors reading direction; feels purposeful |
| Second column (TWO_COLUMN right) | `wipe from_right` | Mirror the left column direction — symmetric reveal |
| Subtitle on cover | `fade` or `float_in from_bottom` | Gentle, doesn't steal from the title |

#### Rule 4: `by_bullet` only for sequential narrative, not reference lists

Use `"text_build": "by_bullet"` when bullets represent steps, arguments, or reveals the presenter will explain one by one. **Do not use** when bullets are a reference list the audience should scan all at once (e.g. a feature comparison, a list of attendees, prerequisites).

```
✓ "bullets": ["Step 1: Assess", "Step 2: Plan", "Step 3: Migrate"]  → by_bullet  (sequential)
✗ "bullets": ["Python", "Go", "Rust", "Java", "C++"]               → by_bullet  (reference — all at once)
```

#### Rule 5: Duration discipline

Keep animations short and professional:

| Element | Duration |
|---|---|
| Title (any type) | 400–800 ms |
| Subtitle / supporting text | 400–600 ms |
| Body bullet (per bullet) | 300–500 ms |
| Second column reveal | 300–500 ms |
| Cover title (cinematic feel) | 600–900 ms |
| Quote/dissolve | 700–1000 ms |
| Emphasis (pulse, grow) | 300–400 ms |

**Never exceed 1000 ms** for any animation in a business presentation. Durations over 1 second make the deck feel sluggish.

#### Rule 6: TWO_COLUMN requires symmetric column animation

Both columns must always be animated. The left column (`body`) wipes from the left; the right column (`right`) wipes from the right. They animate on separate clicks so the presenter can discuss each column. This is the default — never override it to omit the right column.

```json
{"shape": "body",  "effect": "wipe", "trigger": "on_click", "dir": "from_left",  "duration_ms": 400},
{"shape": "right", "effect": "wipe", "trigger": "on_click", "dir": "from_right", "duration_ms": 400}
```

#### Rule 7: Suppress animations only for pure reference slides

Use `"animations": []` only for slides that function as reference material mid-deck (e.g. a compliance table, an appendix). Never suppress on featured content slides.

```
✓ "animations": []  on TABLE slide containing reference data
✗ "animations": []  on CONTENT slide with main argument
```

#### Rule 8: Never use novelty effects in professional decks

These effects are available but inappropriate for business/enterprise presentations:
- `bounce` — informal
- `spin` on text — always wrong
- `blinds`, `checker`, `random_bars`, `strips` — feel dated
- `wheel`, `wedge`, `plus`, `diamond`, `circle` — distracting iris shapes
- Chaining more than 3 `after_prev` elements in sequence

### Animation Examples

**Cover slide (default behavior — no override needed for most decks):**
```json
{
  "type": "COVER",
  "title": "Cloud Migration\nStrategy 2026",
  "subtitle": "<company-name> Digital Solutions"
}
```

**Cover slide with custom cinematic feel:**
```json
{
  "type": "COVER",
  "title": "Cloud Migration\nStrategy 2026",
  "subtitle": "<company-name> Digital Solutions",
  "animations": [
    {"shape": "title",    "effect": "fade",   "trigger": "on_load",   "duration_ms": 800},
    {"shape": "subtitle", "effect": "float_in","trigger": "after_prev","dir": "from_bottom",
     "duration_ms": 600, "delay_ms": 300}
  ]
}
```

**Content slide with bullet-by-bullet reveal (most common pattern):**
```json
{
  "type": "CONTENT",
  "title": "Why Migrate to Cloud?",
  "bullets": ["Reduce costs by 40%", "Scale on demand", "Improve security posture"],
  "animations": [
    {"shape": "title", "effect": "fade", "trigger": "on_load",  "duration_ms": 400},
    {"shape": "body",  "effect": "wipe", "trigger": "on_click", "dir": "from_left",
     "duration_ms": 400, "text_build": "by_bullet"}
  ]
}
```

**TWO_COLUMN slide with symmetric column reveal:**
```json
{
  "type": "TWO_COLUMN",
  "title": "Build vs Buy",
  "left_title": "Build In-House",
  "left_bullets": ["Full control", "High upfront cost"],
  "right_title": "Buy / Partner",
  "right_bullets": ["Faster time-to-value", "Ongoing licensing"],
  "animations": [
    {"shape": "title", "effect": "fade", "trigger": "on_load",  "duration_ms": 400},
    {"shape": "body",  "effect": "wipe", "trigger": "on_click", "dir": "from_left",  "duration_ms": 400},
    {"shape": "right", "effect": "wipe", "trigger": "on_click", "dir": "from_right", "duration_ms": 400}
  ]
}
```

**Quote slide — dissolves in, auto-advances after 5 seconds:**
```json
{
  "type": "QUOTE",
  "title": "93% of enterprises now\nhave a multi-cloud\nstrategy",
  "animations": [
    {"shape": "title", "effect": "dissolve", "trigger": "on_load", "duration_ms": 900}
  ],
  "transition": {"type": "dissolve", "speed": "slow", "advance_ms": 5000}
}
```

**Suppress all animations for a reference table:**
```json
{
  "type": "TABLE",
  "title": "Compliance Status",
  "table": {"headers": ["Framework","Status"], "rows": [["SOC2","Compliant"]]},
  "animations": []
}
```

---

## Slide Transitions

Every slide supports an optional `"transition"` field that controls the animated transition **into** that slide. If omitted, the injector applies a smart default based on the slide's role (see defaults table below). You should always include explicit transitions so the deck feels intentional — do not leave them to chance.

### Transition Object Fields

| Field | Required | Default | Notes |
|-------|----------|---------|-------|
| `type` | Y | *(smart default)* | Transition name from the catalog below. |
| `speed` | N | role-based | `"slow"` \| `"med"` \| `"fast"` |
| `duration_ms` | N | speed-derived | Animation duration in milliseconds (100–10000). Overrides `speed` for fine control. |
| `dir` | N | — | Direction qualifier. Valid values depend on the transition type. |
| `advance_ms` | N | click-only | Auto-advance to next slide after N milliseconds. Omit for click-only advance. |

```json
{
  "type": "CONTENT",
  "title": "Our Cloud Strategy",
  "bullets": ["..."],
  "transition": {
    "type": "wipe",
    "speed": "fast",
    "dir": "l"
  }
}
```

---

### Transition Type Catalog

#### Core Transitions (universally supported)

| Name | Dir values | Best For | Effect |
|------|-----------|----------|--------|
| `fade` | — | COVER, CLOSING, QUOTE | Smooth cross-fade |
| `cut` | — | CONTENT, TABLE, CHART | Instant switch (no animation) |
| `push` | `l r u d` | CHAPTER, SECTION_BLUE, TIMELINE | New slide pushes old off screen |
| `cover` | `l r u d lu ru ld rd` | CHAPTER, SECTION_BLUE | New slide slides over old |
| `pull` | `l r u d lu ru ld rd` | CLOSING, QUOTE | Old slide pulls away |
| `wipe` | `l r u d` | CONTENT, CONTENT_CHART | Edge wipe reveal |
| `dissolve` | — | COVER, CLOSING, QUOTE | Pixel dissolve |
| `split` | `horz vert` | TWO_COLUMN, CONTENT_CHART | Split open from center |
| `zoom` | `in out` | KPI, QUOTE, COVER_FULL | Zoom in or out |
| `wheel` | — | KPI, TIMELINE | Rotating wheel wipe (4 spokes) |
| `blinds` | `horz vert` | TABLE, CONTENT_TABLE | Venetian-blinds |
| `checker` | `horz vert` | KPI | Checkerboard reveal |
| `circle` | — | QUOTE, KPI | Circular iris |
| `diamond` | — | QUOTE | Diamond iris |
| `plus` | — | SECTION_BLUE, SECTION_GREY | Plus/cross wipe |
| `wedge` | — | TIMELINE, CHAPTER | Clock-wipe |
| `comb` | `horz vert` | TABLE, CONTENT_TABLE | Interlocking comb |
| `randomBar` | `horz vert` | CONTENT, CONTENT_SIDEBAR | Random bars |
| `strips` | `lu ru ld rd` | CHAPTER, SECTION_BLUE | Diagonal strips |
| `newsflash` | — | QUOTE, KPI | Spin-zoom newsflash |
| `random` | — | *(avoid)* | Random — unprofessional in corporate decks |

#### PowerPoint 2010+ Transitions (automatically wrapped for back-compat)

| Name | Dir values | Best For | Effect |
|------|-----------|----------|--------|
| `conveyor` | `l r` | TIMELINE, CHAPTER | Conveyor belt scroll |
| `doors` | `horz vert` | COVER, CHAPTER, SECTION_BLUE | Double doors open |
| `ferris` | `l r` | KPI, TWO_COLUMN | Ferris wheel rotation |
| `flip` | `l r` | TWO_COLUMN, COVER_FULL | 3-D card flip |
| `flythrough` | `in out` | COVER, COVER_FULL, CHAPTER | Camera flies through |
| `gallery` | `l r` | CONTENT, CHART | Gallery scroll |
| `glitter` | `l r u d` | COVER, CLOSING, QUOTE | Glitter particle sweep |
| `honeycomb` | — | KPI, SECTION_BLUE | Hexagonal reveal |
| `pan` | `l r u d` | TIMELINE, CONTENT | Camera pan |
| `prism` | `l r u d` | CHAPTER, SECTION_GREY | Prism refraction |
| `reveal` | `l r` | TWO_COLUMN, CONTENT_CHART | Page peel reveal |
| `ripple` | — | QUOTE, CLOSING | Water ripple from center |
| `shred` | `in out` | CLOSING, QUOTE | Slide shreds into pieces |
| `switch` | `l r` | TWO_COLUMN, COVER_ALT | 3-D flip switch |
| `vortex` | `l r u d` | CLOSING, CHAPTER | Vortex swirl |
| `warp` | `in out` | COVER_FULL, QUOTE | Perspective warp zoom |
| `window` | `horz vert` | CONTENT, TABLE, CHART | Window pane reveal |
| `flash` | — | QUOTE, KPI | Flash to white |
| `wheelReverse` | — | KPI, TIMELINE | Reverse wheel wipe |

---

### Smart Defaults (applied automatically when `transition` is omitted)

| Slide Type | Default Transition | Speed | Dir |
|---|---|---|---|
| COVER | `fade` | slow | — |
| COVER_ALT | `fade` | slow | — |
| COVER_FULL | `flythrough` | slow | `in` |
| CHAPTER | `push` | med | `l` |
| SECTION_BLUE | `doors` | med | `horz` |
| SECTION_GREY | `prism` | med | `l` |
| CONTENT | `wipe` | fast | `l` |
| CONTENT_SIDEBAR | `wipe` | fast | `l` |
| TWO_COLUMN | `split` | fast | `horz` |
| QUOTE | `dissolve` | slow | — |
| CLOSING | `fade` | slow | — |
| CHART | `wipe` | fast | `l` |
| TABLE | `wipe` | fast | `l` |
| CONTENT_CHART | `reveal` | fast | `l` |
| CONTENT_TABLE | `reveal` | fast | `l` |
| KPI | `zoom` | med | `in` |
| TIMELINE | `conveyor` | med | `l` |

---

### Transition Design Principles

These are rules, not suggestions. Follow them exactly unless the user gives a specific override reason.

#### Rule 1: Pacing communicates hierarchy

The audience unconsciously reads transition speed as importance. A slow dissolve before the cover feels cinematic. The same slow dissolve before every content slide feels tedious. Calibrate speed by slide role:

| Slide Role | Speed | Duration | Why |
|---|---|---|---|
| Cover, closing, pause (QUOTE) | `slow` | ~1500 ms | Structural moments deserve breathing room |
| Chapter and section dividers | `med` | ~800 ms | Marks a new section — noticeable but not interruptive |
| All content slides | `fast` | ~400 ms | Keeps flow — audience is reading, not watching transitions |

#### Rule 2: One transition family per deck, one break allowed

Choose a single transition for all content slides (e.g. always `wipe l`) and a single transition for section breaks (e.g. always `push l`). The cover and closing can each get their own special transition. That's it — four total: cover, section, content, closing.

```
✓ COVER: fade slow | CHAPTER: push med l | all CONTENT: wipe fast l | CLOSING: fade slow
✗ CONTENT_1: wipe | CONTENT_2: gallery | CONTENT_3: flip | CONTENT_4: doors   (random mixing)
```

#### Rule 3: Direction reinforces narrative flow

Use left-to-right motion (`l`) for all forward-progressing slides. This matches the audience's reading direction and creates a sense of momentum.

- Sequential content, timelines: `push l`, `conveyor l`, `pan l`
- Drilling into a topic (going "deeper"): `push d`, `wipe d`
- Returning to a summary: `push r` (right-to-left suggests going back)
- Symmetrical/comparison slides: `split horz` (opens from center)

**Never use `r` (right-to-left) for forward progression** — it subconsciously signals retreating.

#### Rule 4: Match boldness to the slide's weight

| Slide Moment | Appropriate Transitions |
|---|---|
| Opening cover | `fade`, `flythrough in`, `dissolve` |
| Chapter/section break | `push l`, `doors horz`, `prism l`, `wipe l` |
| Narrative content | `wipe l`, `cut` (virtually invisible) |
| High-impact statement (QUOTE, KPI) | `zoom in`, `dissolve`, `circle` |
| Comparison/two-column | `split horz` |
| Timeline/sequential steps | `conveyor l`, `pan l`, `push l` |
| Closing | `fade`, `dissolve`, `ripple` |

#### Rule 5: Avoid theatrical transitions in corporate decks

These transitions are available but inappropriate for enterprise/business presentations. Do not use them:

- `glitter`, `honeycomb`, `shred`, `newsflash` — informal/flashy
- `checker`, `strips`, `random_bars`, `comb`, `blinds` — feel dated or gimmicky
- `random` — never appropriate; always inconsistent

Exception: a single theatrical transition on the very last CLOSING slide can be a deliberate surprise. Use `ripple` or `vortex` at most once, at the end.

#### Rule 6: Auto-advance only for scripted pause slides

`advance_ms` is only appropriate for QUOTE slides or video-synchronized content. For all other slides, omit it — the presenter controls pace via click.

#### Rule 7: Speed calibration values

| Speed | Duration | Use For |
|---|---|---|
| `slow` | ~1500 ms | COVER, CLOSING, QUOTE |
| `med` | ~800 ms | CHAPTER, SECTION_BLUE, SECTION_GREY, KPI, TIMELINE |
| `fast` | ~400 ms | CONTENT, CONTENT_SIDEBAR, TWO_COLUMN, CHART, TABLE, CONTENT_CHART, CONTENT_TABLE |

### Transition Examples

**Simple fade (cover):**
```json
{
  "type": "COVER",
  "title": "Cloud Migration\nStrategy 2026",
  "subtitle": "<company-name> Digital Solutions",
  "transition": {"type": "fade", "speed": "slow"}
}
```

**Directional push (chapter break):**
```json
{
  "type": "CHAPTER",
  "title": "Migration\nApproach",
  "subtitle": "Phase overview and workstreams",
  "transition": {"type": "push", "speed": "med", "dir": "l"}
}
```

**Zoom in (KPI slide — builds anticipation):**
```json
{
  "type": "KPI",
  "title": "Program Highlights",
  "kpis": [{"value": "$4.2M", "label": "Annual Savings"}],
  "transition": {"type": "zoom", "speed": "med", "dir": "in"}
}
```

**Auto-advancing quote (pauses 5 seconds, then advances):**
```json
{
  "type": "QUOTE",
  "title": "93% of enterprises now\nhave a multi-cloud\nstrategy",
  "transition": {"type": "dissolve", "speed": "slow", "advance_ms": 5000}
}
```

**Custom duration (1.2 s flythrough on bold cover):**
```json
{
  "type": "COVER_FULL",
  "title": "The Future of Work Is Now",
  "transition": {"type": "flythrough", "duration_ms": 1200, "dir": "in"}
}
```

---

## Speaker Notes

Add speaker notes to **any slide type** with the `"speaker_notes"` field. Notes appear in PowerPoint's presenter view and are not visible to the audience.

```json
{
  "type": "CONTENT",
  "title": "Key Findings",
  "bullets": ["..."],
  "speaker_notes": "Emphasize the 40% cost reduction figure.\nMention that this exceeds industry benchmarks by 15%.\nAsk if there are questions before moving to the next section."
}
```

- Use `\n` to create separate paragraphs in the notes pane
- No character limit (notes are not displayed on the slide)
- Available on all 17 slide types including CLOSING

---

## Content Plan Structure

### Full JSON Schema

```json
{
  "slides": [
    {
      "type": "COVER | COVER_ALT | COVER_FULL | CHAPTER | SECTION_BLUE | SECTION_GREY | CONTENT | CONTENT_SIDEBAR | TWO_COLUMN | QUOTE | CLOSING | CHART | TABLE | CONTENT_CHART | CONTENT_TABLE | KPI | TIMELINE",
      "title": "string (with \\n for line breaks)",
      "subtitle": "string (COVER/COVER_ALT/COVER_FULL/CHAPTER only)",
      "bullets": ["string", {"text": "string", "sub_bullets": ["string"]}],
      "body_text": "string (alternative to bullets for CONTENT/CONTENT_SIDEBAR)",
      "left_bullets": ["string (TWO_COLUMN only)"],
      "right_bullets": ["string (TWO_COLUMN only)"],
      "chart": {"type": "string", "categories": [], "series": [{"name": "", "values": []}]},
      "left_chart": {"...chart object (TWO_COLUMN only)"},
      "right_chart": {"...chart object (TWO_COLUMN only)"},
      "table": {"headers": [], "rows": [[]]},
      "kpis": [{"value": "", "label": "", "color": "#hex"}],
      "milestones": [{"label": "", "description": "", "color": "#hex"}],
      "icons": [{"name": "", "position": "top_right|top_left|before_title|{x,y}", "size": 1.0}],
      "speaker_notes": "string (\\n for paragraph breaks)",
      "transition": {
        "type": "fade|cut|push|cover|pull|wipe|dissolve|split|zoom|wheel|blinds|checker|circle|diamond|plus|wedge|comb|randomBar|strips|newsflash|conveyor|doors|ferris|flip|flythrough|gallery|glitter|honeycomb|pan|prism|reveal|ripple|shred|switch|vortex|warp|window|flash|wheelReverse",
        "speed": "slow|med|fast",
        "duration_ms": 800,
        "dir": "l|r|u|d|horz|vert|in|out|lu|ru|ld|rd",
        "advance_ms": 0
      },
      "animations": [
        {
          "shape": "title|body|subtitle|left|right|<shape_id_int>",
          "effect": "fade|fly_in|wipe|zoom|dissolve|appear|split|box|swivel|blinds|checker|circle|diamond|plus|wedge|wheel|random_bars|strips|float_in|bounce|disappear|fade_out|fly_out|wipe_out|zoom_out|pulse|spin|grow_shrink|bold_flash",
          "trigger": "on_click|after_prev|with_prev|on_load",
          "dir": "<effect-dependent>",
          "duration_ms": 500,
          "delay_ms": 0,
          "text_build": "all_at_once|by_bullet"
        }
      ]
    }
  ]
}
```

### Field Applicability by Slide Type

| Type | title | subtitle | bullets | body_text | left/right_bullets | chart | table | kpis | milestones | icons | speaker_notes | transition | animations |
|------|-------|----------|---------|-----------|-------------------|-------|-------|------|------------|-------|--------------|------------|------------|
| COVER | Y | Y | - | - | - | - | - | - | - | Y | Y | Y | Y |
| COVER_ALT | Y | Y | - | - | - | - | - | - | - | Y | Y | Y | Y |
| COVER_FULL | Y | Y | - | - | - | - | - | - | - | Y | Y | Y | Y |
| CHAPTER | Y | Y | - | - | - | - | - | - | - | Y | Y | Y | Y |
| SECTION_BLUE | Y | - | - | - | - | - | - | - | - | Y | Y | Y | Y |
| SECTION_GREY | Y | - | - | - | - | - | - | - | - | Y | Y | Y | Y |
| CONTENT | Y | - | Y | Y | - | Y* | - | - | - | Y | Y | Y | Y |
| CONTENT_SIDEBAR | Y | - | Y | Y | - | - | - | - | - | Y | Y | Y | Y |
| TWO_COLUMN | Y | - | - | - | Y | Y** | - | - | - | Y | Y | Y | Y |
| QUOTE | Y | - | - | - | - | - | - | - | - | Y | Y | Y | Y |
| CLOSING | - | - | - | - | - | - | - | - | - | Y | Y | Y | Y |
| CHART | Y | - | - | - | - | Y | - | - | - | Y | Y | Y | Y† |
| TABLE | Y | - | - | - | - | - | Y | - | - | Y | Y | Y | Y† |
| CONTENT_CHART | Y | - | Y | - | - | Y | - | - | - | Y | Y | Y | Y |
| CONTENT_TABLE | Y | - | Y | - | - | - | Y | - | - | Y | Y | Y | Y |
| KPI | Y | - | - | - | - | - | - | Y | - | Y | Y | Y | Y† |
| TIMELINE | Y | - | - | - | - | - | - | - | Y | Y | Y | Y | Y† |

† Only `"title"` shape targeting is reliable for these slide types. Chart, table, KPI card, and timeline node shapes are injected as free shapes and cannot be targeted by role name.

\* CONTENT: `chart` only used when `bullets` is absent (replaces body with full-width chart)
\*\* TWO_COLUMN: via `left_chart` / `right_chart` fields (replaces corresponding bullet column)

---

## Presentation Design Guidelines

### Slide Flow Best Practices

A well-structured presentation follows this pattern:

```
COVER (or COVER_ALT or COVER_FULL)
  CHAPTER (optional, for long decks)
    SECTION_BLUE or SECTION_GREY (optional, for visual breaks)
      CONTENT / CONTENT_SIDEBAR / TWO_COLUMN (the substance)
      CHART / TABLE / CONTENT_CHART / CONTENT_TABLE (data visualization)
      KPI (metrics dashboard)
      TIMELINE (project phases)
      QUOTE (for impact moments)
CLOSING
```

**Short deck (5-8 slides):** COVER -> 3-5 CONTENT slides -> QUOTE (optional) -> CLOSING

**Medium deck (8-15 slides):** COVER -> SECTION -> CONTENT blocks + CHART/TABLE -> SECTION -> CONTENT blocks -> KPI -> QUOTE -> CLOSING

**Long deck (15+ slides):** COVER -> CHAPTER -> SECTION -> CONTENT/CHART/TABLE blocks -> CHAPTER -> SECTION -> CONTENT blocks -> KPI -> TIMELINE -> QUOTE -> CLOSING

### Content Quality Rules

1. **Never exceed character limits.** The injector will warn, but the slide will look cramped.
2. **Prefer fewer, punchier bullets.** 4 clear bullets beat 6 cramped ones.
3. **Use sub-bullets sparingly.** They eat vertical space fast -- 1-2 sub-bullets per parent max.
4. **Keep titles concise.** They should scan in 2 seconds.
5. **Use QUOTE slides for emphasis.** Drop one after a heavy content section to let the audience breathe.
6. **Alternate slide types for visual variety.** Don't use 5 CONTENT slides in a row. Mix in charts, KPIs, or section dividers.
7. **Match bullet counts in TWO_COLUMN.** Uneven columns look unbalanced.
8. **Use charts for data, not text.** If you have numbers, visualize them.
9. **Limit table complexity.** 4-5 columns, 5-8 rows max for readability.
10. **Use KPI slides for executive summaries.** 3-4 key metrics make an immediate impact.

### Content Limit Quick Reference

| Slide Type | Title Limit | Body Limit | Best For |
|---|---|---|---|
| COVER | 20 chars x 2 lines | 35 chars x 2 lines | Opening |
| COVER_ALT | 20 chars x 3 lines | 45 chars x 2 lines | Alt opening |
| COVER_FULL | 30 chars x 2 lines | 45 chars x 2 lines | Bold statement |
| CHAPTER | 18 chars x 2 lines | 35 chars x 1 line | Chapter break |
| SECTION_BLUE | 17 chars x 2 lines | (icon only) | Section divider |
| SECTION_GREY | 17 chars x 2 lines | (icon only) | Section divider alt |
| CONTENT | 30 chars x 1 line | 55 chars x 5-6 bullets | Main content |
| CONTENT_SIDEBAR | 15 chars sidebar | 55 chars x 7-8 bullets | Text-heavy |
| TWO_COLUMN | 30 chars x 1 line | 25-28 chars x 4-5 bullets/col | Comparisons |
| QUOTE | 30 chars x 3 lines | (none) | Key statements |
| CLOSING | (none) | (none) | End slide |
| CHART | 30 chars x 1 line | Full-width chart | Data visualization |
| TABLE | 30 chars x 1 line | Full-width table (6 cols, 10 rows) | Structured data |
| CONTENT_CHART | 30 chars x 1 line | 40 chars x 6 bullets + chart | Text + visual |
| CONTENT_TABLE | 30 chars x 1 line | 40 chars x 6 bullets + table | Text + data |
| KPI | 30 chars x 1 line | 2-4 metric cards | Executive metrics |
| TIMELINE | 30 chars x 1 line | 2-6 milestones | Project phases |

---

## Complete Example

Here is a full 15-slide content plan demonstrating all features:

```json
{
  "slides": [
    {
      "type": "COVER",
      "title": "Cloud Migration\nStrategy 2026",
      "subtitle": "<company-name> Digital Solutions",
      "speaker_notes": "Welcome everyone. Today we'll walk through our cloud migration strategy."
    },
    {
      "type": "CHAPTER",
      "title": "Executive\nSummary",
      "subtitle": "Strategic overview and goals"
    },
    {
      "type": "KPI",
      "title": "Program Highlights",
      "kpis": [
        {"value": "$4.2M", "label": "Projected Savings", "color": "#0080BA"},
        {"value": "99.99%", "label": "Target Uptime", "color": "#4CAF50"},
        {"value": "47%", "label": "Faster Deploys", "color": "#FF9901"},
        {"value": "Zero", "label": "Critical Incidents", "color": "#00447A"}
      ],
      "speaker_notes": "These are our four north-star metrics for the program."
    },
    {
      "type": "CONTENT",
      "title": "Why Migrate to the Cloud?",
      "bullets": [
        "Reduce infrastructure costs by 40%",
        "Improve scalability and agility",
        "Enhance security posture",
        {
          "text": "Enable digital transformation",
          "sub_bullets": [
            "AI/ML workloads",
            "Real-time analytics"
          ]
        },
        "Accelerate time to market"
      ],
      "icons": [{"name": "cloud", "position": "top_right"}]
    },
    {
      "type": "TWO_COLUMN",
      "title": "Current vs. Target State",
      "left_bullets": [
        "On-premise data centers",
        "Manual provisioning",
        "3-month release cycles",
        "Siloed teams"
      ],
      "right_bullets": [
        "Multi-cloud architecture",
        "Infrastructure as Code",
        "Weekly deployments",
        "Cross-functional DevOps"
      ]
    },
    {
      "type": "SECTION_BLUE",
      "title": "Migration\nApproach"
    },
    {
      "type": "TIMELINE",
      "title": "Migration Roadmap",
      "milestones": [
        {"label": "Discovery", "description": "Assess current state", "color": "#0080BA"},
        {"label": "Planning", "description": "Design target arch", "color": "#FF9901"},
        {"label": "Migration", "description": "Move workloads", "color": "#00447A"},
        {"label": "Optimize", "description": "Tune & scale", "color": "#4CAF50"}
      ]
    },
    {
      "type": "CONTENT_CHART",
      "title": "Cloud Adoption by Department",
      "bullets": [
        "Engineering leads at 85%",
        "Marketing migrated Q2",
        "Finance scheduled Q3",
        "HR pilot in progress"
      ],
      "chart": {
        "type": "bar",
        "categories": ["Engineering", "Marketing", "Finance", "HR", "Legal"],
        "series": [{"name": "% Migrated", "values": [85, 72, 30, 15, 10]}],
        "show_data_labels": true
      }
    },
    {
      "type": "CHART",
      "title": "Infrastructure Cost Trend",
      "chart": {
        "type": "line",
        "categories": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        "series": [
          {"name": "Actual", "values": [180, 165, 150, 132, 120, 108]},
          {"name": "Forecast", "values": [180, 170, 155, 140, 125, 110]}
        ],
        "title": "Monthly Spend ($K)",
        "show_gridlines": true
      }
    },
    {
      "type": "SECTION_GREY",
      "title": "Security &\nCompliance"
    },
    {
      "type": "TABLE",
      "title": "Compliance Status",
      "table": {
        "headers": ["Framework", "Status", "Gap Count", "ETA"],
        "rows": [
          ["SOC 2 Type II", "Compliant", "0", "Current"],
          ["ISO 27001", "In Progress", "3", "Q3 2026"],
          ["PCI DSS", "Planning", "7", "Q4 2026"],
          ["HIPAA", "Assessment", "5", "Q1 2027"]
        ]
      },
      "icons": [{"name": "cybersecurity", "position": "top_right", "size": 0.8}]
    },
    {
      "type": "CONTENT_TABLE",
      "title": "Team & Budget Allocation",
      "bullets": [
        "Total program budget: $6.8M",
        "36 FTEs across all streams",
        "External consulting: 20%",
        "Contingency reserve: 10%"
      ],
      "table": {
        "headers": ["Stream", "FTEs", "Budget"],
        "rows": [
          ["Migration", "12", "$2.4M"],
          ["Security", "8", "$1.6M"],
          ["Platform", "10", "$1.8M"],
          ["PMO", "6", "$1.0M"]
        ]
      }
    },
    {
      "type": "CONTENT",
      "title": "Cost Breakdown",
      "chart": {
        "type": "doughnut",
        "categories": ["Infrastructure", "Personnel", "Licensing", "Training", "Contingency"],
        "series": [{"name": "Budget", "values": [35, 30, 18, 10, 7]}]
      },
      "speaker_notes": "Note that personnel costs include both internal FTEs and external consultants."
    },
    {
      "type": "QUOTE",
      "title": "93% of enterprises now\nhave a multi-cloud\nstrategy"
    },
    {
      "type": "CLOSING",
      "speaker_notes": "Thank you for your time. Happy to take questions."
    }
  ]
}
```

---

## Execution Reference

### Generate a Presentation

```bash
# From the skill bundle directory
python scripts/<company-name>_injector.py content_plan.json \
  --template assets/<company-name>_2021_template.pptx \
  --output output/presentation.pptx
```

### Visual Validation (Optional but Recommended)

```bash
# Convert to PDF
soffice --headless --convert-to pdf output/presentation.pptx --outdir output/

# Convert PDF to images
python -c "
from pdf2image import convert_from_path
images = convert_from_path('output/presentation.pdf', dpi=150)
for i, img in enumerate(images, 1):
    img.save(f'output/slide_{i}.png')
print(f'Saved {len(images)} slide images')
"
```

### Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| `ModuleNotFoundError: pptx` | python-pptx not installed | `pip install python-pptx` |
| Validation warnings | Content exceeds character limits | Shorten text to fit within limits |
| Empty CLOSING slide | Template file corrupted/wrong | Use the original `<company-name>_2021_template.pptx` from assets/ |
| Placeholder not found | Wrong slide type or template mismatch | Verify template has all 27 layouts |
| Unknown chart type | Invalid chart type string | Use one of: column, bar, line, pie, doughnut, stacked_column, stacked_bar, area |
| Table column mismatch | Row has different cell count than headers | Ensure every row array has the same length as headers |
| Unknown icon name | Icon name not in catalog | Check the icon catalog section for valid names |

---

## Template Details

- **Template**: <company-name> 2021 corporate template
- **Dimensions**: 13.33" x 7.5" (16:9 widescreen)
- **Font**: Helvetica Neue (inherited from slide master)
- **Title font size**: 44pt
- **Body font sizes**: 28pt (level 1), 24pt (level 2), 20pt (level 3)
- **Brand colors**: Blue (#0080BA), Orange (#FF9901), Dark Blue (#00447A), Grey (#989998), Green (#4CAF50), Light Blue (#5BC0DE)
- **Total layouts available**: 27 (17 slide types curated for use)
- **Embedded icons**: 95 branded SVG+PNG icon pairs
