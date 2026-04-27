# JZ Dynamic Content Sections - Examples & Usage Guide

This document contains comprehensive examples showing all features of the JZ Dynamic Content Sections extension.

> **Note:** After adding a section through the properties panel, you must toggle any checkbox (like 'Hide section if no data') to refresh the visualization. This is a Qlik Sense limitation.

---

## Table of Contents
1. [KPI Examples](#kpi-examples)
2. [List Examples](#list-examples)
3. [Table Examples](#table-examples)
4. [Grid Examples](#grid-examples)
5. [Concatenate Examples](#concatenate-examples)
6. [Color Syntax Examples](#color-syntax-examples)
7. [Box/Card Examples](#boxcard-examples)
8. [Visualization Examples](#visualization-examples)
9. [Colored List Examples](#colored-list-examples)
10. [Complete Feature Example](#complete-feature-example)
11. [Claude AI Examples](#claude-ai-examples)
12. [Qlik200 Strategy Review Template](#qlik200-strategy-review-template)
13. [Strategic Players Template](#strategic-players-template)

---

## KPI Examples

Simple KPI display with labels:

```markdown
# Sales Dashboard

#[kpi label="Total Revenue"]{{measure1}}#[/kpi]

#[kpi label="Total Orders"]{{measure2}}#[/kpi]

**Performance:** {green:Above Target}
```

**Recommended Settings:**
- Section Style: `card`
- Section Width: `full`

---

## List Examples

### Multi-Column Lists with Different Styles

```markdown
# Product Categories

## 2-Column Bulleted List
#[list type="bulleted" columns="2"]{{dim1}}#[/list]

## 3-Column Plain List
#[list type="plain" columns="3"]{{dim1}}#[/list]

## 4-Column Numbered List (Horizontal)
#[list type="numbered" columns="4" numbering="horizontal"]{{dim1}}#[/list]

## 3-Column Numbered List (Vertical)
#[list type="numbered" columns="3" numbering="vertical"]{{dim1}}#[/list]

## 3-Column List with Dividers
#[list type="bulleted" columns="3" dividers="true"]{{dim1}}#[/list]
```

**List Attributes:**
- `type`: `bulleted`, `plain`, or `numbered`
- `columns`: Number of columns (1-6)
- `numbering`: `horizontal` or `vertical` (for numbered lists)
- `dividers`: `true` or `false` (adds vertical dividers between columns)

**Recommended Settings:**
- Section Style: `card`
- Section Width: `full`

---

## Table Examples

### Tables with Colored Headers and Striped Rows

```markdown
# Sales Report

## Green Header
#[table headerColor="#009845"]
Product|{{dim1}}
Category|{{dim2}}
Revenue|{{measure1}}
Units|{{measure2}}
#[/table]

## Blue Header with Stripes
#[table headerColor="#1976D2" stripedRows="true"]
Product|{{dim1}}
Revenue|{{measure1}}
#[/table]
```

### Tables with Custom Column Widths

```markdown
# Contract Details

## Manual Width Control (20%, 50%, 30%)
#[table headerColor="#009845" widths="20,50,30"]
Term|{{dim1}}
Description|{{dim2}}
Value|{{measure1}}
#[/table]

## Smart Auto-Width (analyzes content automatically)
#[table headerColor="#1976D2"]
ID|{{dim1}}
Long Description Text|{{dim2}}
Amount|{{measure1}}
#[/table]
```

**Table Attributes:**
- `headerColor`: Any hex color (e.g., `#009845`, `#1976D2`)
- `stripedRows`: `true` or `false` (alternating row colors)
- `widths`: Comma-separated percentages (e.g., `"20,50,30"` or `"15%,60%,25%"`)
  - **Manual widths**: Specify exact percentages for each column
  - **Smart auto-width**: Leave empty to automatically calculate based on content length
  - Smart calculation analyzes actual data to optimize column widths

**Recommended Settings:**
- Section Style: `card`
- Section Width: `full`

---

## Grid Examples

### Simple Grids with Optional Dividers

```markdown
# Product Grid

## Standard Grid
#[grid columns="3"]{{dim1}}#[/grid]

## Grid with Dividers
#[grid columns="3" dividers="true"]{{dim1}}#[/grid]
```

**Grid Attributes:**
- `columns`: Number of columns
- `dividers`: `true` or `false` (adds borders between items)

**Recommended Settings:**
- Section Style: `card`
- Section Width: `full`

---

## Concatenate Examples

Join dimension values with custom delimiters:

```markdown
# Available Products

#[concat delimiter=" • "]{{dim1}}#[/concat]

**Total Items:** {{measure1}}
```

**Concat Attributes:**
- `delimiter`: Any text string (e.g., `" • "`, `" | "`, `", "`)

**Recommended Settings:**
- Section Style: `card`
- Section Width: `full`

---

## Color Syntax Examples

### Inline Text Coloring

```markdown
# Status Dashboard

**Status:** {green:Active}

**Priority:** {red:High}

**Category:** {blue:Premium}

**Amount:** {#009845:${{measure1}}}
```

**Color Syntax:**
- Named colors: `{green:text}`, `{red:text}`, `{blue:text}`, `{orange:text}`, `{purple:text}`, `{yellow:text}`, `{gray:text}`
- Hex colors: `{#009845:text}`

**Recommended Settings:**
- Section Style: `highlighted`
- Section Width: `half` or `full`

---

## Box/Card Examples

### Colored Information Boxes

```markdown
# Colored Boxes

#[box color="#009845" bgColor="#f0fff4"]
**Success Message**

Everything is working perfectly!
#[/box]

#[box color="#d32f2f" bgColor="#fff0f0"]
**Alert**

Action required on this item.
#[/box]

#[box color="#1976d2" bgColor="#e3f2fd"]
**Information**

Here's some helpful information for you.
#[/box]

#[box color="#f57c00" bgColor="#fff8e1"]
**Warning**

Please review this carefully.
#[/box]
```

**Box Attributes:**
- `color`: Text color (hex)
- `bgColor`: Background color (hex)

**Recommended Settings:**
- Section Style: `plain`
- Section Width: `full`

---

## Image Examples

### Using Images from Media Library

```markdown
# Images from Media Library (Recommended)

## Simple filename (automatic media library lookup)
#[image src="logo.png" width="200px" align="center"]#[/image]

## Explicit media library reference
#[image src="media:Qlik_200_Initiative.png" width="150px" height="25px" align="inline-right"]#[/image]

## Floating image with text wrap
#[image src="product-image.jpg" width="150px" align="float-left"]#[/image]
Your text content flows around the image...
```

**Image Attributes:**
- `src`: Filename from media library (e.g., `logo.png` or `media:logo.png`)
- `width`, `height`: Size in px or other CSS units
- `maxWidth`, `maxHeight`: Maximum dimensions
- `align`: `left`, `center`, `right`, `float-left`, `float-right`, `inline-left`, `inline-right`
- `alt`: Alternative text for accessibility

**Recommended Settings:**
- Upload images to the Qlik media library (Assets > Media library)
- Use simple filenames without paths

---

## Visualization Examples

### Embedding Qlik Visualizations

```markdown
# Embedded Visualizations

#[box color="#1976d2" bgColor="#e3f2fd"]
**Two ways to embed:**
- By ID: `#[viz id="abc123"]#[/viz]`
- By Name: `#[viz name="Trend"]#[/viz]`
#[/box]

## Method 1: By Name (Easiest)
#[viz name="Trend" height="400px"]#[/viz]

## Method 2: By ID
#[viz id="your-masteritem-id" height="300px"]#[/viz]

**Tip:** Using `name` is easier - just use the exact master item name!
```

**Viz Attributes:**
- `name`: Master item name (easier)
- `id`: Master item ID
- `height`: Height in pixels (e.g., `400px`)

**Recommended Settings:**
- Section Style: `card`
- Section Width: `full`

---

## Colored List Examples

### Lists with Automatic Dimension Coloring

```markdown
#[header bgColor="#87205d"]Strategic Players|Colored by Role#[/header]

#[list type="plain" colorBy="[Strategic Player Role]"]
{{[Strategic Player]}}
#[/list]

**Legend:** Colors come from the master item dimension settings. If your dimension has colors defined (green, red, etc.), they'll automatically appear in the list!

#[box color="#1976d2" bgColor="#e3f2fd"]
**How colorBy works:**

The colorBy attribute maps each list item to a dimension. If that dimension has master item colors, those colors are applied automatically!

**Example with your own data:**

Replace "Strategic Player" and "Strategic Player Role" with your master item names.
#[/box]
```

**ColorBy Feature:**
- The `colorBy` attribute maps list items to a dimension
- If the dimension has master item colors defined in Qlik, they're automatically applied
- Format: `colorBy="[Dimension Name]"`

**Recommended Settings:**
- Section Style: `card`
- Section Width: `full`
- Group Width: `half`

---

## Complete Feature Example

All features in one comprehensive example:

```markdown
# {blue:Sales Overview}

#[box color="#1976d2" bgColor="#e3f2fd"]
**Welcome!** This example shows all available features.
#[/box]

## Key Metrics
#[kpi label="Total Revenue"]{{measure1}}#[/kpi]

## Top Products (Multi-Column)
#[list type="numbered" columns="2"]{{dim1}}#[/list]

## Categories
#[concat delimiter=" | "]{{dim2}}#[/concat]

## Detailed Report (Colored Header)
#[table headerColor="#009845" stripedRows="true"]
Product|{{dim1}}
Category|{{dim2}}
Revenue|{{measure1}}
#[/table]

## Product Grid
#[grid columns="3"]{{dim1}}#[/grid]

#[box color="#009845" bgColor="#f0fff4"]
**Status:** {green:Active} | **Priority:** {red:High}
#[/box]
```

**Recommended Settings:**
- Section Style: `card`
- Section Width: `full`

---

## Claude AI Examples

### Complex Nested AI Analysis

> **Note:** Enable Claude AI in extension properties and add your API token to use these features.

```markdown
# 📊 AI-Powered Sales Dashboard

#[box color="#1976D2" bgColor="#e3f2fd"]
## 🎯 Top Performers
#[list type="numbered"]{{dim1}}#[/list]

**AI Analysis:**
#[claude prompt="Analyze the top 3 performers and provide key insights" data="{{dim1}},{{measure1}}"]
#[/claude]
#[/box]

#[box color="#4CAF50" bgColor="#e8f5e9"]
## 📈 Performance Grid
#[grid columns="2"]{{dim1}}#[/grid]

**AI Recommendations:**
#[claude prompt="Provide 3 actionable recommendations based on this data" data="{{dim1}},{{measure1}}"]
#[/claude]
#[/box]

#[box color="#FF9800" bgColor="#fff3e0"]
## 🔍 Quick Stats
- **Items:** #[concat delimiter=", "]{{dim1}}#[/concat]
- **Summary:** #[claude prompt="Summarize this data in one sentence" data="{{dim1}},{{measure1}}"]#[/claude]
#[/box]
```

**Claude Attributes:**
- `prompt`: The question or instruction for Claude AI
- `data`: The data to analyze (comma-separated placeholders)

**Recommended Settings:**
- Section Style: `card`
- Section Width: `full`

---

## Qlik200 Strategy Review Template

A complete 3-column strategy review layout:

### Group 1: Situational Appraisal (Width: third)

```markdown
#[header]Situational Appraisal|Where we are#[/header]

**ARR Trend:** — Flat  |  **Relationship Score:** -

### Charter Statement

Our charter statement content goes here...

### Strategic Players

#[list type="plain"]
{green:Brad Davis (Sr. Manager, Data)}

{green:Elena Garvey (IT Director - Data Analytics & Business Insights)}

{green:Joe Medsker (Software Engineer)}

{red:Samir Daiya (Chief Information Officer)}

Gregory Kirk (Managing Director of Architecture)
#[/list]

### Key Information

#[table headerColor="#006580"]
Qlik Solutions|Qlik Catalog, Qlik Sense Enterprise SaaS
Contracts|Term 12: $37,000 (SaaS), Term 13: $252,466 (SaaS)
Competitive Landscape|Tableau, Power BI
Channel Partners|Amazon Web Services
#[/table]
```

### Group 2: Account Strategy (Width: third)

```markdown
#[header]Account Strategy|Where we want to go#[/header]

### Our Strategic Strengths

**Drive competitive advantage** within MSI's go-to-market offerings: Qlik's unique OEM capabilities can bring advanced, self-service analytics and AI capabilities to Motorola's products in the market

**Help accelerate value** delivered from acquisitions: Qlik can help MSI drive value faster from acquired companies via our data integration expertise

**Supply Chain Optimization:** Qlik's ability to analyze massive amounts of data in real-time can help Motorola's supply chain team optimize decisions

### Our Critical Vulnerabilities

⚠️ Motorola has no Qlik Analytics skills across the organization

⚠️ Motorola lacks skills, knowledge, expertise in using Qlik Analytics, vs. Tableau
```

### Group 3: Action Plan (Width: third)

```markdown
#[header]Action Plan|How we get there#[/header]

### Next Steps

#[list type="numbered" columns="2"]
Schedule executive briefing with CIO

Conduct OEM capabilities workshop

Propose POC for supply chain analytics

Engage with AWS partnership team

Develop training program proposal

Schedule quarterly business reviews
#[/list]

### Timeline & Milestones

#[table headerColor="#006580" stripedRows="true"]
Quarter|Milestone|Owner|Status
Q1 2026|Executive Alignment|{green:Dean Bruckman}|{green:Complete}
Q2 2026|POC Launch|Jeff Jordan|{blue:In Progress}
Q3 2026|Production Deployment|TBD|Planned
Q4 2026|Expansion Discussion|TBD|Planned
#[/table]

### Key Actions

✓ Complete technical discovery

✓ Submit OEM proposal

✓ Align on success metrics

⏳ Finalize training roadmap
```

**Recommended Settings:**
- Section Style: `plain`
- Section Width: `full` (for each section within the groups)
- Group Width: `third` (for all 3 groups)

---

## Strategic Players Template

Multiple sections in a single full-width group:

### Section 1: Strategic Players (Width: full)

```markdown
#[title separator="true" style="bold" size="18px"]{#87205d:Strategic Players}#[/title]

{green:Brad Davis (Sr. Manager, Data)}

{green:Elena Garvey (IT Director - Data Analytics & Business Insights)}

{green:Joe Medsker (Software Engineer)}

{red:Samir Daiya (Chief Information Officer)}

Gregory Kirk (Managing Director of Architecture)
```

### Section 2: Qlik Solutions (Width: half)

```markdown
**{#87205d:Qlik Solutions}**

Qlik Catalog, Qlik Sense Enterprise SaaS
```

### Section 3: Contracts (Width: half)

```markdown
**{#87205d:Contracts}**

Term 12: 37,000.00 (SaaS)

Term 13: 252,466.07 (SaaS, Subscription)

Term 14: 279,194.02 (Subscription)

Term 36: 100,000.00 (Subscription)
```

### Section 4: Competitive Landscape (Width: half)

```markdown
**{#87205d:Competitive Landscape}**

Tableau, Power BI, Looker
```

### Section 5: Channel Partners (Width: half)

```markdown
**{#87205d:Channel Partners}**

Amazon Web Services, Microsoft Azure
```

**Recommended Settings:**
- Section Style: `card`
- Section Width: Varies (see above)
- Group Width: `full`

---

## Reference Guide

### Data References

| Syntax | Description |
|--------|-------------|
| `{{dim1}}`, `{{dim2}}` | Direct dimension reference |
| `{{measure1}}`, `{{measure2}}` | Direct measure reference |
| `{{[Master Item Name]}}` | Master item by name (recommended) |

### Content Tags

| Tag | Purpose | Key Attributes |
|-----|---------|----------------|
| `#[kpi]` | Display key metrics | `label` |
| `#[list]` | Create lists | `type`, `columns`, `numbering`, `dividers`, `colorBy` |
| `#[table]` | Create tables | `headerColor`, `stripedRows` |
| `#[grid]` | Create grids | `columns`, `dividers` |
| `#[concat]` | Join values | `delimiter` |
| `#[box]` | Colored boxes | `color`, `bgColor` |
| `#[image]` | Display images | `src`, `width`, `height`, `align`, `maxWidth`, `maxHeight`, `alt` |
| `#[viz]` | Embed visualizations | `name` or `id`, `height` |
| `#[header]` | Section headers | `bgColor` |
| `#[title]` | Custom titles | `separator`, `style`, `size` |
| `#[claude]` | AI analysis | `prompt`, `data` |

### Inline Styling

| Syntax | Purpose |
|--------|---------|
| `{green:text}` | Green text |
| `{red:text}` | Red text |
| `{blue:text}` | Blue text |
| `{#HEXCODE:text}` | Custom hex color |

---

## Tips & Best Practices

1. **Master Items**: Always use master item names in double square brackets: `{{[Sales Amount]}}`
2. **Refresh Limitation**: After adding examples, toggle any checkbox in properties to refresh
3. **Color Consistency**: Use the Color Maps feature for consistent theming across all sections
4. **Responsive Design**: Use width settings (`full`, `half`, `third`) to create responsive layouts
5. **Performance**: Use `hideIfNoData` to improve load times when sections have no data
6. **Multi-User**: Enable Multi-User Collaboration for shared editing across teams
7. **AI Features**: Configure Claude AI integration for automated insights and analysis
8. **Images**: Upload images to Qlik media library and reference by filename (e.g., `#[image src="logo.png"]#[/image]`)

---

## Additional Resources

- **Extension Properties**: Configure all settings in the Qlik properties panel
- **Color Maps**: Define reusable color mappings in the Color Maps accordion
- **MCP Integration**: Enable multi-user collaboration and SharePoint storage
- **Version**: Check the About section in properties for current version information

