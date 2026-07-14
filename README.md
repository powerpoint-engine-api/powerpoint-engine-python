# PowerPoint Engine — Python SDK

Python client for the [PowerPoint Engine API](https://powerpointengine.io): generate `.pptx` from Markdown or a structured template, edit / merge / replace inside existing decks, and convert PPTX to PDF — over plain HTTP, no PowerPoint or LibreOffice install needed.

## Install

```bash
pip install powerpoint-engine-api
```

## Quick start

```python
from powerpoint_engine import PowerPointEngine

engine = PowerPointEngine()  # anonymous: works, but files are watermarked

result = engine.generate(markup="""
# Q3 Business Review

## Highlights
- Revenue up 24% QoQ
- 3 new enterprise customers

## Quarterly Numbers
| Region | Q1  | Q2  | Q3  |
|--------|-----|-----|-----|
| EU     | 1.0 | 1.2 | 1.4 |
| US     | 2.0 | 2.4 | 2.6 |
""", theme="corporate")

engine.download(result, "review.pptx")
```

Markdown tables become native PowerPoint tables. A fenced ```` ```chart ```` block under a `## heading` becomes a real editable chart (col / bar / line / pie / combo with a secondary axis):

```text
## Revenue vs Margin
```chart
type: combo
categories: Q1, Q2, Q3
Revenue: 120, 150, 170
Margin % (line, secondary): 10, 12, 14
```
```

## Work with existing decks

```python
# Replace {{placeholders}} without breaking formatting
engine.replace("template.pptx", {"client": "ACME", "year": "2026"})

# Duplicate / delete / reorder slides (1-based)
engine.edit("deck.pptx", [
    {"op": "duplicate", "slide": 2},
    {"op": "delete", "slide": 5},
    {"op": "move", "slide": 3, "to": 1},
])

# Merge 2-5 decks; each keeps its own design
engine.merge(["intro.pptx", "results.pptx", "outro.pptx"])

# Convert to PDF
pdf = engine.to_pdf("deck.pptx")
engine.download(pdf, "deck.pdf")

# Translate in place (layout preserved)
engine.translate("deck.pptx", "Spanish")
```

Every call returns a dict with `result.downloadUrl` — a signed link valid for 24 hours. `engine.download(result, path)` fetches it.

## Branding

```python
engine.generate(
    markup="# Deck\n\n## Slide\n- point",
    brand={"primary": "#E4002B", "secondary": "#111827"},
    font="Georgia",
)
```

## Accounts and credits

Pass your account id to tie calls to your quota and remove the watermark (Pro plan):

```python
engine = PowerPointEngine(session_id="your-account-id")
```

Billing is credit-based: 1 credit = up to 10 slides. Free tier: 5 credits/month (watermarked); Pro ($9/mo) removes the watermark. Details: [powerpointengine.io/#pricing](https://powerpointengine.io/#pricing).

## AI agents

The same engine is exposed as a remote [MCP server](https://powerpointengine.io/llms.txt) — one line in Claude Code / any MCP client:

```bash
claude mcp add --transport http powerpoint-engine https://powerpointengine.io/api/mcp/mcp
```

## Errors

All errors raise `PowerPointEngineError` subclasses: `ValidationError` (400), `AuthenticationError` (401), `NotFoundError` (404), `RateLimitError` (429), `ServerError` (5xx).

## License

MIT
