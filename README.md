# power-point-mcp

An MCP (Model Context Protocol) server that lets Claude read, create, and
edit **a single, pre-bound PowerPoint file**. Bind the server to one
`.pptx` and one optional template via environment variables; the server
refuses to touch anything else on disk. Claude can then ask for slide
metadata, list slides, read a slide's shapes, create the file from your
template, and add new slides whose text comes only from the user.

## Prerequisites

- Python **3.14** (matches `.python-version`).
- [`uv`](https://docs.astral.sh/uv/) for environment and dependency
  management.

That is the entire toolchain. Nothing is installed globally.

## Setup

```bash
git clone <this-repo-url> power-point-mcp
cd power-point-mcp
uv sync                  # runtime deps into a local .venv
uv sync --group dev      # add pytest for running the test suite
```

## Configuration

The server is bound to a single PPTX file via environment variables:

| Variable        | Required | Purpose                                                 |
| --------------- | -------- | ------------------------------------------------------- |
| `PPTX_TARGET`   | yes      | Absolute path to the one `.pptx` the server may touch.  |
| `PPTX_TEMPLATE` | no       | Absolute path to a `.pptx` or `.potx` template, if any. |

If `PPTX_TARGET` is unset, the server refuses to start.

## Running locally

```bash
PPTX_TARGET=/abs/path/to/deck.pptx \
PPTX_TEMPLATE=/abs/path/to/template.pptx \
uv run power-point-mcp
```

The server speaks MCP over stdio (FastMCP's default transport).

## Claude Desktop configuration

Add an entry like this to your Claude Desktop MCP config:

```json
{
  "mcpServers": {
    "power-point-mcp": {
      "command": "uv",
      "args": ["run", "power-point-mcp"],
      "env": {
        "PPTX_TARGET": "/abs/path/to/deck.pptx",
        "PPTX_TEMPLATE": "/abs/path/to/template.pptx"
      }
    }
  }
}
```

## Tools exposed

- `presentation_info()` — slide count, dimensions, layouts, master.
- `list_slides()` — index, layout, title, text snippet for every slide.
- `read_slide(slide_index)` — every shape on one slide, with its text.
- `create_presentation_from_template(overwrite=False)` — initialise the
  bound target from `PPTX_TEMPLATE`.
- `add_slide(layout_name, placeholders)` — append a slide using a named
  layout; only fills placeholders you supply.

## Recommended companion skill

Pair this server with the
[academic-pptx-skill](https://github.com/Gabberflast/academic-pptx-skill).
It provides the slide-design conventions and prompts; this MCP gives
Claude the file-system hands to actually write the deck.

## Security note

The server only ever reads or writes `PPTX_TARGET`. Every path the tools
receive is routed through a single security check that compares it
against the bound target and rejects anything else. There is no
network code in the package.

## Development

```bash
uv sync --group dev
uv run pytest
```

See `docs/architecture.md` for the layering and where to add new tools.
