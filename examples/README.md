# Examples — from clone to first slide in 60 seconds

This folder ships a tiny standalone script that builds a sample template,
plus a walkthrough you can copy-paste.

## What's here

- `build_template.py` — programmatically writes `examples/template.pptx`
  using `python-pptx`'s default master (Title Slide, Title and Content,
  Section Header, and a handful more layouts). Not imported by the
  package; just a fixture-builder.
- `template.pptx` — generated output of the script. Tracked in git as a
  ready-to-use fixture so you do not need to bring your own template
  for a first run.

## 60-second walkthrough

From the repo root:

```bash
# 1. Sync dependencies (one-time).
uv sync

# 2. Materialise the example template (already tracked in git, but you
#    can rebuild it any time):
uv run python examples/build_template.py

# 3. Set up your .env file:
cp .env.example .env
#    Edit .env so the two variables look like (use absolute paths):
#      PPTX_TARGET=/abs/path/to/power-point-mcp/examples/out.pptx
#      PPTX_TEMPLATE=/abs/path/to/power-point-mcp/examples/template.pptx
#    The target file does NOT need to exist yet.

# 4. Confirm the environment is sane:
uv run power-point-mcp --doctor

# 5. Start the server (it speaks MCP over stdio; Ctrl-C to stop):
uv run power-point-mcp
```

Once Claude Desktop (or any MCP client) is wired up to this server, ask
Claude to call `create_presentation_from_template` followed by
`add_slide(layout_name="Title Slide", placeholders={...})`. The server
will read the template you pointed at, write the target deck, and
populate placeholders with the text you provided — and only that text.

## Notes

- The server is bound to a *single* file. `PPTX_TARGET` is the only
  path it will ever touch. Do not point it at a deck you cannot afford
  to overwrite.
- `os.environ` wins over `.env` values. Useful if you want to override
  one variable for a single run:
  `PPTX_TARGET=/tmp/scratch.pptx uv run power-point-mcp`.
