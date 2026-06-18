# SETUP — power-point-mcp

Step-by-step guide for getting the server running on a fresh machine and
wiring it into Claude Desktop.

If you have done this before, the `README.md` quick-start is enough. This
document is for the first time, or for debugging a broken install.

---

## 1. Prerequisites

| Tool       | Version  | Purpose                          | Install hint                                |
| ---------- | -------- | -------------------------------- | ------------------------------------------- |
| Python     | **3.14** | Runtime                          | macOS: `brew install python@3.14`           |
| `uv`       | latest   | Project & dependency management  | `brew install uv` or [astral.sh/uv](https://docs.astral.sh/uv/) |
| `git`      | any      | Cloning the repo                 | macOS: pre-installed via Xcode CLT           |

Verify each before continuing:

```bash
python3 --version    # → Python 3.14.x
uv --version         # → uv 0.x.y
git --version        # → git version 2.x.y
```

> **Why Python 3.14?** It is the version pinned in `.python-version`. `uv`
> will fetch a matching interpreter automatically if your system Python is
> different — you do **not** have to install 3.14 system-wide.

---

## 2. Clone and sync

```bash
git clone <this-repo-url> power-point-mcp
cd power-point-mcp

# Runtime dependencies (mcp, python-pptx) into a local .venv:
uv sync

# Optional — only if you want to run the test suite:
uv sync --group dev
```

After `uv sync`, the project layout looks like this:

```
power-point-mcp/
├── .venv/                  # local virtualenv (gitignored)
├── src/power_point_mcp/    # the server
├── tests/                  # pytest suite
├── docs/                   # plan & architecture docs
├── pyproject.toml
├── uv.lock
├── README.md / SETUP.md / CLAUDE.md / changelog.md
└── .python-version
```

Confirm the install:

```bash
uv run python -c "import power_point_mcp; print(power_point_mcp.__version__)"
# → 0.2.0
```

---

## 3. Pick your target file and (optional) template

The server is **bound to one `.pptx` at startup**. It cannot read or write
any other file. You decide which file by setting environment variables:

| Variable        | Required | What it points to                                           |
| --------------- | -------- | ----------------------------------------------------------- |
| `PPTX_TARGET`   | yes      | Absolute path to the deck the server may read/write.        |
| `PPTX_TEMPLATE` | no       | Absolute path to a `.pptx` or `.potx` template, if any.     |

The target file does **not** have to exist yet — call
`create_presentation_from_template` to initialise it from your template.

---

## 4. Run the server (manual smoke test)

Before wiring it into Claude Desktop, prove it starts cleanly:

```bash
PPTX_TARGET=/abs/path/to/deck.pptx \
PPTX_TEMPLATE=/abs/path/to/template.pptx \
uv run power-point-mcp
```

You should see no output (the server speaks MCP over stdio and is waiting
for a client). Press **Ctrl+C** to stop. If you see an error like
`PPTX_TARGET is not set`, double-check the env var.

---

## 5. Wire it into Claude Desktop

Edit your Claude Desktop MCP config. On macOS the file lives at
`~/Library/Application Support/Claude/claude_desktop_config.json`.

Add (or merge) the following:

```json
{
  "mcpServers": {
    "power-point-mcp": {
      "command": "uv",
      "args": [
        "run",
        "--project",
        "/abs/path/to/power-point-mcp",
        "power-point-mcp"
      ],
      "env": {
        "PPTX_TARGET": "/abs/path/to/deck.pptx",
        "PPTX_TEMPLATE": "/abs/path/to/template.pptx"
      }
    }
  }
}
```

Restart Claude Desktop. In a new conversation you should see the
`power-point-mcp` server connected with ten tools (`presentation_info`,
`list_slides`, `read_slide`, `create_presentation_from_template`,
`add_slide`, `set_slide_placeholder`, `set_slide_placeholder_by_idx`,
`set_slide_title`, `delete_slide`, `reorder_slide`).

> Use `--project` so `uv` resolves the right project even when Claude
> Desktop launches it from an unrelated working directory.

---

## 6. (Recommended) Pair it with academic-pptx-skill

The MCP server provides the **mechanism** (read, create, modify slides).
For high-quality slide *design* — layouts, typography, narrative flow —
install the companion skill:

[github.com/Gabberflast/academic-pptx-skill](https://github.com/Gabberflast/academic-pptx-skill)

Together: the skill tells Claude *how* a good slide looks; this MCP lets
Claude actually write it to your file.

---

## 7. Run the tests (optional)

```bash
uv sync --group dev
uv run pytest -q
# → 39 passed in <1 s
```

The suite is fast and runs entirely against `tmp_path` fixtures — it does
**not** touch your real `PPTX_TARGET`.

---

## Troubleshooting

**“PPTX_TARGET is not set”** — the server refuses to start without a
target. Set the env var; for Claude Desktop, set it in `env` inside the
MCP config block.

**“path is outside the bound target”** — a tool tried to operate on a
file other than `PPTX_TARGET`. This is the security boundary doing its
job. If you actually want to work on a different file, restart the
server with a different `PPTX_TARGET`.

**“no placeholder named X on slide N”** — `set_slide_placeholder` could
not find that placeholder. Use `read_slide(slide_index=N)` to see the
real shape names; or use `set_slide_placeholder_by_idx` with the layout's
placeholder idx.

**“no title placeholder”** — `set_slide_title` only works on slides
whose layout has a title placeholder. Use a layout that includes one
(most templates do) or use `add_slide` to create such a slide first.

**Server starts but Claude Desktop shows no tools** — restart Claude
Desktop fully (quit + relaunch). Make sure `command` is `uv`, not
`uv.exe` on Windows or a relative path. Run the manual smoke test in
step 4 to isolate the issue.

**Tests fail with `ModuleNotFoundError: No module named 'pytest'`** —
you ran `uv sync` but not `uv sync --group dev`. Re-run with the group
flag.
