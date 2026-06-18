# Architecture

## Layering

```
+-------------------+   env vars only
|   __main__.py     |   loads ServerConfig, runs server.run() (stdio)
+---------+---------+
          |
          v
+-------------------+   defines @mcp.tool() functions; never opens files
|     server.py     |   without going through security.assert_within_target
+---------+---------+
          |
          v
+-------------------+   pure-Python wrappers on python-pptx; takes a
|    pptx_ops.py    |   Presentation object, returns plain dicts
+---------+---------+
          |
          v
+-------------------+   the only module that compares paths; raises
|    security.py    |   SecurityError on any escape attempt
+---------+---------+
          |
          v
+-------------------+   immutable ServerConfig built once at startup
|    config.py      |
+-------------------+
```

Each layer only depends on the one below it. `pptx_ops.py` has no MCP
imports so it stays unit-testable in isolation.

## Security boundary

```
   user / MCP client
          |
          v
   +-------------+
   | server tool |---> assert_within_target(path, cfg) ---> ALLOW only if
   +-------------+                                          path == cfg.target_path
          |
          v
   pptx_ops + python-pptx + filesystem
```

There is no second filesystem path elsewhere in the code. If you add a
new tool, route through `assert_within_target` even when you "know" the
path is fine.

## Why env vars over CLI args

MCP clients (Claude Desktop, etc.) launch servers from JSON config.
Environment variables map cleanly onto the `env` block in that JSON,
they survive process restarts without surgery in argv parsing, and they
keep the server's `main()` trivial. Switching to argparse later is a
local change in `__main__.py` only.

## Where iteration 2 hooks go

- New read tool? Add a function in `pptx_ops.py`, register it in
  `server.py`, route through `assert_within_target` if it touches paths.
- New mutation tool? Same, plus call `prs.save(str(target))` at the end
  of the server-level wrapper, the way `add_slide` does.
- New config knob? Extend `ServerConfig` and `load_config_from_env`
  together; never read `os.environ` outside `config.py`.
- New security rule? It belongs in `security.py`. Do not sprinkle path
  checks across tools.
