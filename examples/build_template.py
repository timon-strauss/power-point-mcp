"""Build examples/template.pptx — a tiny three-layout fixture.

Run from the repo root:

    uv run python examples/build_template.py

This script is **not** part of the package. It exists so a fresh clone has a
working template to point PPTX_TEMPLATE at.

The output file uses python-pptx's default master/layouts (which already
include "Title Slide", "Title and Content", and "Section Header" among
others). We deliberately keep this minimal: python-pptx ships sensible
placeholder defaults, and adding our own styling would just fight the
template the user actually wants to bring.
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation


REQUIRED_LAYOUTS = ("Title Slide", "Title and Content", "Section Header")


def build_template(out_path: Path) -> Path:
    """Materialise a fresh .pptx with python-pptx's default layouts."""
    prs = Presentation()
    available = {layout.name for layout in prs.slide_layouts}
    missing = [name for name in REQUIRED_LAYOUTS if name not in available]
    if missing:
        raise RuntimeError(
            f"python-pptx default master is missing required layouts: {missing}. "
            f"Available layouts: {sorted(available)}"
        )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def main() -> None:
    here = Path(__file__).resolve().parent
    out = build_template(here / "template.pptx")
    print(f"Wrote {out}")
    prs = Presentation(str(out))
    print("Layouts in template:")
    for idx, layout in enumerate(prs.slide_layouts):
        print(f"  [{idx}] {layout.name}")


if __name__ == "__main__":
    main()
