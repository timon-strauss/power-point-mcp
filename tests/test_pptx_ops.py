"""Tests for power_point_mcp.pptx_ops."""

from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation

from power_point_mcp import pptx_ops


def _layouts(path: Path) -> list[str]:
    return [layout.name for layout in Presentation(str(path)).slide_layouts]


def test_create_from_template_copies_layouts(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=False)
    assert target.exists()
    assert _layouts(target) == _layouts(tmp_pptx_template)


def test_create_from_template_no_overwrite(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=False)
    with pytest.raises(FileExistsError):
        pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=False)
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)


def test_summarize_zero_then_one_slide(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)

    prs = Presentation(str(target))
    info = pptx_ops.summarize_presentation(prs)
    assert info["slide_count"] == 0
    assert "Title Slide" in info["layouts"]

    pptx_ops.add_slide(prs, "Title Slide", {})
    prs.save(str(target))

    prs2 = Presentation(str(target))
    info2 = pptx_ops.summarize_presentation(prs2)
    assert info2["slide_count"] == 1


def test_add_slide_fills_only_supplied_placeholders(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))

    # Find the actual placeholder name on the Title Slide layout.
    layout = next(l for l in prs.slide_layouts if l.name == "Title Slide")
    title_ph_name = None
    for ph in layout.placeholders:
        if ph.placeholder_format.idx == 0:
            title_ph_name = ph.name
            break
    assert title_ph_name is not None

    result = pptx_ops.add_slide(prs, "Title Slide", {title_ph_name: "Hello"})
    assert result["populated"] == [title_ph_name]
    assert result["warnings"] == []

    result2 = pptx_ops.add_slide(
        prs, "Title Slide", {title_ph_name: "Hi", "Bogus": "X"}
    )
    assert title_ph_name in result2["populated"]
    assert any("Bogus" in w for w in result2["warnings"])


def test_read_slide_returns_title_text(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))

    layout = next(l for l in prs.slide_layouts if l.name == "Title Slide")
    title_ph_name = next(
        ph.name for ph in layout.placeholders if ph.placeholder_format.idx == 0
    )
    pptx_ops.add_slide(prs, "Title Slide", {title_ph_name: "Hello there"})
    prs.save(str(target))

    prs2 = Presentation(str(target))
    detail = pptx_ops.read_slide(prs2, 0)
    title_texts = [s["text"] for s in detail["shapes"] if s["is_placeholder"]]
    assert any("Hello there" in t for t in title_texts)


def test_read_slide_out_of_range(tmp_path: Path, tmp_pptx_template: Path) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))
    with pytest.raises(IndexError):
        pptx_ops.read_slide(prs, 0)
    with pytest.raises(IndexError):
        pptx_ops.read_slide(prs, 99)


def _add_title_slide(prs) -> str:
    """Add a Title-Slide and return the title placeholder name."""
    layout = next(l for l in prs.slide_layouts if l.name == "Title Slide")
    title_ph_name = next(
        ph.name for ph in layout.placeholders if ph.placeholder_format.idx == 0
    )
    pptx_ops.add_slide(prs, "Title Slide", {})
    return title_ph_name


def test_set_slide_placeholder_by_name(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))
    title_ph_name = _add_title_slide(prs)

    result = pptx_ops.set_slide_placeholder(prs, 0, title_ph_name, "Hello name")
    assert result["new_text"] == "Hello name"
    assert result["previous_text"] == ""
    assert result["placeholder_name"] == title_ph_name
    assert result["warnings"] == []

    # second call captures previous_text correctly
    result2 = pptx_ops.set_slide_placeholder(prs, 0, title_ph_name, "Replaced")
    assert result2["previous_text"] == "Hello name"
    assert result2["new_text"] == "Replaced"


def test_set_slide_placeholder_unknown_name(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))
    _add_title_slide(prs)
    with pytest.raises(KeyError):
        pptx_ops.set_slide_placeholder(prs, 0, "NoSuchPlaceholder", "x")


def test_set_slide_placeholder_out_of_range(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))
    with pytest.raises(IndexError):
        pptx_ops.set_slide_placeholder(prs, 0, "anything", "x")


def test_set_slide_placeholder_by_idx(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))
    _add_title_slide(prs)

    result = pptx_ops.set_slide_placeholder_by_idx(prs, 0, 0, "By idx")
    assert result["new_text"] == "By idx"
    assert result["placeholder_idx"] == 0


def test_set_slide_placeholder_by_idx_unknown(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))
    _add_title_slide(prs)
    with pytest.raises(KeyError):
        pptx_ops.set_slide_placeholder_by_idx(prs, 0, 999, "x")


def test_set_slide_title(tmp_path: Path, tmp_pptx_template: Path) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))
    _add_title_slide(prs)

    result = pptx_ops.set_slide_title(prs, 0, "The Title")
    assert result["new_text"] == "The Title"
    assert result["previous_text"] == ""

    # round-trip via save/reopen
    prs.save(str(target))
    prs2 = Presentation(str(target))
    detail = pptx_ops.read_slide(prs2, 0)
    title_texts = [s["text"] for s in detail["shapes"] if s["is_placeholder"]]
    assert any("The Title" in t for t in title_texts)


def test_set_slide_title_no_title_layout(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    """If a layout has no idx-0 placeholder, set_slide_title should KeyError."""
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))

    # find a layout that has no placeholder with idx 0; skip otherwise
    no_title_layout = None
    for layout in prs.slide_layouts:
        idxs = [ph.placeholder_format.idx for ph in layout.placeholders]
        if 0 not in idxs:
            no_title_layout = layout
            break
    if no_title_layout is None:
        pytest.skip("no layout without idx-0 placeholder available")

    pptx_ops.add_slide(prs, no_title_layout.name, {})
    with pytest.raises(KeyError):
        pptx_ops.set_slide_title(prs, 0, "Title")


def test_list_layouts_returns_layouts_with_placeholders(
    tmp_path: Path, tmp_pptx_template: Path
) -> None:
    target = tmp_path / "out.pptx"
    pptx_ops.create_from_template(tmp_pptx_template, target, overwrite=True)
    prs = Presentation(str(target))

    layouts = pptx_ops.list_layouts(prs)
    assert isinstance(layouts, list)
    assert len(layouts) >= 3
    names = {entry["name"] for entry in layouts}
    assert "Title Slide" in names
    for entry in layouts:
        assert "name" in entry
        assert "idx" in entry
        assert "placeholder_names" in entry
        assert isinstance(entry["placeholder_names"], list)
        assert isinstance(entry["idx"], int)
    # idx values should be 0..N-1 in order
    assert [entry["idx"] for entry in layouts] == list(range(len(layouts)))
