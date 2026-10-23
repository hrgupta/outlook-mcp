"""Docs must not drift from the tool surface.

README.md and SKILL.md are what a visitor and an MCP client read *before*
installing. A stale tool count there is the same class of bug as a stale schema:
the server is honest and the documentation is not. These assertions read the
live registration and the prose together, so adding or removing a tool fails CI
until the docs are updated with it.

Counts are parsed out of the docs rather than hardcoded here on purpose — a
hardcoded number in the test would drift exactly like the one in the README.
"""

import re
from pathlib import Path

from outlook_mcp.permissions import VALID_CATEGORIES
from outlook_mcp.server import mcp

ROOT = Path(__file__).resolve().parent.parent
README = (ROOT / "README.md").read_text(encoding="utf-8")
SKILL = (ROOT / "SKILL.md").read_text(encoding="utf-8")


def _registered_tools() -> set[str]:
    return set(mcp._tool_manager._tools.keys())


def _feature_block() -> str:
    """The README's '**N tools** across M categories:' listing, up to the next section."""
    match = re.search(r"\*\*\d+ tools\*\* across \d+ categories:", README)
    assert match, "README no longer states '**N tools** across M categories:'"
    start = match.start()
    end = README.index("**Design principles:**", start)
    return README[start:end]


def test_readme_total_matches_registration():
    match = re.search(r"\*\*(\d+) tools\*\* across (\d+) categories:", README)
    assert match, "README no longer states '**N tools** across M categories:'"
    assert int(match.group(1)) == len(_registered_tools()), (
        f"README claims {match.group(1)} tools, server registers {len(_registered_tools())}"
    )


def test_readme_category_bullets_match_registration():
    block = _feature_block()
    counts = [int(n) for n in re.findall(r"^- \*\*[^(]+\((\d+)\)", block, re.M)]
    header = re.search(r"across (\d+) categories", block)
    assert header, "README category header lost its count"
    assert len(counts) == int(header.group(1)), (
        f"README says {header.group(1)} categories but lists {len(counts)}"
    )
    assert sum(counts) == len(_registered_tools()), (
        f"README category counts sum to {sum(counts)}, server registers {len(_registered_tools())}"
    )


def test_readme_permission_table_matches_code():
    section = README.split("**Available categories:**")[1].split("**Example policies:**")[0]
    documented = set(re.findall(r"^\| `([a-z_]+)`", section, re.M))
    assert documented == set(VALID_CATEGORIES), (
        f"README documents {sorted(documented)}, code has {sorted(VALID_CATEGORIES)}"
    )


def test_skill_tool_heading_matches_registration():
    match = re.search(r"^## Tools \((\d+)\)", SKILL, re.M)
    assert match, "SKILL.md no longer has a '## Tools (N)' heading"
    assert int(match.group(1)) == len(_registered_tools()), (
        f"SKILL.md heading claims {match.group(1)} tools, server registers "
        f"{len(_registered_tools())}"
    )


def test_skill_description_matches_registration():
    match = re.search(r"^description: .*?(\d+) typed Graph tools", SKILL, re.M)
    assert match, "SKILL.md description no longer states 'N typed Graph tools'"
    assert int(match.group(1)) == len(_registered_tools())


# Every prose mention of a count, with what it must equal. The first version of
# this file checked only the two tool-count phrases below that happened to be on
# hand, so "this has N tools", "N discrete tools" and "the N categories" could go
# stale while every test passed. Listing the patterns explicitly keeps a new
# sentence from silently escaping the net: a phrase that disappears fails the
# "prose count missing" assert rather than being skipped.
TOOL_COUNT_PATTERNS = (
    r"All (\d+) tool schemas load",
    r"unset loads all (\d+)",
    r"this has (\d+) tools",
    r"(\d+) discrete tools",
)
# "across N categories" is the README's documentation grouping (asserted against the
# bullet list above); "the N categories" is the permission categories in SKILL.md,
# which track VALID_CATEGORIES. Two different counts, so two different tests.
PERMISSION_CATEGORY_PATTERNS = (r"the (\d+) categories",)


def test_other_prose_counts_match_registration():
    """Every prose mention of the tool count agrees with the registration."""
    total = len(_registered_tools())
    for pattern in TOOL_COUNT_PATTERNS:
        match = re.search(pattern, README + SKILL)
        assert match, f"prose count missing: {pattern}"
        assert int(match.group(1)) == total, (
            f"{pattern!r} says {match.group(1)}, server registers {total}"
        )


def test_prose_category_counts_match_code():
    """Category counts in the prose agree with VALID_CATEGORIES."""
    expected = len(VALID_CATEGORIES)
    for pattern in PERMISSION_CATEGORY_PATTERNS:
        match = re.search(pattern, README + SKILL)
        assert match, f"prose category count missing: {pattern}"
        assert int(match.group(1)) == expected, (
            f"{pattern!r} says {match.group(1)}, code defines {expected}"
        )
