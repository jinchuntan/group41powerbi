#!/usr/bin/env python3
"""Draw the learning and reflection diagram with Matplotlib.

Run from the repository root:
    python scripts/build_reflection_diagram.py

The script needs no workbook. Edit STEPS to change the three slide messages.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.axes import Axes
from matplotlib.patches import Circle, FancyArrowPatch


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "figures" / "recommendations"
FILE_STEM = "04_learning_reflection"
CREAM = "#F4EFE3"
ORANGE = "#D95E26"
INK = "#211D19"
MUTED = "#71695E"

STEPS = (
    ("Check the data", "Match revenue and costs."),
    ("Interpret the results", "Assess ROAS alongside\nproduct costs."),
    ("Plan controlled trials", "Set profit and customer\nexperience targets."),
)


def draw_step(axes: Axes, number: int, title: str, caption: str, y: float) -> None:
    """Place a numbered step with a short explanation on the same row."""
    axes.add_patch(Circle((0.16, y), radius=0.055, facecolor=ORANGE, edgecolor="none"))
    axes.text(
        0.16, y, str(number), ha="center", va="center",
        fontsize=19, fontweight="bold", color=INK,
    )
    axes.text(
        0.29, y + 0.025, title, ha="left", va="center",
        fontsize=20, fontweight="bold", color=INK,
    )
    axes.text(
        0.29, y - 0.025, caption, ha="left", va="top",
        fontsize=15.5, linespacing=1.45, color=MUTED,
    )


def build_diagram(output_dir: Path, dpi: int) -> None:
    """Export a square PNG and an SVG with editable text and no metadata element."""
    figure, axes = plt.subplots(figsize=(7.5, 7.5))
    figure.patch.set_facecolor(CREAM)
    figure.subplots_adjust(left=0, right=1, bottom=0, top=1)
    axes.set(xlim=(0, 1), ylim=(0, 1), aspect="equal")
    axes.set_facecolor("none")
    axes.set_axis_off()
    axes.text(
        0.5, 0.95, "How we reached\nour recommendations",
        ha="center", va="top", fontsize=23, fontweight="bold",
        linespacing=1.25, color=INK,
    )

    top_y, bottom_y = 0.72, 0.20
    spacing = (top_y - bottom_y) / (len(STEPS) - 1)
    for index, (title, caption) in enumerate(STEPS):
        y = top_y - index * spacing
        draw_step(axes, index + 1, title, caption, y)
        if index < len(STEPS) - 1:
            axes.add_patch(FancyArrowPatch(
                (0.16, y - 0.075), (0.16, y - spacing + 0.075),
                arrowstyle="-|>", mutation_scale=18,
                linewidth=1.8, color=ORANGE,
            ))

    output_dir.mkdir(parents=True, exist_ok=True)
    figure.savefig(
        output_dir / f"{FILE_STEM}.png", dpi=dpi,
        facecolor=CREAM, metadata={"Software": None},
    )
    figure.savefig(
        output_dir / f"{FILE_STEM}.svg", transparent=True,
        metadata={"Date": None, "Creator": None, "Format": None, "Type": None},
    )
    plt.close(figure)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--dpi", type=int, default=300)
    args = parser.parse_args()
    if args.dpi < 72:
        parser.error("--dpi must be at least 72")
    if len(STEPS) < 2:
        parser.error("STEPS must contain at least two steps")

    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "svg.fonttype": "none",
        "svg.hashsalt": "nexamart-reflection",
    })
    build_diagram(args.output_dir, args.dpi)
    print(f"Wrote {FILE_STEM}.png and {FILE_STEM}.svg to {args.output_dir}")


if __name__ == "__main__":
    main()
