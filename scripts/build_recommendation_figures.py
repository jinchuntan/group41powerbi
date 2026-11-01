#!/usr/bin/env python3
"""Build the three recommendation charts from NexaMart's raw Excel sheets.

Run from the repository root:
    python scripts/build_recommendation_figures.py

The workbook is read only. PNGs match the slide background, while SVGs have
a transparent background and preserve text as editable text.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter, MaxNLocator, MultipleLocator
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = ROOT / "Grp 41 Kampus Digital Student Workbook (1).xlsx"
DEFAULT_OUTPUT = ROOT / "figures" / "recommendations"
CREAM = "#F4EFE3"
ORANGE = "#D95E26"
AMBER = "#EBAF3D"
GREY = "#B6AD9F"
INK = "#211D19"
MUTED = "#71695E"
GRID = "#DED6C8"
ZERO = Decimal("0")


def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def number(value: Any, field: str) -> Decimal:
    try:
        result = Decimal(text(value))
    except InvalidOperation as error:
        raise ValueError(f"{field} must be numeric; found {value!r}.") from error
    if not result.is_finite():
        raise ValueError(f"{field} must be finite; found {value!r}.")
    return result


def discount(value: Any) -> Decimal:
    raw = text(value)
    result = number(raw[:-1], "Discount") / 100 if raw.endswith("%") else number(value, "Discount")
    if not ZERO <= result <= 1:
        raise ValueError(f"Discount must be between 0 and 1; found {value!r}.")
    return result


def records(workbook: Any, sheet_name: str, required: set[str]) -> list[dict[str, Any]]:
    """Locate the source header and ignore title/navigation columns."""
    sheet = workbook[sheet_name]
    rows = sheet.iter_rows(values_only=True)
    headers: list[str] | None = None
    for row_number, row in enumerate(rows, start=1):
        candidate = [text(value) for value in row]
        if required.issubset(candidate):
            headers = candidate
            break
        if row_number >= 20:
            break
    if headers is None:
        raise ValueError(f"Cannot find the required headers in {sheet_name!r}.")
    result = []
    for row in rows:
        record = {name: value for name, value in zip(headers, row) if name}
        if any(value is not None and text(value) for value in record.values()):
            result.append(record)
    return result


def unique(records_: list[dict[str, Any]], key: str) -> tuple[dict[str, dict[str, Any]], int]:
    """Remove identical duplicates and reject conflicting records for one ID."""
    result: dict[str, dict[str, Any]] = {}
    duplicates = 0
    for record in records_:
        identifier = text(record[key])
        if not identifier:
            raise ValueError(f"Missing {key} in a source record.")
        if identifier in result:
            if result[identifier] != record:
                raise ValueError(f"Conflicting records for {key} {identifier!r}.")
            duplicates += 1
        else:
            result[identifier] = record
    return result, duplicates


def calculate(workbook_path: Path) -> dict[str, Any]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        orders, order_duplicates = unique(records(workbook, "Raw Orders", {
            "Order ID", "Order Date", "Product ID", "Quantity", "Unit Price (RM)",
            "Discount", "Order Status",
        }), "Order ID")
        products, _ = unique(records(workbook, "Raw Products", {
            "Product ID", "Category", "Unit Cost (RM)",
        }), "Product ID")
        campaigns, _ = unique(records(workbook, "Raw Marketing", {
            "Campaign ID", "Campaign Type", "Campaign Spend (RM)",
            "Revenue Generated (RM)",
        }), "Campaign ID")
        experience, _ = unique(records(workbook, "Raw Customer Experience", {
            "Order ID", "Complaint Category", "Returned Order",
        }), "Order ID")
    finally:
        workbook.close()

    channels: dict[str, dict[str, Decimal]] = defaultdict(lambda: {"spend": ZERO, "revenue": ZERO})
    for campaign in campaigns.values():
        channel = text(campaign["Campaign Type"])
        if not channel:
            raise ValueError("A marketing campaign has no channel.")
        spend = number(campaign["Campaign Spend (RM)"], "Campaign Spend (RM)")
        if spend < 0:
            raise ValueError("Campaign spend cannot be negative.")
        channels[channel]["spend"] += spend
        channels[channel]["revenue"] += number(campaign["Revenue Generated (RM)"], "Revenue Generated (RM)")
    marketing = []
    for channel, values in channels.items():
        if values["spend"] <= 0:
            raise ValueError(f"Cannot calculate ROAS for {channel!r} with zero spend.")
        marketing.append({"channel": channel, "spend_rm": float(values["spend"]),
                          "attributed_revenue_rm": float(values["revenue"]),
                          "roas": float(values["revenue"] / values["spend"])})
    marketing.sort(key=lambda row: row["roas"], reverse=True)
    total_spend = sum((row["spend"] for row in channels.values()), ZERO)
    selected_spend = sum((values["spend"] for channel, values in channels.items()
                          if channel in {"Influencer", "Social Media"}), ZERO)

    if set(experience) != set(orders):
        raise ValueError("Order and customer experience IDs must match before plotting complaint rates.")
    labels = {"late delivery": "Late Delivery", "product not suitable": "Product Not Suitable",
              "product quality": "Product Quality"}
    complaints: Counter[str] = Counter()
    returned_orders = 0
    priority_returns = 0
    for record in experience.values():
        raw_label = text(record["Complaint Category"])
        if not raw_label:
            raise ValueError("A complaint category is missing; confirm its meaning before counting.")
        label = labels.get(raw_label.casefold(), raw_label)
        if raw_label.casefold() != "none":
            complaints[label] += 1
        returned = text(record["Returned Order"]).casefold()
        if returned not in {"yes", "no"}:
            raise ValueError(f"Unrecognised Returned Order value: {record['Returned Order']!r}.")
        if returned == "yes":
            returned_orders += 1
            if label in {"Late Delivery", "Product Not Suitable"}:
                priority_returns += 1
    total_complaints = sum(complaints.values())
    if total_complaints == 0:
        raise ValueError("There are no complaints to chart.")
    complaint_data = [{"category": category, "complaints": count}
                      for category, count in complaints.most_common()]
    priority_complaints = complaints["Late Delivery"] + complaints["Product Not Suitable"]

    categories: dict[str, dict[str, Any]] = defaultdict(lambda: {
        "revenue": ZERO, "cost": ZERO, "orders": 0,
    })
    completed_orders = 0
    unknown_cost_orders = 0
    unknown_cost_revenue = ZERO
    for order in orders.values():
        if text(order["Order Status"]).casefold() != "completed":
            continue
        completed_orders += 1
        quantity = number(order["Quantity"], "Quantity")
        price = number(order["Unit Price (RM)"], "Unit Price (RM)")
        if quantity <= 0 or price < 0:
            raise ValueError("Completed orders need a positive quantity and nonnegative unit price.")
        revenue = quantity * price * (1 - discount(order["Discount"]))
        product = products.get(text(order["Product ID"]))
        if product is None or product["Unit Cost (RM)"] is None:
            unknown_cost_orders += 1
            unknown_cost_revenue += revenue
            continue
        category = text(product["Category"])
        if not category:
            raise ValueError("A cost-matched product has no category.")
        unit_cost = number(product["Unit Cost (RM)"], "Unit Cost (RM)")
        if unit_cost < 0:
            raise ValueError("Unit cost cannot be negative.")
        categories[category]["revenue"] += revenue
        categories[category]["cost"] += quantity * unit_cost
        categories[category]["orders"] += 1
    revenue = sum((row["revenue"] for row in categories.values()), ZERO)
    cost = sum((row["cost"] for row in categories.values()), ZERO)
    if revenue <= 0:
        raise ValueError("The cost-matched completed-order revenue must be positive.")
    category_data = []
    for category, values in categories.items():
        if values["revenue"] <= 0:
            raise ValueError(f"Cannot calculate margin for {category!r} with nonpositive revenue.")
        profit = values["revenue"] - values["cost"]
        category_data.append({"category": category, "revenue_rm": float(values["revenue"]),
                              "cost_rm": float(values["cost"]), "gross_profit_rm": float(profit),
                              "completed_orders": values["orders"],
                              "weighted_margin_pct": float(100 * profit / values["revenue"])})
    category_data.sort(key=lambda row: row["weighted_margin_pct"])
    home_office_revenue = categories.get("Home Office", {}).get("revenue", ZERO)
    return {
        "source_workbook": workbook_path.name,
        "source_sha256": hashlib.sha256(workbook_path.read_bytes()).hexdigest(),
        "marketing": marketing,
        "marketing_summary": {"campaigns": len(campaigns), "total_spend_rm": float(total_spend),
                              "influencer_social_spend_pct": float(100 * selected_spend / total_spend)},
        "complaints": complaint_data,
        "service_summary": {"unique_orders": len(orders), "total_complaints": total_complaints,
                            "delivery_suitability_complaints": priority_complaints,
                            "delivery_suitability_complaints_pct": 100 * priority_complaints / total_complaints,
                            "returned_orders": returned_orders,
                            "delivery_suitability_returns": priority_returns},
        "categories": category_data,
        "financial_summary": {"completed_orders": completed_orders,
                              "cost_matched_completed_orders": completed_orders - unknown_cost_orders,
                              "revenue_rm": float(revenue), "cost_rm": float(cost),
                              "gross_profit_rm": float(revenue - cost),
                              "weighted_margin_pct": float(100 * (revenue - cost) / revenue),
                              "home_office_revenue_share_pct": float(100 * home_office_revenue / revenue),
                              "excluded_unknown_cost_orders": unknown_cost_orders,
                              "excluded_unknown_cost_revenue_rm": float(unknown_cost_revenue)},
        "exact_order_duplicates_removed": order_duplicates,
    }


def canvas(title: str, subtitle: str) -> tuple[Any, Any]:
    fig, ax = plt.subplots(figsize=(9, 6.5))
    fig.patch.set_facecolor(CREAM)
    ax.set_facecolor("none")
    fig.subplots_adjust(left=0.31, right=0.92, bottom=0.27, top=0.79)
    fig.text(0.04, 0.94, title, fontsize=21, fontweight="bold", color=INK, va="top")
    fig.text(0.04, 0.875, subtitle, fontsize=12.5, color=MUTED, va="top")
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)
    ax.tick_params(axis="both", length=0, labelsize=12.5, colors=MUTED)
    ax.tick_params(axis="y", colors=INK, pad=10)
    ax.set_axisbelow(True)
    ax.xaxis.grid(True, color=GRID, linewidth=0.8)
    return fig, ax


def bars(ax: Any, labels: list[str], values: list[float], colors: list[str], suffix: str) -> None:
    ax.barh(labels, values, color=colors, height=0.57, edgecolor="none", zorder=3)
    ax.invert_yaxis()
    low, high = min(0, min(values)), max(0, max(values))
    span = high - low or 1
    for index, value in enumerate(values):
        label = f"{value:.2f}{suffix}" if suffix else str(int(value))
        label_position = value + span * 0.025 if value >= 0 else value - span * 0.025
        ax.text(label_position, index, label, va="center", ha="left" if value >= 0 else "right",
                color=INK, fontsize=14, fontweight="bold", zorder=5,
                bbox={"facecolor": CREAM, "edgecolor": "none", "pad": 1.5})
    ax.set_xlim(low - (span * 0.19 if low < 0 else 0), high + span * 0.19)


def save(fig: Any, output_dir: Path, stem: str, dpi: int) -> None:
    fig.savefig(output_dir / f"{stem}.png", dpi=dpi, facecolor=CREAM,
                metadata={"Software": "NexaMart recommendation figures"})
    fig.savefig(output_dir / f"{stem}.svg", transparent=True,
                metadata={"Date": None, "Creator": "NexaMart recommendation figures"})
    plt.close(fig)


def plot_marketing(data: dict[str, Any], output_dir: Path, dpi: int) -> None:
    rows = data["marketing"]
    fig, ax = canvas("Marketing ROAS by channel", "Attributed revenue per RM1 of campaign spend")
    colors = [ORANGE if row["channel"] in {"Email Retargeting", "Campus Activation"} else GREY for row in rows]
    bars(ax, [row["channel"] for row in rows], [row["roas"] for row in rows], colors, "x")
    ax.xaxis.set_major_locator(MultipleLocator(2))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:g}x"))
    ax.set_xlabel("Return on advertising spend (ROAS)", fontsize=12, color=MUTED, labelpad=12)
    share = data["marketing_summary"]["influencer_social_spend_pct"]
    fig.text(0.04, 0.13, f"Influencer and Social Media received {share:.1f}% of campaign spending.",
             fontsize=12.5, color=INK)
    fig.text(0.04, 0.07, "ROAS measures attributed revenue, before product and operating costs.",
             fontsize=11, color=MUTED)
    save(fig, output_dir, "01_marketing_roas", dpi)


def plot_complaints(data: dict[str, Any], output_dir: Path, dpi: int) -> None:
    rows = data["complaints"]
    summary = data["service_summary"]
    fig, ax = canvas("Complaints by category", f"{summary['total_complaints']} complaints across {summary['unique_orders']} unique orders")
    colors = [ORANGE if row["category"] in {"Late Delivery", "Product Not Suitable"} else GREY for row in rows]
    bars(ax, [row["category"] for row in rows], [row["complaints"] for row in rows], colors, "")
    ax.xaxis.set_major_locator(MaxNLocator(integer=True, nbins=5))
    ax.set_xlabel("Number of complaints", fontsize=12, color=MUTED, labelpad=12)
    share = summary["delivery_suitability_complaints_pct"]
    fig.text(0.04, 0.13, f"Delivery and suitability account for {share:.1f}% of complaints.", fontsize=13, color=INK)
    fig.text(0.04, 0.07,
             f"{summary['delivery_suitability_returns']} of {summary['returned_orders']} returned orders fall within these two categories.",
             fontsize=11.5, color=MUTED)
    save(fig, output_dir, "02_customer_complaints", dpi)


def plot_margins(data: dict[str, Any], output_dir: Path, dpi: int) -> None:
    rows = data["categories"]
    summary = data["financial_summary"]
    fig, ax = canvas("Category gross margins", "Weighted margin on completed orders with available costs")
    colors = [ORANGE if row["category"] == "Home Office" else
              AMBER if row["category"] in {"Accessories", "Stationery"} else GREY for row in rows]
    bars(ax, [row["category"] for row in rows], [row["weighted_margin_pct"] for row in rows], colors, "%")
    overall = summary["weighted_margin_pct"]
    ax.axvline(overall, color=INK, linewidth=1.3, linestyle=(0, (4, 4)), zorder=2,
               label=f"Overall {overall:.2f}%")
    ax.legend(loc="upper left", bbox_to_anchor=(0, 1.085), frameon=False, fontsize=11.5,
              labelcolor=INK, handlelength=2.6, borderaxespad=0)
    ax.xaxis.set_major_locator(MultipleLocator(20))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:g}%"))
    ax.set_xlabel("Gross profit / revenue", fontsize=12, color=MUTED, labelpad=12)
    share = summary["home_office_revenue_share_pct"]
    fig.text(0.04, 0.13, f"Home Office contributes {share:.1f}% of cost-matched completed revenue.", fontsize=12.5, color=INK)
    fig.text(0.04, 0.07, "Margins exclude marketing, delivery and overhead costs.", fontsize=11, color=MUTED)
    save(fig, output_dir, "03_category_gross_margin", dpi)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--dpi", type=int, default=300, help="PNG resolution; default 300 DPI")
    args = parser.parse_args()
    if args.dpi < 72:
        parser.error("--dpi must be at least 72")
    if not args.workbook.is_file():
        parser.error(f"Workbook not found: {args.workbook}")
    plt.rcParams.update({"font.family": "DejaVu Sans", "svg.fonttype": "none",
                         "svg.hashsalt": "nexamart-recommendations", "axes.unicode_minus": False})
    data = calculate(args.workbook)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    for plot in (plot_marketing, plot_complaints, plot_margins):
        plot(data, args.output_dir, args.dpi)
    for name, rows in [("marketing_roas", data["marketing"]),
                       ("customer_complaints", data["complaints"]),
                       ("category_gross_margin", data["categories"])]:
        with (args.output_dir / f"{name}.csv").open("w", newline="", encoding="utf-8") as stream:
            writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
            writer.writeheader()
            writer.writerows(rows)
    (args.output_dir / "metrics.json").write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote 3 PNGs, 3 SVGs, 3 CSVs and metrics.json to {args.output_dir}")
    print(f"Marketing budget share: {data['marketing_summary']['influencer_social_spend_pct']:.1f}%")
    print(f"Delivery/suitability complaint share: {data['service_summary']['delivery_suitability_complaints_pct']:.1f}%")
    print(f"Weighted overall gross margin: {data['financial_summary']['weighted_margin_pct']:.2f}%")


if __name__ == "__main__":
    main()
