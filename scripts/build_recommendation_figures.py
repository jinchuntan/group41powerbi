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
from matplotlib.axes import Axes
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter, MaxNLocator, MultipleLocator
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook


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
Record = dict[str, Any]
RecordsById = dict[str, Record]
TEST_DESTINATION_CHANNELS = {"Email Retargeting", "Campus Activation"}
TEST_SOURCE_CHANNELS = {"Influencer", "Social Media"}
SERVICE_PRIORITY_CATEGORIES = {"Late Delivery", "Product Not Suitable"}
ADD_ON_CATEGORIES = {"Accessories", "Stationery"}


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def parse_number(value: Any, field: str) -> Decimal:
    """Keep financial arithmetic decimal and reject unavailable source values."""
    try:
        result = Decimal(clean_text(value))
    except InvalidOperation as error:
        raise ValueError(f"{field} must be numeric; found {value!r}.") from error
    if not result.is_finite():
        raise ValueError(f"{field} must be finite; found {value!r}.")
    return result


def parse_discount(value: Any) -> Decimal:
    """Accept Excel percentages and source text such as '10%'."""
    raw = clean_text(value)
    if raw.endswith("%"):
        result = parse_number(raw[:-1], "Discount") / 100
    else:
        result = parse_number(value, "Discount")
    if not ZERO <= result <= 1:
        raise ValueError(f"Discount must be between 0 and 1; found {value!r}.")
    return result


def read_sheet_records(
    workbook: Workbook, sheet_name: str, required_columns: set[str]
) -> list[Record]:
    """Locate the source header and ignore title/navigation columns."""
    sheet = workbook[sheet_name]
    rows = sheet.iter_rows(values_only=True)
    headers: list[str] | None = None
    for row_number, row in enumerate(rows, start=1):
        candidate = [clean_text(value) for value in row]
        if required_columns.issubset(candidate):
            headers = candidate
            break
        if row_number >= 20:
            break
    if headers is None:
        raise ValueError(f"Cannot find the required headers in {sheet_name!r}.")
    result = []
    for row in rows:
        record = {name: value for name, value in zip(headers, row) if name}
        if any(clean_text(value) for value in record.values()):
            result.append(record)
    return result


def deduplicate_records(
    source_records: list[Record], key: str
) -> tuple[RecordsById, int]:
    """Remove identical duplicates and reject conflicting records for one ID."""
    result: RecordsById = {}
    duplicates = 0
    for record in source_records:
        identifier = clean_text(record[key])
        if not identifier:
            raise ValueError(f"Missing {key} in a source record.")
        if identifier in result:
            if result[identifier] != record:
                raise ValueError(f"Conflicting records for {key} {identifier!r}.")
            duplicates += 1
        else:
            result[identifier] = record
    return result, duplicates


def calculate_marketing_metrics(
    campaigns: RecordsById,
) -> tuple[list[Record], Record]:
    """Calculate each channel's ROAS from its total revenue and total spend."""
    if not campaigns:
        raise ValueError("There are no marketing campaigns to chart.")
    channels: dict[str, dict[str, Decimal]] = defaultdict(lambda: {"spend": ZERO, "revenue": ZERO})
    for campaign in campaigns.values():
        channel = clean_text(campaign["Campaign Type"])
        if not channel:
            raise ValueError("A marketing campaign has no channel.")
        spend = parse_number(campaign["Campaign Spend (RM)"], "Campaign Spend (RM)")
        if spend < 0:
            raise ValueError("Campaign spend cannot be negative.")
        channels[channel]["spend"] += spend
        channels[channel]["revenue"] += parse_number(
            campaign["Revenue Generated (RM)"], "Revenue Generated (RM)"
        )
    marketing = []
    for channel, values in channels.items():
        if values["spend"] <= 0:
            raise ValueError(f"Cannot calculate ROAS for {channel!r} with zero spend.")
        marketing.append({
            "channel": channel,
            "spend_rm": float(values["spend"]),
            "attributed_revenue_rm": float(values["revenue"]),
            "roas": float(values["revenue"] / values["spend"]),
        })
    marketing.sort(key=lambda row: row["roas"], reverse=True)
    total_spend = sum((row["spend"] for row in channels.values()), ZERO)
    selected_spend = sum(
        (values["spend"] for channel, values in channels.items()
         if channel in TEST_SOURCE_CHANNELS),
        ZERO,
    )
    summary = {
        "campaigns": len(campaigns),
        "total_spend_rm": float(total_spend),
        "influencer_social_spend_pct": float(100 * selected_spend / total_spend),
    }
    return marketing, summary


def calculate_service_metrics(
    orders: RecordsById, experience: RecordsById
) -> tuple[list[Record], Record]:
    """Count complaints across every unique order, including financial exclusions."""

    if set(experience) != set(orders):
        raise ValueError("Order and customer experience IDs must match before plotting complaint rates.")
    labels = {
        "late delivery": "Late Delivery",
        "product not suitable": "Product Not Suitable",
        "product quality": "Product Quality",
    }
    complaints: Counter[str] = Counter()
    returned_orders = 0
    priority_returns = 0
    for record in experience.values():
        raw_label = clean_text(record["Complaint Category"])
        if not raw_label:
            raise ValueError("A complaint category is missing; confirm its meaning before counting.")
        label = labels.get(raw_label.casefold(), raw_label)
        if raw_label.casefold() != "none":
            complaints[label] += 1
        returned = clean_text(record["Returned Order"]).casefold()
        if returned not in {"yes", "no"}:
            raise ValueError(f"Unrecognised Returned Order value: {record['Returned Order']!r}.")
        if returned == "yes":
            returned_orders += 1
            if label in SERVICE_PRIORITY_CATEGORIES:
                priority_returns += 1
    total_complaints = sum(complaints.values())
    if total_complaints == 0:
        raise ValueError("There are no complaints to chart.")
    complaint_data = [
        {"category": category, "complaints": count}
        for category, count in complaints.most_common()
    ]
    priority_complaints = sum(complaints[category] for category in SERVICE_PRIORITY_CATEGORIES)
    summary = {
        "unique_orders": len(orders),
        "total_complaints": total_complaints,
        "delivery_suitability_complaints": priority_complaints,
        "delivery_suitability_complaints_pct": 100 * priority_complaints / total_complaints,
        "returned_orders": returned_orders,
        "delivery_suitability_returns": priority_returns,
    }
    return complaint_data, summary


def calculate_financial_metrics(
    orders: RecordsById, products: RecordsById
) -> tuple[list[Record], Record]:
    """Use matching revenue and costs so missing-cost orders cannot dilute margin."""

    categories: dict[str, dict[str, Any]] = defaultdict(lambda: {
        "revenue": ZERO, "cost": ZERO, "orders": 0,
    })
    completed_orders = 0
    unknown_cost_orders = 0
    unknown_cost_revenue = ZERO
    for order in orders.values():
        if clean_text(order["Order Status"]).casefold() != "completed":
            continue
        completed_orders += 1
        quantity = parse_number(order["Quantity"], "Quantity")
        price = parse_number(order["Unit Price (RM)"], "Unit Price (RM)")
        if quantity <= 0 or price < 0:
            raise ValueError("Completed orders need a positive quantity and nonnegative unit price.")
        revenue = quantity * price * (1 - parse_discount(order["Discount"]))
        product = products.get(clean_text(order["Product ID"]))
        if product is None or product["Unit Cost (RM)"] is None:
            unknown_cost_orders += 1
            unknown_cost_revenue += revenue
            continue
        category = clean_text(product["Category"])
        if not category:
            raise ValueError("A cost-matched product has no category.")
        unit_cost = parse_number(product["Unit Cost (RM)"], "Unit Cost (RM)")
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
        category_data.append({
            "category": category,
            "revenue_rm": float(values["revenue"]),
            "cost_rm": float(values["cost"]),
            "gross_profit_rm": float(profit),
            "completed_orders": values["orders"],
            "weighted_margin_pct": float(100 * profit / values["revenue"]),
        })
    category_data.sort(key=lambda row: row["weighted_margin_pct"])
    home_office_revenue = categories.get("Home Office", {}).get("revenue", ZERO)
    summary = {
        "completed_orders": completed_orders,
        "cost_matched_completed_orders": completed_orders - unknown_cost_orders,
        "revenue_rm": float(revenue),
        "cost_rm": float(cost),
        "gross_profit_rm": float(revenue - cost),
        "weighted_margin_pct": float(100 * (revenue - cost) / revenue),
        "home_office_revenue_share_pct": float(100 * home_office_revenue / revenue),
        "excluded_unknown_cost_orders": unknown_cost_orders,
        "excluded_unknown_cost_revenue_rm": float(unknown_cost_revenue),
    }
    return category_data, summary


def build_metrics(workbook_path: Path) -> Record:
    """Load the four required raw tables and assemble the chart data."""
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        orders, order_duplicates = deduplicate_records(read_sheet_records(
            workbook, "Raw Orders", {
                "Order ID", "Product ID", "Quantity", "Unit Price (RM)",
                "Discount", "Order Status",
            }
        ), "Order ID")
        products, _ = deduplicate_records(read_sheet_records(
            workbook, "Raw Products", {"Product ID", "Category", "Unit Cost (RM)"}
        ), "Product ID")
        campaigns, _ = deduplicate_records(read_sheet_records(
            workbook, "Raw Marketing", {
                "Campaign ID", "Campaign Type", "Campaign Spend (RM)",
                "Revenue Generated (RM)",
            }
        ), "Campaign ID")
        experience, _ = deduplicate_records(read_sheet_records(
            workbook, "Raw Customer Experience", {
                "Order ID", "Complaint Category", "Returned Order",
            }
        ), "Order ID")
    finally:
        workbook.close()

    marketing, marketing_summary = calculate_marketing_metrics(campaigns)
    complaints, service_summary = calculate_service_metrics(orders, experience)
    categories, financial_summary = calculate_financial_metrics(orders, products)
    return {
        "source_workbook": workbook_path.name,
        "source_sha256": hashlib.sha256(workbook_path.read_bytes()).hexdigest(),
        "marketing": marketing,
        "marketing_summary": marketing_summary,
        "complaints": complaints,
        "service_summary": service_summary,
        "categories": categories,
        "financial_summary": financial_summary,
        "exact_order_duplicates_removed": order_duplicates,
    }


def create_chart_canvas(
    title: str, x_label: str, y_label: str
) -> tuple[Figure, Axes]:
    """Keep slide figures limited to the chart title, axes and plotted values."""
    fig, ax = plt.subplots(figsize=(9, 6.5))
    fig.patch.set_facecolor(CREAM)
    ax.set_facecolor("none")
    fig.subplots_adjust(left=0.32, right=0.93, bottom=0.17, top=0.85)
    fig.suptitle(title, y=0.96, fontsize=22, fontweight="bold", color=INK)
    ax.set_xlabel(x_label, fontsize=14, color=INK, labelpad=14)
    ax.set_ylabel(y_label, fontsize=14, color=INK, labelpad=16)
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)
    ax.tick_params(axis="both", length=0, labelsize=14, colors=MUTED)
    ax.tick_params(axis="y", colors=INK, pad=10)
    ax.set_axisbelow(True)
    ax.xaxis.grid(True, color=GRID, linewidth=0.8)
    return fig, ax


def draw_horizontal_bars(
    ax: Axes, labels: list[str], values: list[float], colors: list[str], suffix: str
) -> None:
    ax.barh(labels, values, color=colors, height=0.57, edgecolor="none", zorder=3)
    ax.invert_yaxis()
    low, high = min(0, min(values)), max(0, max(values))
    span = high - low or 1
    for index, value in enumerate(values):
        label = f"{value:.2f}{suffix}" if suffix else str(int(value))
        label_position = value + span * 0.025 if value >= 0 else value - span * 0.025
        ax.text(label_position, index, label, va="center", ha="left" if value >= 0 else "right",
                color=INK, fontsize=16, fontweight="bold", zorder=5,
                bbox={"facecolor": CREAM, "edgecolor": "none", "pad": 1.5})
    ax.set_xlim(low - (span * 0.19 if low < 0 else 0), high + span * 0.19)


def save_chart(fig: Figure, output_dir: Path, stem: str, dpi: int) -> None:
    """Omit optional image metadata while retaining PNG resolution and SVG text."""
    fig.savefig(output_dir / f"{stem}.png", dpi=dpi, facecolor=CREAM,
                metadata={"Software": None})
    fig.savefig(output_dir / f"{stem}.svg", transparent=True,
                metadata={"Date": None, "Creator": None, "Format": None, "Type": None})
    plt.close(fig)


def plot_marketing(data: dict[str, Any], output_dir: Path, dpi: int) -> None:
    rows = data["marketing"]
    fig, ax = create_chart_canvas(
        "Marketing ROAS by channel", "Revenue per RM1 spent (ROAS)", "Marketing channel"
    )
    colors = [
        ORANGE if row["channel"] in TEST_DESTINATION_CHANNELS else GREY
        for row in rows
    ]
    draw_horizontal_bars(
        ax, [row["channel"] for row in rows],
        [row["roas"] for row in rows], colors, "x"
    )
    ax.xaxis.set_major_locator(MultipleLocator(2))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:g}x"))
    save_chart(fig, output_dir, "01_marketing_roas", dpi)


def plot_complaints(data: dict[str, Any], output_dir: Path, dpi: int) -> None:
    rows = data["complaints"]
    fig, ax = create_chart_canvas(
        "Complaints by category", "Number of complaints", "Complaint category"
    )
    colors = [
        ORANGE if row["category"] in SERVICE_PRIORITY_CATEGORIES else GREY
        for row in rows
    ]
    draw_horizontal_bars(
        ax, [row["category"] for row in rows],
        [row["complaints"] for row in rows], colors, ""
    )
    ax.xaxis.set_major_locator(MaxNLocator(integer=True, nbins=5))
    save_chart(fig, output_dir, "02_customer_complaints", dpi)


def plot_margins(data: dict[str, Any], output_dir: Path, dpi: int) -> None:
    rows = data["categories"]
    fig, ax = create_chart_canvas(
        "Gross margin by category", "Weighted gross margin (%)", "Product category"
    )
    colors = []
    for row in rows:
        if row["category"] == "Home Office":
            colors.append(ORANGE)
        elif row["category"] in ADD_ON_CATEGORIES:
            colors.append(AMBER)
        else:
            colors.append(GREY)
    draw_horizontal_bars(
        ax, [row["category"] for row in rows],
        [row["weighted_margin_pct"] for row in rows], colors, "%"
    )
    ax.xaxis.set_major_locator(MultipleLocator(20))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:g}%"))
    save_chart(fig, output_dir, "03_category_gross_margin", dpi)


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
    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "svg.fonttype": "none",
        "svg.hashsalt": "nexamart-recommendations",
        "axes.unicode_minus": False,
    })
    data = build_metrics(args.workbook)
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
    complaint_share = data["service_summary"]["delivery_suitability_complaints_pct"]
    print(f"Delivery/suitability complaint share: {complaint_share:.1f}%")
    print(f"Weighted overall gross margin: {data['financial_summary']['weighted_margin_pct']:.2f}%")


if __name__ == "__main__":
    main()
