# Group 41 NexaMart

The workbook and Power BI report support the Kampus Digital business performance challenge.

## Recommendation figures

Three charts accompany the marketing, customer experience and product profitability recommendations. These charts contain historical results. A fourth diagram summarises the learning and reflection points. Its final step describes planning a controlled trial. Proposed trial targets remain in the presentation text.

| Recommendation | PNG for slides | SVG with editable text |
| --- | --- | --- |
| 1. Marketing budget reallocation | [PNG](figures/recommendations/01_marketing_roas.png) | [SVG](figures/recommendations/01_marketing_roas.svg) |
| 2. Delivery and product suitability | [PNG](figures/recommendations/02_customer_complaints.png) | [SVG](figures/recommendations/02_customer_complaints.svg) |
| 3. Home Office profitability | [PNG](figures/recommendations/03_category_gross_margin.png) | [SVG](figures/recommendations/03_category_gross_margin.svg) |
| Learning and reflection | [PNG](figures/recommendations/04_learning_reflection.png) | [SVG](figures/recommendations/04_learning_reflection.svg) |

The PNGs use the slide's cream background (`#F4EFE3`) and export at 2700 x 1950 pixels. The SVG backgrounds are transparent. Marketing ROAS uses a horizontal bar chart, complaint shares use a donut chart, and category gross margins use a dot plot. The bar chart and dot plot label both axes. The donut labels each category and its percentage of complaints. Explanatory captions and callouts stay in the slide text. Orange highlights the main comparison and amber identifies suitability complaints or the suggested add-on categories. Use a chart beside two or three recommendation points.

![Marketing ROAS](figures/recommendations/01_marketing_roas.png)

![Customer complaints](figures/recommendations/02_customer_complaints.png)

![Category gross margin](figures/recommendations/03_category_gross_margin.png)

![Learning and reflection](figures/recommendations/04_learning_reflection.png)

The reflection diagram is square at 2250 x 2250 pixels. It uses three numbered steps and two connecting arrows, with one short caption per step.

## Rebuild

Use Python 3.10 or later. From the repository root:

```bash
python -m pip install -r scripts/requirements-figures.txt
python scripts/build_recommendation_figures.py
python scripts/build_reflection_diagram.py
```

An updated workbook can be passed explicitly:

```bash
python scripts/build_recommendation_figures.py --workbook "updated.xlsx" --output-dir figures/recommendations
```

The script reads the workbook without modifying it. It exports three PNGs, three SVGs, three CSV tables and `metrics.json`. Calculations use the raw sheets, so stale cached summary formulas do not determine the figures. SVG text remains editable. The script and pinned dependencies reproduce the charts without an image generation service.

The downloadable figures bundle does not contain another copy of the workbook. When using that bundle, pass the path to your existing Excel file with `--workbook`, as in the command above.

`build_reflection_diagram.py` is standalone and needs only Matplotlib. Change its `STEPS` tuple to update the labels or captions. It exports a cream-background PNG and a transparent SVG with editable text, using the same metadata settings as the charts.

## Updating the code

The script keeps source parsing, business calculations and chart styling in separate functions. `build_metrics` reads the source tables, then calls `calculate_marketing_metrics`, `calculate_service_metrics` and `calculate_financial_metrics`. The plotting functions use those results without repeating the calculations. Colours and the recommendation categories are named constants near the top of the file.

`save_chart` controls both image formats. PNGs retain their resolution information but omit optional text, software, EXIF and XMP metadata. SVGs omit their metadata element while preserving editable text. These are charts plotted from the workbook with Matplotlib.

Run the calculation and export checks with:

```bash
python -m unittest discover -s tests -v
```

The tests cover weighted ratios, missing product costs, duplicate IDs, percentage parsing, service counts and metadata-free exports.

## Calculation definitions

- **Marketing ROAS** is total attributed campaign revenue divided by total campaign spend for each channel. It does not establish channel profit. Influencer and Social Media receive 46.1% of the recorded budget.
- **Complaints** include all 240 unique orders, with the exact duplicate order removed. The script normalises `late delivery` to `Late Delivery`. Of 58 complaints, 35 concern late delivery and 18 concern product suitability. These categories account for 53/58 complaints, or 91.4%, and contain all 21 returned orders. The complaint share does not predict how many complaints a corrective action will prevent.
- **Category gross margin** is summed gross profit divided by summed revenue, using completed orders with recorded product costs. The financial population contains 207 orders, RM57,609.65 revenue, RM34,407.00 product cost and RM23,202.65 gross profit. Its weighted margin is 40.28%. One completed P999 order with RM716.40 revenue has no product cost and is excluded from both the margin numerator and denominator. That order remains in service counts.
- **Home Office revenue share** uses the same cost-matched financial population as the margin comparison. Its RM24,654.00 revenue represents 42.8% of that population. Category margins exclude marketing, delivery and overhead costs.

Missing customer segments and satisfaction ratings do not remove orders from the complaint chart. Missing ratings are not assigned numerical scores. No region labels are recoded for these figures.

The CSV files contain the plotted values. `metrics.json` also records the source workbook's SHA-256 digest, duplicate count and financial exclusions. The original source workbook and PBIX report remain unchanged.
