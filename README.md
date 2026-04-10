# KL Sales Report Generator

A desktop tool that reads raw order exports from TikTok or Shopee, merges them with a product cost/charges reference file, and produces a formatted Excel sales report with per-SKU breakdowns and a grand total.

---

## Features

- Supports **TikTok** and **Shopee** order export formats
- Automatically parses platform charge percentages (e.g. `18.88%`) from the product file and applies them against revenue
- Groups orders by SKU and unit price, with subtotals and a grand total
- Outputs a styled `.xlsx` report (borders, yellow highlight on totals, fixed column widths)
- Cross-platform GUI — works on **Windows** and **macOS**

---

## Requirements

- Python 3.8 or higher
- `tkinter` — included with standard Python installers
  - On macOS via Homebrew, install separately: `brew install python-tk`

---

## Installation

```bash
# 1. Clone the repository
git clone <your-repo-url>
cd <repo-folder>

# 2. (Optional) Create a virtual environment
python -m venv .venv
source .venv/bin/activate      # macOS / Linux
.venv\Scripts\activate         # Windows

# 3. Install dependencies
pip install -r requirements.txt
```

---

## Usage

```bash
python kl_report_v2.py
```

A window will open with the following fields:

| Field | Description |
|---|---|
| **Input File** | The raw order export from TikTok or Shopee (`.xlsx`) |
| **Product File** | The product code, cost, and platform charges reference file (`.xlsx`) |
| **Output File Name** | Name for the generated report (`.xlsx` is appended automatically if omitted) |
| **Platform** | Select **TikTok** or **Shopee** from the dropdown |

Click **Generate Report** to run. Progress and any errors are shown in the log area at the bottom.

---

## Expected File Formats

### Input File — TikTok

| Column | Description |
|---|---|
| `Seller SKU` | Product SKU identifier |
| `Quantity` | Units sold |
| `SKU Subtotal After Discount` | Revenue after discount |
| `SKU Platform Discount` | Discount amount (added back to compute adjusted subtotal) |

> The TikTok export has a blank second row — this is handled automatically.

### Input File — Shopee

| Column | Description |
|---|---|
| `SKU Reference No.` | Product SKU identifier |
| `Quantity` | Units sold |
| `Deal Price` | Per-unit selling price |

### Product File

| Column | Description |
|---|---|
| `Product Code` | SKU identifier (matched against the input file) |
| `Product Name & Color` | Display name used in the report |
| `Cost` | Per-unit cost |
| `Tiktok Platform Charges` | TikTok charge rate, e.g. `18.88%` |
| `Shopee Platform Charges` | Shopee charge rate, e.g. `12.00%` |

Platform charges are percentages of the total revenue for that SKU group.

---

## Output

The generated report contains one row per SKU/price-tier combination, followed by a **yellow subtotal row** per SKU and a **grand total row** at the bottom.

| Column | Description |
|---|---|
| Seller SKU | Product SKU |
| Product Name | From the product file |
| Per Unit Adjusted Subtotal | Effective selling price per unit |
| Total Quantity | Units sold in this group |
| Total Amount | Revenue (unit price × quantity) |
| Total Cost | Cost (unit cost × quantity) |
| Platform Charges | Platform fee (charge rate × revenue) |
| Profit | Total Amount − Total Cost − Platform Charges |
