# Collection Report Module — Odoo 18

## Overview
This module adds a **Collection Report** wizard under the Sales menu that generates a downloadable Excel (.xlsx) report showing sales invoices aged **14, 21, or 28 days**, grouped by city.

## Features
- 🧙 Wizard with date selector (14 / 21 / 28 days or custom range)
- 🏙️ Optional city filter
- 📊 Excel output with **3 sheets**:
  1. **Summary by City** — pivot-style table per city with aging amounts
  2. **Invoice Details** — full row-level data with auto-filter
  3. **Aging Analysis** — bucket breakdown with totals
- 🎨 Color-coded aging buckets (Green → Yellow → Red)

## Installation
1. Copy `collection_report/` folder to your Odoo addons path
2. Update Apps list in Odoo
3. Install **Collection Report (14/21/28 Days)**

## Usage
`Sales → Collection Reports → Collection Report (14/21/28 Days)`

## Requirements
- Odoo 18
- Python package: `openpyxl` (usually pre-installed with Odoo)
- Depends on: `sale_management`, `account`
