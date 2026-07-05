# FIFO Inventory Analyzer

A desktop ERP tool for calculating inventory valuation, cost of goods sold, and profit/loss using the **FIFO (First-In-First-Out)** method. Built for wholesale/retail businesses that track purchases, sales, and damages in Excel.

## Features

- Import Excel transaction ledgers (`.xlsx`)
- Automatic FIFO cost calculation:
  - **FIFO1** — Current average inventory cost
  - **FIFO2** — Weighted average cost of units sold
- Real-time profit/loss per transaction
- Multi-sheet Excel report export:
  - FIFO Details (transaction-level)
  - FIFO Summary (product-level)
  - Customer Profit Report
  - Invoice Profit Report
  - Monthly Purchase/Sale Breakdown
  - Current Stock / Batch Report
- Arabic/English bilingual UI
- Password-protected login
- Threaded background processing (no UI freezing)
- Handles negative stock, transfers, and damaged goods

## Tech Stack

- Python 3
- Tkinter (GUI)
- Pandas & NumPy
- openpyxl (Excel read/write)
- tkcalendar (optional date picker)

## Prerequisites

Make sure you have **Python 3.8+** installed.

## Installation & Build

1. **Clone or download** this repository.
2. **Install dependencies** via pip:

## How to Use

Click "تحميل ملف Excel" and select your transaction ledger.
Pick the analysis date using the date picker.
Click "تحليل البيانات" to process.
Export reports using the buttons:
تقرير الربح و تكلفة المخزن حسب المادة
تقرير الربح حسب العميل
تقرير الربح حسب الفاتورة
تقرير المخزون المتبقي
Expected Excel Format
Your input sheet should contain Arabic column headers including:
تاربخ (Date)
نوع (Transaction Type)
كود المادة (Product Code)
اسم المادة (Product Name)
سعر البيع (Unit Price)
كمية الوارد (Incoming Qty)
كمية الصادر (Outgoing Qty)
كود المشتري / اسم المشتري (Customer info)



```bash
pip install pandas numpy openpyxl tkcalendar
python fifo_calculator.py

