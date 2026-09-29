# Pulse-bi-solution
A solution that helps user create dashboards easily
# PulseBI - Executive Analytics & Multi-File Intelligence Suite

**PulseBI** is a lightweight, high-performance, client-side Business Intelligence (BI) web application designed for data professionals and portfolio showcases. It allows users to ingest multiple datasets simultaneously, inspect automated column schemas, and instantly generate interactive executive dashboards.

---

## 🚀 Key Features

* **Landing & Authentication Flow:** Clean onboarding experience with persistent local storage user sessions.
* **Multi-File Ingestion Engine:** Seamlessly upload and merge multiple **CSV** and **Excel (`.xlsx`, `.xls`)** files simultaneously using SheetJS.
* **Automated Data Type Inference:** Dynamically inspects column values and identifies data types (`integer`, `decimal`, `string`, `date`, `boolean`) before dashboard generation.
* **Interactive Visualization Studio:** Automatically aggregates records and renders responsive charts using Chart.js.
* **Unified Data Grid:** Inspect, search, and analyze records with clean schema type badges.

---

## 🛠️ Tech Stack

* **Frontend Framework:** HTML5, Tailwind CSS (Dark Mode UI)
* **Scripting:** Vanilla JavaScript (ES6+)
* **Data Parsing Libraries:** 
  * [SheetJS (xlsx)](https://sheetjs.com/) – For high-speed browser-based spreadsheet parsing.
  * [Chart.js](https://www.chart.js/) – For responsive data visualization.

---

## 🔍 How Data Type Inference Works

PulseBI inspects the first 10 rows of every column in your uploaded files to evaluate and tag data types dynamically:
1. **Integer:** Numbers without decimal points.
2. **Decimal:** Floating-point numbers.
3. **Boolean:** True/False values.
4. **Date:** Standard recognizable date-time strings.
5. **String:** Textual or unclassified categorical data.

---

## 📦 Local Setup & Running

1. Clone or download this repository.
2. Open the project folder in **VS Code**.
3. Install the **Live Server** extension.
4. Right-click `index.html` and select **"Open with Live Server"**.

---
&copy; 2026 pulse bi solution. product solution
