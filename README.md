# 📊 National Political Parties Social Media Audit Index

A high-efficiency, multi-platform data extraction pipeline engineered to monitor, aggregate, and analyze real-time digital campaign metrics for national political parties (INC, BJP, and AAP) across **Instagram, Facebook, and YouTube**. 

This repository houses the raw data streams, automated data transformation pipelines, and a production-ready executive tracking sheet generated via programmatic openpyxl manipulation.

---

## 🛠️ System Architecture & Engineering Paradigm

Unlike standard browser-automation frameworks (e.g., Selenium/Playwright) which suffer from high memory overhead and frequent anti-bot breakage, this project prioritizes **network-layer inspection and internal API data harvesting**. 

### Core Features:
*   **Performance Optimization:** Bypasses heavy DOM rendering by targeting server-side NextJS hydration data (`__NEXT_DATA__`) and internal GIS payloads directly.
*   **Multi-Platform Aggregation:** Consolidates profile metrics, active interaction counts ("Talking About" indices), and content output velocity.
*   **Automated Spreadsheet Styling:** Programmatically compiles data into an enterprise-ready dashboard with custom themes, explicit cell data type formatting, and dynamic column auto-fitting.

---

## 📁 Repository Structure

```text
├── Political_Parties_Master_Dashboard.xlsx  # Styled Corporate Output
├── generate_political_dashboard.py         # Openpyxl UI Transformation Engine
├── GROWTH_STRATEGY.md                           # Deep-dive Strategic Playbook (INC)
├── SUMMARY.md                                   # Cross-Party Analytical Executive Summary
└── README.md                                    # Project Documentation
