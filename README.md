# 🏦 Automated Cash & E-Wallet Financial Operations System

[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/excel)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=flat&logo=python&logoColor=white)](https://www.python.org/)
[![VBA](https://img.shields.io/badge/VBA-Automation-darkblue?style=flat)](https://learn.microsoft.com/office/vba/)
[![Status](https://img.shields.io/badge/Status-Production%20Ready-success)](#)

An enterprise-grade financial management and operational analytics system designed to track, reconcile, and audit high-volume cash, multi-provider e-wallet, and banking transactions. 

Engineered end-to-end using **Python (`openpyxl`, `pywin32`)** to programmatically generate formulas, dimensional tables, VBA automation scripts, and native executive dashboards.

---

## 📌 Business Problem & Motivation

Retail and cash-transfer businesses operating in Egypt handle hundreds of daily transactions across cash drawers, banking accounts (InstaPay/CIB/NBE/Banque Misr), and multiple e-wallet lines (Vodafone Cash, Orange Cash, Etisalat Cash, WE Pay). 

**Key Operational Challenges:**
* **Liquidity Fragmentation:** Capital split across 18+ telecom lines, bank accounts, and physical cash.
* **Hidden Fee Leakage:** Discrepancies between gross customer fees and network provider deductions.
* **Audit & Reversal Bottlenecks:** Human errors in recording transactions distorting historical ledgers.
* **Cash Drawer Variance:** End-of-day discrepancies between counted cash and accounting balances.

---

## 🏗️ System Architecture & Financial Logic

The system follows double-entry bookkeeping principles adapted for retail liquidity management:
