# Azure Subscription Transfer – Impacted Resources Assessor
A PowerShell tool that discovers Azure resources that are commonly **impacted by subscription transfer**, and exports a **single Excel workbook** with four tabs:

- **Summary** – counts by resource type (KV, AKS, UAIs, SQL DBs, ADLS Gen2, etc.)
- **Details** – one row per impacted resource, with useful flags (MSI, HNS, encryption hints)
- **RBAC** – role assignments at **subscription**, **resource group**, and **resource** scope (for re‑creation planning)
- **Networking** – **Private Endpoints**, associated **Private DNS Zone Groups**, and **Private DNS VNet Links**

> Designed for tenant‑wide discovery, management group scope, or specific subscriptions. Includes Azure Resource Graph (ARG) pagination, helpful defaults, and optional CSV exports.

---

## ✨ Features

- 🔎 Scans **impacted resource types** and any resource with **SystemAssigned** Managed Identity.
- 📦 Includes **ADLS Gen2** Storage (HNS‑enabled).
- 🔐 Exports **RBAC** assignments to help recreate permissions after transfer.
- 🌐 Maps **Private Endpoints → targets**, associated **Private DNS zones**, and **VNet links**.
- 📊 Produces a single **Excel workbook** (auto‑sized, filtered tables; UK‑friendly).
- ⚙️ **No KQL `let`** statements – avoids Azure Resource Graph parser quirks.
- 🚦 Lightweight GitHub **CI** with PSScriptAnalyzer lint.

---

## 🛠 Prerequisites

- PowerShell 7.x recommended (Windows PowerShell 5.1 also supported).
- Modules:
  - `Az` (Az.Accounts, Az.ResourceGraph, Az.Resources)
  - `ImportExcel` (auto‑installs if missing)

Install if needed:
```powershell
Install-Module Az -Scope CurrentUser
Install-Module Az.ResourceGraph -Scope CurrentUser
# ImportExcel is auto-installed by the script if missing
