# ⚡ Enterprise Multi-Source REST API Data Pipeline & Scheduled Sync Engine

![Platform](https://img.shields.io/badge/Platform-Google%20Workspace%20%7C%20Sheets-10B981?style=for-the-badge&logo=googlesheets&logoColor=white)
![Runtime](https://img.shields.io/badge/Runtime-Google%20Apps%20Script%20(V8)-3B82F6?style=for-the-badge&logo=google&logoColor=white)
![Architecture](https://img.shields.io/badge/Architecture-Zero%20Middleware%20REST%20ETL-0F172A?style=for-the-badge)

An enterprise-grade, self-hosted **ETL (Extract, Transform, Load) REST API Data Pipeline** built natively inside Google Sheets and Google Apps Script (V8). Engineered to replace costly recurring middleware subscriptions (**Zapier, Make.com, Coefficient, Supermetrics**) with a zero-maintenance, fault-tolerant ingestion engine.

---

## 🏛️ System Architecture & Data Flow

```text
+-----------------------------------------------------------------------------------+
|                        EXTERNAL REST API PROVIDERS (JSON)                         |
|   [CoinGecko Market API]     [Open Exchange Rates v6]     [DummyJSON Catalog API] |
+--------------+--------------------------+---------------------------+-------------+
               |                          |                           |
               v                          v                           v
+-----------------------------------------------------------------------------------+
|                  GOOGLE APPS SCRIPT V8 ENGINE (UrlFetchApp)                       |
|                                                                                   |
|  1. Scheduler Governance  <--- Reads state & intervals from [⚡ Sync_Controls]    |
|  2. Endpoint Registry     <--- Reads active URLs/Auth from  [⚙️ API_Config]       |
|  3. Resilient HTTP Client ---> Exponential Backoff (2^n * 400ms) on HTTP 429/5xx  |
|  4. Schema Normalizer     ---> Flattens heterogeneous JSON into unified 10-col UI |
|  5. Batch Memory Writer   ---> Single-call Range.setValues() (Zero-Loop I/O)      |
+-------------------------------+-----------------------------------+---------------+
                                |                                   |
                                v                                   v
         +---------------------------------------------+   +------------------------+
         |           📊 Market_Pipeline                |   |    📜 System_Logs      |
         | • Live Telemetry Table & Status Chips       |   | • Execution IDs        |
         | • Real-time KPI Cards (Latency, Health)     |   | • HTTP Codes (200/429) |
         | • Overwrite Snapshot or Append Time-Series  |   | • Latency (ms) & KB    |
         +---------------------------------------------+   +------------------------+
