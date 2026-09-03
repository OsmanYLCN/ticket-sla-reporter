# Prometeon IT Infrastructure - Ticket SLA Reporting & Quarterly Consolidation Engine

Automated data pipeline and analytics engine designed for Prometeon IT Infrastructure management. The system consolidates quarterly incident records (`Tickets-Q1` through `Tickets-Q4`), applies multi-layered scope filtering and IT service classification rules, and generates an executive `Summary` dashboard driven entirely by dynamic Excel formulas.

---

## Overview

Corporate IT incident management requires precise tracking of Service Level Agreements (SLAs) for critical infrastructure. This utility provides an automated, auditable, and non-destructive workflow to process raw ticket exports:

- Ingests quarterly data sources in read-only mode to guarantee raw data integrity.
- Consolidates quarterly records into a unified `Year-2026` master repository.
- Performs domain-specific scope classification across Network, Industrial Operations Technology (OT), and Enterprise Application boundaries.
- Produces an executive `Summary` dashboard with live Excel formulas (`COUNTIF`, `COUNTIFS`, `AVERAGEIFS`, `IF`), ensuring total auditability without hardcoded metrics.
- Prevents file corruption and operating system file-lock conflicts via proactive process monitoring.

---

## Governance & Scope Rules

In accordance with Prometeon IT Infrastructure SLA policy, tickets are evaluated against the following criteria:

### 1. Mandatory Baseline Criteria
- **Business Line:** Must strictly equal `Infrastructure Services`. Tickets originating from Workplace Services, Elmec Connect, or unrelated business units are excluded.
- **Severity Level:** Restricted to Severity `1 - Molto alta` (Critical) and `2 - Alta` (High). Severity 3 (Medium) and 4 (Low) are excluded.

### 2. Domain Classification Rules
- **Managed Network Services:** Tickets are flagged as In-Scope if text fields (`Oggetto`, `Servizio`, `Tipologia Calcolata`, `Configuration Item`, `Descrizione Articolo`) match recognized network service identifiers, equipment tokens, or network protocol terminology (e.g., Cato, routing, switching, firewalls, VPN, Wi-Fi, LAN, WAN, proxy, Cisco ISE, Zscaler/ZPA). Exact word boundaries are enforced to eliminate false positives.
- **Industrial Infrastructure (Manufacturing Plants / OT):** Centralized corporate headquarters (`Cesano Maderno`, `Milano`, `Cinisello`, `TIM Data Center`) and cloud tenants (`Microsoft Azure`, `Google Cloud Platform`) are strictly excluded from server-level scope. Only on-premises virtualization, physical compute, and backup systems located within manufacturing plants (`Kocaeli`, `Alexandria`, `Santo Andre`, `Gravatai`) are designated as In-Scope Industrial OT.
- **Enterprise Application Exception (SAP):** Standard application-layer SAP transaction failures, print spooler alerts, and background job queues are excluded from infrastructure scope. However, SAP accessibility incidents caused by network transport or Cato SD-WAN tunnel degradation are classified under Managed Network Services.

### 3. SLA Performance Metric
- **Resolution Target:** Tickets must reach resolved status (`Data Soluzione`) within 4.0 hours from initial creation (`Data Creazione`).
- **Calculation:** `Resolution_Hours = (Data Soluzione - Data Creazione) * 24`
- **SLA Target Compliance:** Minimum 90.0% of In-Scope critical incidents resolved within 4.0 hours.

### 4. Generated Workbook Structure
- **`Summary` Dashboard Sheet:**
  1. **Quarterly SLA Performance Matrix (Q1 - Q4):** Side-by-side comparison matrix displaying ticket volume, resolution within 4.0h, SLA breaches, and compliance results across each quarter.
  2. **Full Year 2026 Consolidated Table:** Executive summary consolidating annual infrastructure SLA performance using live Excel formulas.
  3. **Granular Breakdown Tables:** Domain breakdown (Managed Network vs. Industrial OT), severity distribution (Severity 1 vs. 2), and manufacturing plant breakdown (Kocaeli, Alexandria, Santo André, Gravataí, Other Sites).
- **`Year-2026` Master Dataset Sheet:**
  - Consolidates all quarterly records into a single auditable sheet with injected formula columns: `In_Scope` (Col AH), `Resolution_Hours` (Col AI), `SLA_Status` (Col AJ), `Scope_Category` (Col AK), `Plant_Site` (Col AL), and `Quarter` (Col AM).

---

## Repository Structure

```text
ticket-sla-reporter/
|-- sample_data/            # Staging directory for source raw Excel exports (read-only)
|   +-- .gitkeep
|-- outputs/                # Target directory for generated executive reports
|   +-- .gitkeep
|-- src/                    # Core processing modules
|   |-- __init__.py
|   |-- config.py           # Classification regex patterns, site mappings, and styles
|   |-- classifier.py       # Deterministic ticket categorization and boundary logic
|   |-- processor.py        # Streamlined workbook ingestion and datetime normalization
|   +-- excel_builder.py    # Master workbook consolidation, layout, and formula injection
|-- generate_sla_summary.py # Primary command-line interface with interactive menu
|-- setup.bat               # Windows one-click automated setup and dependency installer
|-- run.bat                 # Windows one-click launcher with drag-and-drop support
|-- requirements.txt        # Python dependency manifest
|-- .gitignore              # Git exclusion rules for Python, virtual environments, and data
+-- README.md               # Technical documentation
```

---

## Quick Start for Windows (Recommended)

1. Double-click **`setup.bat`**: Automatically checks for Python, provisions an isolated virtual environment (`.venv`), and installs required dependencies.
2. Place raw quarterly Excel files into `sample_data/`.
3. Double-click **`run.bat`** (or drag and drop any Excel file directly onto `run.bat`).

---

## Manual Prerequisites & Installation

### Environment Requirements
- Python 3.9 or higher
- Microsoft Windows, macOS, or Linux

### Setup
1. Clone or download the repository:
   ```bash
   git clone <repository_url>
   cd ticket-sla-reporter
   ```
2. Create and activate an isolated virtual environment:
   ```bash
   python -m venv .venv
   # Windows (PowerShell)
   .venv\Scripts\Activate.ps1
   # Windows (Command Prompt)
   .venv\Scripts\activate.bat
   # Linux / macOS
   source .venv/bin/activate
   ```
3. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## Operational Execution

Place raw export workbooks into the `sample_data/` directory.

### Method 1: Interactive Console (Recommended)
Run the script without positional arguments to launch the interactive selector:
```bash
python generate_sla_summary.py
```
The interface scans `sample_data/`, displays available `.xlsx` files ordered by modification timestamp (newest first), and prompts for selection:
```text
===========================================================================
  PROMETEON IT INFRASTRUCTURE - SLA SUMMARY & QUARTERLY REPORT GENERATOR
===========================================================================

[?] Available Excel Files in 'sample_data/' (Ordered by Date):
    (1) tickets_export_20260831_084937.xlsx           (Modified: 2026-09-01 14:15:04) [LATEST]
    (M) Enter custom file path manually
    (Q) Quit

Select file to process [1-1, M, Q or Enter for default (1)]:
```

### Method 2: Explicit File Argument
Specify input and output destinations directly via CLI arguments:
```bash
python generate_sla_summary.py sample_data/tickets_export_20260831_084937.xlsx
```
Or use explicit options:
```bash
python generate_sla_summary.py -i sample_data/tickets_export_20260831_084937.xlsx -o outputs/Q2_SLA_Report.xlsx
```

### Method 3: Non-Interactive / Batch Execution
For automated scheduling and task orchestration, pass `--non-interactive`:
```bash
python generate_sla_summary.py --non-interactive
```

---

## File Safety & Fault Tolerance

- **Raw Data Preservation:** The source workbook is opened exclusively in read-only mode (`read_only=True`). Input files remain unmodified.
- **Concurrent File Lock Protection:** If the target output workbook is open in Microsoft Excel when the generation process runs, the system intercepts the operating system lock (`PermissionError`) and prompts the user to close the document before retrying, preventing execution aborts.
- **Graceful Termination:** Operations can be safely cancelled at any prompt using `Q` or `Ctrl+C` without terminal tracebacks.
