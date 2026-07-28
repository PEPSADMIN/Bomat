# PEPS BOM Automation Tool — v2.1

A web-based Bill of Materials (BOM) generation and management system for PEPS India.  
Automates BOM creation across **250+ Excel macro files**, tracks every change, and provides a full audit trail with rollback capability.

---

## Table of Contents

1. [Overview](#overview)
2. [Tech Stack](#tech-stack)
3. [Getting Started](#getting-started)
4. [Features](#features)
   - [Login](#1-login)
   - [Product Browser](#2-product-browser)
   - [BOM Generator](#3-bom-generator)
   - [Formula Calculator](#4-formula-calculator)
   - [Global Replace](#5-global-replace)
   - [History & Rollback](#6-history--rollback)
   - [Nearest BOM](#7-nearest-bom)
   - [Formula Guide](#8-formula-guide)
   - [Approvals](#9-approvals)
   - [Admin Panel](#10-admin-panel)
5. [Server & Auto-Start](#server--auto-start)
6. [File Structure](#file-structure)

---

## Overview

The BOM Tool eliminates manual Excel work by letting users select a product, enter dimensions, and instantly generate a complete, formatted BOM Excel file. All changes are logged, approval workflows are built in, and every item-code replacement across the entire file library can be tracked and rolled back.

**Access:** `https://192.168.0.133:5020` (LAN only)

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3.11 · Flask · Cheroot WSGI |
| Database | SQLite (`bom_tool.db`) |
| BOM Engine | openpyxl · xlwings (reads `.xlsm` macros) |
| Frontend | Vanilla JS SPA (single HTML page) · Bootstrap 5 |
| Auth | Session-based login with role/tab permissions |
| Server | HTTPS (self-signed) · Windows Task Scheduler watchdog |

---

## Getting Started

### Prerequisites
- Python 3.11+
- Windows 10/11
- SSL certificate files in `ssl/`

### First-time setup
```bash
cd "D:\Hari JR. DATA\Development\Bom Tool"
python -m venv .venv
.venv\Scripts\pip install -r requirements_v2_1.txt
```

### Start the server
The server runs automatically via Windows Task Scheduler (`PEPS BOM Tool` task).  
To start manually:
```bat
run_forever.bat
```

The watchdog (`PEPS BOM Watchdog` task) checks health every 5 minutes and auto-restarts if the server goes down.

---

## Features

### 1. Login

Secure session-based login. Each user has role-based access — only permitted tabs are visible after login.

![Login Page](docs/screenshots/01_login.png)

---

### 2. Product Browser

Browse and search the full product catalogue. Filter by name, item code, section, or department. Select a product to open its BOM panel.

![Product Browser](docs/screenshots/02_product_browser.png)

**Key actions:**
- Search by product name or item code
- Filter by Section / Department
- View product details and size varieties
- Open BOM Generator directly from any product card

---

### 3. BOM Generator

The core feature. Select a product, enter the mattress dimensions (Length × Width), and generate a complete BOM.

![BOM Generator — Single Product](docs/screenshots/03_bom_generator.png)

**Options:**
| Option | Description |
|---|---|
| Download Single | Generate BOM for one specific size |
| Download All Sizes | Generate BOM for every size variant at once |
| Run (Background) | Queue a BOM run and download when ready |

![BOM Generator — All Sizes](docs/screenshots/04_bom_all_sizes.png)

---

### 4. Formula Calculator

Inspect and edit the formula behind any BOM line item. The calculator parses the formula and shows a visual breakdown so you can verify the logic before saving.

![Formula Calculator](docs/screenshots/05_formula_calculator.png)

**Supported formula patterns:**

| Pattern | Example | Description |
|---|---|---|
| Perimeter | `2*(L+W)*k` | Tape, border, quilting thread |
| Perimeter + offset | `(L+W+offset)*k` | Border fabric with seam allowance |
| Area | `L*W*k` | Glue, foam panels, fabric |
| Volume | `L*W*H*k` | Foam blocks, spring units |

The **LAYER** box appears automatically for area/volume formulas — set multiple layers and the formula updates to `L*W*k*n`.

![Formula Calculator — Layer Box](docs/screenshots/06_formula_layer.png)

**Foam/Spring Code Panel:**  
When the item code or description contains a foam or spring code, a dedicated panel appears for editing dimensions without breaking the code format.

![Foam Code Panel](docs/screenshots/07_foam_panel.png)

---

### 5. Global Replace

Replace an item code across **all 250+ BOM Excel files** in one operation. Every replace is snapshotted so it can be rolled back.

![Global Replace — Preview](docs/screenshots/08_global_replace_preview.png)

**Workflow:**
1. Enter the old item code and new item code
2. Add reason and approval reference
3. Click **Preview** — see every file and row that will change
4. Download preview as Excel for sign-off
5. Click **Execute** — all files updated atomically

![Global Replace — Execute](docs/screenshots/09_global_replace_execute.png)

**Replace History:**  
Every executed replace is stored with `who`, `when`, `reason`, `approval ref`, and full before/after data.

![Global Replace — History](docs/screenshots/10_global_replace_history.png)

---

### 6. History & Rollback

Every BOM run is logged. If an error is found after generation, any run can be individually rolled back to restore the original file state.

![History](docs/screenshots/11_history.png)

**Columns logged per run:**

| Field | Value |
|---|---|
| Product | Name + item code |
| Dimensions | L × W |
| Run by | User name |
| Timestamp | Date & time |
| Status | Success / Failed |
| Action | Download · Rollback |

---

### 7. Nearest BOM

Find the closest existing product by specification and use it as a starting point. Useful for quoting new products that don't have a BOM yet.

![Nearest BOM — Single](docs/screenshots/12_nearest_bom.png)

**Bulk mode:** Upload a CSV/Excel with multiple specs and download all nearest BOMs in one batch.

![Nearest BOM — Bulk](docs/screenshots/13_nearest_bom_bulk.png)

---

### 8. Formula Guide

Reference library of all formula patterns with examples. Helps BOM administrators write and validate formulas consistently.

![Formula Guide](docs/screenshots/14_formula_guide.png)

---

### 9. Approvals

Changes that require sign-off (e.g., new item codes, formula edits) go into the Approvals queue. Approvers receive an in-app notification badge.

![Approvals Queue](docs/screenshots/15_approvals.png)

**Workflow:**
- Submitter raises a change → status: `Pending`
- Approver reviews and approves / rejects → submitter notified
- Admin can override at any stage

---

### 10. Admin Panel

Manage users, roles, tab permissions, and view full activity logs.

![Admin — Users](docs/screenshots/16_admin_users.png)

**User management:**
- Create / disable user accounts
- Assign roles (Admin · BOM Admin · Viewer · etc.)
- Control which tabs each user can access
- Force password reset

![Admin — Roles](docs/screenshots/17_admin_roles.png)

**Activity Log:** Full audit trail of every action taken by every user.

![Admin — Activity Log](docs/screenshots/18_admin_activity.png)

---

## Server & Auto-Start

The server is managed by two Windows Scheduled Tasks:

| Task | Trigger | Runs as | Purpose |
|---|---|---|---|
| `PEPS BOM Tool` | At system startup | Hari | Starts `run_forever.bat` loop |
| `PEPS BOM Watchdog` | Every 5 minutes | SYSTEM | HTTP health check; kills zombie and restarts if down |

**Watchdog logic (`watchdog_bom.ps1`):**
1. Calls `GET /api/health` — expects `{"status":"ok"}`
2. If no response → kills any process holding port 5020 (including zombies with blank CommandLine)
3. Starts `run_forever.bat` if it is not already running
4. Verifies server came back up; logs result to `bom_server.log`

**`run_forever.bat`** also kills any zombie on port 5020 at startup before launching Python, so manual restarts never get stuck.

---

## File Structure

```
Bom Tool/
├── app_v2_1.py                    # Flask app — all API routes
├── database_v2_1.py               # DB schema, queries, history logging
├── bom_engine_v2_0.py             # BOM generation engine (openpyxl)
├── bom_scanner_v2_0.py            # Scans xlsm files and builds product index
├── product_structure_generator_v2_1.py  # New product BOM builder
├── config_v2_0.py                 # Paths, port, SSL config
├── templates/
│   └── products_v2_1.html         # Full SPA frontend (single file)
├── ssl/                           # HTTPS certificate (not in git)
├── spec_cache/                    # Runtime formula cache (not in git)
├── bom_tool.db                    # SQLite database (not in git)
├── run_forever.bat                # Server start + watchdog loop
├── watchdog_bom.ps1               # Health-check watchdog script
├── run_watchdog.bat               # One-shot watchdog launcher
├── docs/
│   └── screenshots/               # UI screenshots (add here)
├── PEPS BOM Tool — User Manual.html
└── PEPS_BOM_Tool_User_Guide.pdf
```

---

## Adding Screenshots

Screenshots are referenced from `docs/screenshots/`. To add them:

1. Open `https://192.168.0.133:5020` in your browser
2. Take a screenshot of each section (use `Win + Shift + S` or Snipping Tool)
3. Save as the filename listed in the table below
4. Place in `docs/screenshots/`

| File | Page / Panel |
|---|---|
| `01_login.png` | Login page |
| `02_product_browser.png` | Product list / search |
| `03_bom_generator.png` | BOM panel — single product |
| `04_bom_all_sizes.png` | BOM — all sizes download |
| `05_formula_calculator.png` | Formula Calculator panel |
| `06_formula_layer.png` | Formula Calculator — LAYER box |
| `07_foam_panel.png` | Foam/Spring Code panel |
| `08_global_replace_preview.png` | Global Replace — preview results |
| `09_global_replace_execute.png` | Global Replace — execute confirm |
| `10_global_replace_history.png` | Global Replace — snapshot history |
| `11_history.png` | BOM run history |
| `12_nearest_bom.png` | Nearest BOM — single lookup |
| `13_nearest_bom_bulk.png` | Nearest BOM — bulk upload |
| `14_formula_guide.png` | Formula Guide |
| `15_approvals.png` | Approvals queue |
| `16_admin_users.png` | Admin — user management |
| `17_admin_roles.png` | Admin — roles |
| `18_admin_activity.png` | Admin — activity log |

---

*PEPS India · BOM Automation Tool v2.1 · Internal use only*
