# Technical Documentation: AutoCAD Automation Engine

## 1. Project Goal & Objectives
The goal of this application is to automate the geographic mapping, asset placement, and network routing of utility infrastructure inside AutoCAD. By parsing structured coordinate and hardware data from Excel spreadsheets, the engine dynamically generates optimized AutoCAD Script (`.scr`) files. This eliminates manual drafting overhead, enforces proper layering conventions, prevents overlapping assets, and ensures exact spatial coordination.

---

## 2. Tech Stack
* **Core Engine:** Python 3.x
* **Data Parsing:** `openpyxl` (Excel read/write operations)
* **Geospatial Projection:** `pyproj` (Transforming WGS84 Lat/Lon coordinates into UTM Zone 47 South plane coordinates)
* **Hierarchical Parsing:** Python `re` (Regular expressions for parent-child relationship tracking and auto-routing logic)
* **Target Environment:** AutoCAD Runtime Environment (via Master and Area `.scr` execution profiles)

---

## 3. Implemented & Planned Features

### Automatic Coordinate & Routing Transformation
* Converts raw latitude/longitude metrics into localized Cartesian mapping metrics.
* Parses naming patterns to automatically discover parent poles and draw lines (`LINE`) interconnecting the grid seamlessly.

### Optimized Layering & Block Conventions
Assets are automatically grouped into specialized, separate layers to simplify drawing management:
* **Main Poles:** Instantiated on Layer `0` (Blocks: `TIANG_MERAH`, `TIANG_BIRU`, `TIANG_MAG`, `TIANG_SP`).
* **Auxiliary Gear:** Stays organized across specific infrastructure layers:
  * Umbang: Layer `UMBANG` (Blocks: `U1`–`U9`)
  * Services: Layer `JUMLAH SERVIS` (Blocks: `S1`, `S2`, etc.)
  * Blackbox 160A: Layer `BBOX 160A` (Block: `BK`)
  * Blackbox 400A: Layer `BBOX 400A` (Block: `BM`)
* **Cables:** Dynamically sorted into distinct layers based on dominant cross-section thickness (e.g., `CABLE 185`, `CABLE 95`, `CABLE 316`).

### Precise Micro-Layout Spacing (Offset Logic)
To completely prevent text and hardware blocks from overlapping, all elements use strict mathematical positioning relative to the main pole center point `(x, y)`:
* **Main Pole Label:** `X + 3.0`, `Y + 4.4` (Elevated profile)
* **Old Pole / N/A Label:** `X + 3.0`, `Y + 1.5` (Stacks cleanly directly underneath the main label)
* **Service Block:** `X - 2.8`, `Y - 3.2` (Bottom-left sector)
* **Blackbox Stacking Engine:** * Starts at `X + 3.5`, `Y - 3.2` (Bottom-right sector).
  * Cascades vertically downwards by a spacing index of `3.2` units per item if multiple units exist.
  * **Smart Shift:** If a pole contains both `BK` and `BM` units, the `BM` column automatically offsets horizontally by `5.2` units to prevent collision. If `BK` is absent, `BM` auto-occupies the primary column slot.

### Optimized Draw Order Configuration
* **Render Sequence:** Script executes `LINE` routing -> `SERVICE` blocks -> `MAIN POLE` blocks -> `TEXT` and `BLACKBOXES`.
* **Impact:** By calling the Service block *before* the main pole block, the pole symbol naturally anchors on top, maintaining a professional blueprint presentation.

---

## 4. Data Structure (Excel Schema)

The input spreadsheet (`data_input.xlsx`) follows a specific transactional column layout read sequentially by the Python processor:

| Column index | Data Property | Expected Values / Types | Description |
| :--- | :--- | :--- | :--- |
| **Col 1** | Geographic Data | `Latitude, Longitude` (String) | Raw global coordinates |
| **Col 2** | Main Pole ID | String (e.g., `PE A 9 /10`) | Node signature used for routing hierarchy |
| **Col 3** | Legacy Node | String / `N/A` | Old pole name tracking |
| **Col 4** | Umbang Specs | Integer (`1`–`9`) | Selects matching `U` block index |
| **Col 5** | Blackbox 160A | Integer Quantity (>= 0) | Cascading count for `BK` blocks |
| **Col 6** | Blackbox 400A | Integer Quantity (>= 0) | Cascading count for `BM` blocks |
| **Col 7** | Service ID | Integer / String | Determines matching service type block |
| **Col 8–11** | Asset Classification | Binary flags (`1` or empty) | Dictates block selection (`MERAH`, `BIRU`, etc.) |
| **Col 12–21** | Cable Specifications | Flag markers | Scan loop identifies the dominant cable network type |

---

## 5. Current Project Status & Next Steps

### Current Milestones Achieved
* Successfully resolved the global script orbiting issue by refactoring item transformation into specialized independent attribute vectors.
* Corrected the overlapping hardware bug by migrating structural detail arrays from generic text strings into dedicated CAD block components (`BK`, `BM`, `S1`).
* Deployed the localized layer update, safely splitting infrastructure tracking into separate layer filters (`BBOX 160A`, `BBOX 400A`, `JUMLAH SERVIS`).
* Re-ordered execution flow within the script constructor to keep asset rendering sequences visually perfect.

### Next Steps & Action Items
1. **Antigravity IDE Migration:** Load the consolidated source folder structure into the Google Antigravity workspace environment to initialize autonomous agent-driven development.
2. **Define System Prompt Rules:** Input our custom mathematical offsets and spreadsheet structures as locked-down developer parameters inside the agent config layout.
3. **Modular Extension:** Expand the dataset processing block to read and map secondary industrial assets (like transformers or streetlamps) using the same dynamic offset methodology.