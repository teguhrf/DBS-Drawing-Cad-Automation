# AutoCAD Automated Pole Routing & Data Cleaning

This Python script automates the drafting process for electrical pole networks in AutoCAD. It takes raw survey data from an Excel file, cleans up inconsistencies, calculates parent-child pole relationships, and generates a `.scr` (AutoCAD Script) file to draw the entire network accurately without manual data entry.

## Why This Exists
Copy-pasting thousands of rows directly into the AutoCAD command line often causes missing lines or overlapping text due to OS buffer limits. This tool solves that by generating a native AutoCAD script file with controlled execution pacing. It also acts as a rigorous data cleaner, resolving duplicate cable entries and preventing overlapping lines in the final drawing.

## Core Features
* **Priority Cable Cleaning (In-Place Erase):** Scans multiple cable columns from largest to smallest capacity. It locks in the dominant cable, ignores the rest, and clears duplicate entries in the output Excel file to ensure a sterile dataset and a single routing line.
* **Dynamic Parent-Child Routing:** Understands complex naming structures and suffixes (e.g., pole `A 10 /5B` connects to `A 10 /4B`, and `A 10 /1B` connects to the main `A 10` pole). It dynamically searches for the correct parent pole or joint pole to execute the line routing.
* **Coordinate Conversion:** Automatically converts Latitude/Longitude coordinates to UTM Zone 47S for accurate spatial plotting.
* **Smart Text Parsing:** Preserves standard "PE" (Feeder Pillar) names from being truncated while cleanly stripping unnecessary area codes from standard pole labels.
* **AutoCAD Fallback Logic:** Defaults to a standard pole block (`TIANG_BIRU`) and default cable type (`Cable 1X 16 Nmp`) if surveyor data is missing or corrupted, preventing the script from crashing.

## Prerequisites
You need Python installed on your system along with these two libraries:
```bash
pip install openpyxl pyproj
