# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PlanForge Pro is a Korean desktop application for shipment planning optimization. It's a single-file Python application (`PlanForge.py`) built with CustomTkinter that generates optimal daily shipment plans based on production schedules and inventory data.

## Core Dependencies

Install dependencies with:
```bash
pip install -r requirements.txt
```

Key libraries:
- `customtkinter` - Modern UI framework (main GUI)
- `pandas` - Data processing for production plans and inventory
- `matplotlib` - Charts and graphs for inventory visualization
- `openpyxl` - Excel file reading/writing
- `tkcalendar` - Date picker widgets

## Running the Application

```bash
python PlanForge.py
```

The application is a desktop GUI with three main workflow steps:
1. Load production plan (Excel import)
2. Inventory simulation (text input processing)
3. Export shipment plan (Excel export)

## Build and Release

Build executable for distribution:
```bash
pyinstaller --name "PlanForge" --onedir --windowed --icon="assets/logo.ico" --add-data "assets;assets" --hidden-import pygame --hidden-import Pillow --hidden-import keyboard PlanForge.py
```

Releases are automated via GitHub Actions when tags starting with 'v' are pushed (see `.github/workflows/release.yml`).

## Architecture

**Single-file application structure:**
- `ConfigManager` - JSON configuration management (`config.json`)
- `PlanProcessor` - Core business logic for shipment optimization
- `ProductionPlannerApp` - Main tkinter application window
- Multiple dialog classes for UI components (SearchableComboBox, InventoryInputDialog, etc.)

**Key data flows:**
1. Excel production plans → pandas DataFrame processing
2. Text inventory input → regex parsing and DataFrame conversion
3. Optimization algorithm → shipment plan generation
4. Results → Excel export with formatting

**Configuration system:**
- `config.json` - Runtime settings (pallet size, lead time, truck capacity, etc.)
- `assets/Item.csv` - Product master data with priorities
- Settings can be modified via UI and persist between sessions

**Business logic core:**
The `PlanProcessor` class handles the shipment optimization algorithm considering:
- Lead times and delivery schedules
- Truck capacity constraints (pallets per truck, max trucks per day)
- Priority-based allocation when multiple items compete for capacity
- Proactive shipment to maximize truck utilization

## File Structure

```
PlanForge.py          # Main application (2700+ lines)
config.json           # Runtime configuration
requirements.txt      # Python dependencies
assets/
  ├── Item.csv        # Product master data
  ├── logo.ico        # Application icon
  └── logo.png        # Logo image
.github/workflows/
  └── release.yml     # Automated build and release
```

## Configuration Management

The application uses JSON-based configuration (`config.json`) with automatic fallback to defaults. Key settings include:
- `PALLET_SIZE` - Items per pallet
- `LEAD_TIME_DAYS` - Shipment lead time
- `PALLETS_PER_TRUCK` - Truck capacity
- `MAX_TRUCKS_PER_DAY` - Daily truck limit
- `DELIVERY_DAYS` - Weekday delivery schedule
- UI preferences (font size, appearance mode, auto-save path)

Configuration is managed through the `ConfigManager` class which handles loading, saving, and providing defaults for missing values.

## Development Notes

- Korean language UI and business logic
- Windows-focused (uses Windows-specific paths in auto-update)
- Built for PyInstaller distribution
- Contains auto-update functionality via GitHub API
- Logging configured to DEBUG level during development
- No test framework - application is GUI-driven with manual testing