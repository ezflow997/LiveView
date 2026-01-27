# LiveView Project

---
## MANDATORY: LiveView.ahk Editing Rule

**DO NOT EDIT `LiveView.ahk` WITHOUT EXPLICIT USER PERMISSION.**

Before making ANY changes to `LiveView.ahk`:
1. Create a COMPLETE edit plan first
2. Present the plan to the user
3. Ask: "Ready to apply these changes to LiveView.ahk?"
4. **WAIT for the user to say "yes" before using Edit/Write tools**

The file may be actively edited by another process. Editing without permission causes conflicts and data loss.

---

## Overview
LiveView.ahk is an AutoHotkey v2 application that creates a live thumbnail viewer with customizable regions, widgets, and backgrounds.

## Main File
- `LiveView.ahk` - Main application file containing the `ThumbnailViewer` class

## Key Features

### DWM Thumbnail Regions
- Multiple regions can capture different source windows
- Each region has independent source/destination rectangles
- Regions can be moved, resized, reordered (bring to front/send to back)
- Source area selection via drawing on the source window

### Widgets
- **Clock Widget**: Displays time and date with customizable font, color, format
- **Weather Widget**: Fetches weather from wttr.in API
- Widgets can be moved/resized with mouse or arrow keys
- Widget settings: font name, font color, background color, formats

### Background System
- Cycles through images from `backgrounds/` folder
- Supports .jpg, .jpeg, .png, .bmp files
- 30-second cycle interval (configurable via `bgCycleInterval`)
- Background stays behind all other controls via z-order management

### Fullscreen Modes
- **Edit Fullscreen (E key)**: Fullscreen but still editable, shows exit button
- **Locked Fullscreen (F11)**: True fullscreen, hides cursor, no editing
- Both modes enforce always-on-top via periodic `SetWindowPos` calls (every 100ms)

### Hotkeys (when window focused)
- `W` - Select source window
- `S` - Select source area
- `A` - Add new region
- `D` - Delete region
- `H` - Toggle source visibility (hide/show source windows)
- `E` - Edit fullscreen
- `F11` - Locked fullscreen
- `Escape` - Exit current mode
- `Arrow keys` - Move selected region/widget
- `Shift+Arrow keys` - Resize selected region/widget
- `PgUp/PgDn` - Change region z-order
- `Ctrl+S` - Save config
- `Ctrl+O` - Load config

## Folders
- `backgrounds/` - Background images for cycling display

## Configuration
- Configs saved as .ini files
- Stores region positions, sources, widget settings, weather location

## Technical Notes
- Uses DWM (Desktop Window Manager) thumbnail API for live window capture
- Picture control for background with z-order managed via SetWindowPos
- Timers: widget updates (1s), background cycling (1s check), thumbnail refresh (16ms), force redraw (100ms)

## Common Issues & Solutions
- **Background covers widgets**: Call `SendBackgroundToBack()` after adding controls
- **Background not fullscreen**: Call `UpdateBackgroundSize()` after window resize
- **File edit conflicts**: Use PowerShell scripts for multi-line replacements when file is being actively modified
- **Controls disappear when window inactive**: `ForceRedraw()` timer (100ms) uses `RedrawWindow` API with `RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN` flags to keep GUI controls visible
