# LiveView

A lightweight AutoHotkey v2 application that creates live thumbnail views of windows with customizable regions, widgets, and cycling backgrounds.

## Features

### DWM Thumbnail Regions
- Capture live thumbnails from any window using Windows DWM API
- Multiple independent regions, each with its own source window
- **Virtual desktop support**: Select windows from other desktops (automatically moved off-screen and restored on exit)
- Thumbnail preview for crop selection with zoom (scroll wheel) and pan (right-drag)
- Move, resize, and reorder regions (bring to front/send to back)

### Widgets
- **Clock Widget**: Customizable time/date display with font, color, and format options
- **Weather Widget**: Live weather data via WeatherAPI.com (requires free API key)

### Background System
- Automatic cycling through images in the `backgrounds` folder
- Supports `.jpg`, `.jpeg`, `.png`, `.bmp` formats
- 30-minute cycle interval

### Fullscreen Modes
- **Edit Fullscreen (E)**: Fullscreen but still editable
- **Locked Fullscreen (F11)**: True fullscreen, hides cursor, no editing

## Requirements

- Windows 10/11
- [AutoHotkey v2](https://www.autohotkey.com/) (to run the `.ahk` script)
- `curl.exe` in the script directory (for weather features) - [Download curl](https://curl.se/windows/)

## Installation

1. Download or clone this repository
2. Install AutoHotkey v2 if not already installed
3. (Optional) Download `curl.exe` and place it in the LiveView folder for weather functionality
4. Run `LiveView.ahk`

## Usage

### Controls

**Mouse:**
- Left-click + drag: Move region
- Right-click + drag: Resize region (on region/widget)
- Right-click on empty space: Open context menu
- Middle-click anywhere: Open context menu

**Keyboard - Region/Widget:**
- Arrow keys: Move (10px)
- Shift + Arrow keys: Resize (10px)
- PgUp/PgDn: Change z-order
- S: Select source area
- A: Add new region
- D: Delete region

**Keyboard - App:**
- W: Select source window
- Ctrl+S: Save configuration
- Ctrl+O: Load configuration
- E: Edit fullscreen
- F11: Locked fullscreen
- H: Toggle source visibility (moves windows off-screen)
- Escape: Exit current mode

### Crop Selection Preview
When selecting a source area (S key), a preview window opens:
- **Left-drag**: Draw selection rectangle
- **Scroll wheel**: Zoom in/out
- **Right-drag**: Pan around when zoomed
- **Escape**: Cancel selection

### Other Desktop Windows
Windows from other virtual desktops appear with `[Other Desktop]` prefix in the window selector. When selected:
- Window is automatically moved to the current desktop
- Window is positioned off-screen to avoid interference
- On app exit, windows are restored to their original position

### Setting Up Backgrounds

1. Place images in the `backgrounds` folder (created automatically)
2. Supported formats: `.jpg`, `.jpeg`, `.png`, `.bmp`
3. Images cycle automatically every 30 minutes

### Weather Widget Setup

1. Get a free API key from [WeatherAPI.com](https://www.weatherapi.com/)
2. Download `curl.exe` from [curl.se](https://curl.se/windows/) and place it in the LiveView folder
3. Add a Weather widget from the Widgets menu
4. Enter your API key and search for your location

## Configuration

Configurations are saved as `.ini` files and store:
- Region positions and source windows
- Widget settings (font, colors, formats)
- Weather location and API key

## License

MIT License
