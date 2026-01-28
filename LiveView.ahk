#Requires AutoHotkey v2.0
#SingleInstance Force

class ThumbnailViewer {
    thumbnails := []
    regions := []       ; Each region now has: srcL, srcT, srcR, srcB, destL, destT, destR, destB, hSource, sourceTitle, sourceExe
    selectedRegion := 1
    isFullscreen := false
    isEditFullscreen := false
    savedGuiX := 0
    savedGuiY := 0
    savedGuiW := 0
    savedGuiH := 0

    ; Drag state
    isDragging := false
    dragType := ""
    dragStartX := 0
    dragStartY := 0

    ; Source selection state
    clientOriginX := 0
    clientOriginY := 0
    clientW := 0
    clientH := 0
    isDrawing := false
    selectingSource := false
    initialized := false

    ; Hidden sources tracking
    hiddenSources := Map()

    ; Windows moved from other desktops (hwnd -> {origX, origY})
    movedFromOtherDesktop := Map()

    ; Widgets
    widgets := []
    widgetControls := []
    selectedWidget := 0
    editingWidgets := false  ; true when widget is selected instead of region

    ; Background image cycling (GDI+)
    bgImages := []
    bgImageIndex := 1
    bgCycleInterval := 1800000  ; 30 minutes per image
    bgLastChange := 0
    gdipToken := 0
    bgBitmap := 0

    ; Canvas scaling (design resolution for fullscreen preview)
    canvasW := 1920
    canvasH := 1080

    ; Not Live overlay controls
    notLiveOverlays := []

    ; Preview window for source selection
    previewGui := ""
    previewThumb := 0
    previewZoom := 1.0
    previewPanX := 0
    previewPanY := 0
    previewSourceW := 0
    previewSourceH := 0
    previewDrawing := false
    previewStartX := 0
    previewStartY := 0
    previewRect := ""

    ; Fullscreen click tracking
    fsClickCount := 0
    fsLastClickTime := 0
    messageGui := ""

    ; Inactive state tracking
    wasInactive := false
    lastMinute := ""
    lastWeatherText := ""
    lastBgIndex := 0

    __New(regionConfigs*) {
        ; Initialize regions - each region can have its own source
        for config in regionConfigs {
            ; Ensure region has source fields
            if !config.HasOwnProp("hSource")
                config.hSource := 0
            if !config.HasOwnProp("sourceTitle")
                config.sourceTitle := ""
            if !config.HasOwnProp("sourceExe")
                config.sourceExe := ""
            if !config.HasOwnProp("sourceClass")
                config.sourceClass := ""
            this.regions.Push(config)
        }

        ; Main viewer window with menu
        this.gui := Gui("+Resize +0x02000000", "Live View")  ; WS_CLIPCHILDREN
        this.gui.BackColor := "000000"
        this.gui.OnEvent("Close", (*) => this.Cleanup())
        this.gui.OnEvent("Size", (*) => this.OnResize())

        ; Initialize canvas to primary monitor resolution
        this.InitCanvas()

        ; Initialize GDI+
        this.InitGDIPlus()

        ; Load background images from folder
        this.LoadBackgroundImages()

        ; Load first background image if available
        if this.bgImages.Length > 0 {
            this.LoadBackgroundBitmap(this.bgImages[1])
            this.bgLastChange := A_TickCount
        }

        ; Handle WM_PAINT for GDI+ background drawing
        OnMessage(0x000F, ObjBindMethod(this, "OnPaint"))  ; WM_PAINT

        ; Handle window move messages
        OnMessage(0x0003, ObjBindMethod(this, "OnWindowMove"))  ; WM_MOVE
        OnMessage(0x0005, ObjBindMethod(this, "OnWindowMove"))  ; WM_SIZE

        ; Create menu
        this.fileMenu := Menu()
        this.fileMenu.Add("Select Source Window`tW", (*) => this.ShowWindowSelector())
        this.fileMenu.Add()
        this.fileMenu.Add("Save Config`tCtrl+S", (*) => this.SaveConfig())
        this.fileMenu.Add("Load Config`tCtrl+O", (*) => this.LoadConfig())
        this.fileMenu.Add()
        this.fileMenu.Add("Edit Fullscreen`tE", (*) => this.ToggleEditFullscreen())
        this.fileMenu.Add("Fullscreen (Locked)`tF11", (*) => this.ToggleFullscreen())
        this.fileMenu.Add()
        this.fileMenu.Add("Toggle Source Visibility`tH", (*) => this.ToggleSourceVisibility())
        this.fileMenu.Add()
        this.fileMenu.Add("Exit App", (*) => this.Cleanup())

        this.regionMenu := Menu()
        this.regionMenu.Add("Select Source Area`tS", (*) => this.StartSourceSelection())
        this.regionMenu.Add("Add New Region`tA", (*) => this.AddRegion())
        this.regionMenu.Add("Delete Region`tD", (*) => this.DeleteRegion())
        this.regionMenu.Add()
        this.regionMenu.Add("Bring to Front`tPgUp", (*) => this.BringToFront())
        this.regionMenu.Add("Send to Back`tPgDn", (*) => this.SendToBack())
        this.regionMenu.Add()
        this.regionMenu.Add("Copy Config to Clipboard", (*) => this.CopyConfig())

        this.widgetMenu := Menu()
        this.widgetMenu.Add("Add Clock Widget", (*) => this.AddClockWidget())
        this.widgetMenu.Add("Add Weather Widget", (*) => this.AddWeatherWidget())
        this.widgetMenu.Add()
        this.widgetMenu.Add("Delete Selected Widget", (*) => this.DeleteWidget())
        this.widgetMenu.Add()
        this.widgetMenu.Add("Configure Clock...", (*) => this.ConfigureClockWidget())
        this.widgetMenu.Add("Configure Weather...", (*) => this.ConfigureWeather())

        this.helpMenu := Menu()
        this.helpMenu.Add("Controls", (*) => this.ShowHelp())

        ; Create right-click context menu (combines all menus)
        this.contextMenu := Menu()
        this.contextMenu.Add("Select Source Window`tW", (*) => this.ShowWindowSelector())
        this.contextMenu.Add("Select Source Area`tS", (*) => this.StartSourceSelection())
        this.contextMenu.Add()
        this.contextMenu.Add("Add New Region`tA", (*) => this.AddRegion())
        this.contextMenu.Add("Delete Region`tD", (*) => this.DeleteRegion())
        this.contextMenu.Add()
        this.contextMenu.Add("Bring to Front`tPgUp", (*) => this.BringToFront())
        this.contextMenu.Add("Send to Back`tPgDn", (*) => this.SendToBack())
        this.contextMenu.Add()
        this.contextMenu.Add("Widgets", this.widgetMenu)
        this.contextMenu.Add()
        this.contextMenu.Add("Save Config`tCtrl+S", (*) => this.SaveConfig())
        this.contextMenu.Add("Load Config`tCtrl+O", (*) => this.LoadConfig())
        this.contextMenu.Add()
        this.contextMenu.Add("Edit Fullscreen`tE", (*) => this.ToggleEditFullscreen())
        this.contextMenu.Add("Fullscreen (Locked)`tF11", (*) => this.ToggleFullscreen())
        this.contextMenu.Add()
        this.contextMenu.Add("Toggle Source Visibility`tH", (*) => this.ToggleSourceVisibility())
        this.contextMenu.Add("Help", (*) => this.ShowHelp())
        this.contextMenu.Add()
        this.contextMenu.Add("Exit App", (*) => this.Cleanup())

        ; Status bar / region selector at bottom
        this.gui.SetFont("s10")
        this.regionDropdown := this.gui.AddDropDownList("w150 Choose1", this.GetRegionList())
        this.regionDropdown.OnEvent("Change", (*) => this.OnRegionSelect())

        ; Exit button for edit fullscreen mode (hidden by default)
        this.exitButton := this.gui.AddButton("w100 h30 Hidden", "Exit (Esc)")
        this.exitButton.OnEvent("Click", (*) => this.ToggleEditFullscreen())

        ; Widget selector dropdown
        this.widgetDropdown := this.gui.AddDropDownList("w150 Choose0", ["No Widgets"])
        this.widgetDropdown.OnEvent("Change", (*) => this.OnWidgetSelect())
        this.widgetDropdown.Visible := false

        ; Weather settings
        this.weatherLocation := "New York"
        this.weatherLat := 40.71
        this.weatherLon := -74.01
        this.weatherUnit := "fahrenheit"
        this.weatherText := "Weather: --"
        this.weatherApiKey := ""
        this.weatherRefreshInterval := 15  ; minutes (0 = never)
        this.lastWeatherFetch := 0

        ; Missing source check timing
        this.appStartTime := A_TickCount
        this.missingSourceCheckInterval := 30000  ; starts at 30 seconds

        ; Load API config if exists
        this.LoadAPIConfig()

        ; Start widget update timer
        SetTimer(() => this.UpdateWidgets(), 1000)

        ; Start weather refresh timer (checks every minute)
        SetTimer(() => this.CheckWeatherRefresh(), 60000)

        ; Start background image cycling timer
        SetTimer(() => this.AnimateBackground(), 1000)

        ; Periodic redraw timer to keep controls visible when window is inactive
        SetTimer(() => this.ForceRedraw(), 500)

        ; Check for missing source windows periodically (starts immediately)
        SetTimer(() => this.CheckMissingSources(), -1000)

        this.gui.Show("w800 h600 x300")

        ; Enable DWM composition for this window (MARGINS: left, right, top, bottom all = -1)
        margins := Buffer(16, 0)
        NumPut("Int", -1, margins, 0)   ; left
        NumPut("Int", -1, margins, 4)   ; right
        NumPut("Int", -1, margins, 8)   ; top
        NumPut("Int", -1, margins, 12)  ; bottom
        DllCall("dwmapi\DwmExtendFrameIntoClientArea", "Ptr", this.gui.Hwnd, "Ptr", margins)

        ; Position dropdown at bottom
        this.PositionControls()

        ; Register thumbnails for regions that have sources
        for i, region in this.regions {
            if region.hSource
                this.RegisterThumbnailForRegion(i)
        }
        this.UpdateAllThumbnails()

        ; Hotkeys - only active when this window is focused
        HotIfWinActive("ahk_id " this.gui.Hwnd)
        HotKey("h", (*) => this.ToggleSourceVisibility())
        HotKey("Escape", (*) => this.HandleEscape())
        HotKey("e", (*) => this.ToggleEditFullscreen())
        HotKey("w", (*) => this.ShowWindowSelector())
        HotKey("s", (*) => this.StartSourceSelection())
        HotKey("a", (*) => this.AddRegion())
        HotKey("d", (*) => this.DeleteRegion())
        HotKey("PgUp", (*) => this.BringToFront())
        HotKey("PgDn", (*) => this.SendToBack())

        ; Arrow keys to move region
        HotKey("Up", (*) => this.NudgeRegion(0, -10))
        HotKey("Down", (*) => this.NudgeRegion(0, 10))
        HotKey("Left", (*) => this.NudgeRegion(-10, 0))
        HotKey("Right", (*) => this.NudgeRegion(10, 0))

        ; Shift+Arrow to resize region
        HotKey("+Up", (*) => this.ResizeRegion(0, -10))
        HotKey("+Down", (*) => this.ResizeRegion(0, 10))
        HotKey("+Left", (*) => this.ResizeRegion(-10, 0))
        HotKey("+Right", (*) => this.ResizeRegion(10, 0))

        ; Save/Load config
        HotKey("^s", (*) => this.SaveConfig())
        HotKey("^o", (*) => this.LoadConfig())
        HotIfWinActive()  ; Reset hotkey context

        ; Global hotkeys - work even in click-through fullscreen
        HotKey("^+Escape", (*) => this.Cleanup())
        HotKey("F11", (*) => this.ToggleFullscreen())
        HotKey("^F11", (*) => this.ToggleFullscreen())

        ; Mouse tracking
        this.SetupMouseTracking()

        ; Mark as initialized
        this.initialized := true

        ; Send background to back of z-order
        this.SendBackgroundToBack()

        ; Continuous thumbnail refresh to handle window movement
        SetTimer(() => this.RefreshThumbnails(), 16)  ; ~60fps

        ; Try to load default config silently
        defaultConfig := A_ScriptDir "\LiveViewConfig.ini"
        if FileExist(defaultConfig)
            this.LoadConfigSilent(defaultConfig)
        else if this.regions.Length > 0 && !this.regions[1].hSource
            SetTimer(() => this.ShowWindowSelector(), -100)  ; Run once after GUI shown
    }

    LoadBackgroundImages() {
        ; Scan backgrounds folder for images
        bgFolder := A_ScriptDir "\backgrounds"
        if !DirExist(bgFolder) {
            DirCreate(bgFolder)
            return
        }

        ; Find all image files
        Loop Files, bgFolder "\*.jpg" {
            this.bgImages.Push(A_LoopFileFullPath)
        }
        Loop Files, bgFolder "\*.jpeg" {
            this.bgImages.Push(A_LoopFileFullPath)
        }
        Loop Files, bgFolder "\*.png" {
            this.bgImages.Push(A_LoopFileFullPath)
        }
        Loop Files, bgFolder "\*.bmp" {
            this.bgImages.Push(A_LoopFileFullPath)
        }
    }

    InitGDIPlus() {
        ; Load GDI+ library explicitly
        if !DllCall("GetModuleHandle", "Str", "gdiplus", "Ptr")
            DllCall("LoadLibrary", "Str", "gdiplus")

        ; Initialize GDI+
        si := Buffer(24, 0)
        NumPut("UInt", 1, si, 0)  ; GdiplusVersion
        token := 0
        result := DllCall("gdiplus\GdiplusStartup", "Ptr*", &token, "Ptr", si, "Ptr", 0, "Int")
        this.gdipToken := (result = 0) ? token : 0
    }

    ShutdownGDIPlus() {
        ; Free background bitmap first
        if this.bgBitmap {
            try DllCall("gdiplus\GdipDisposeImage", "Ptr", this.bgBitmap)
            this.bgBitmap := 0
        }

        ; Shutdown GDI+ (only if it was initialized)
        if this.gdipToken {
            try DllCall("gdiplus\GdiplusShutdown", "Ptr", this.gdipToken)
            this.gdipToken := 0
        }
    }

    LoadBackgroundBitmap(imagePath) {
        if !this.gdipToken
            return

        ; Free old bitmap
        if this.bgBitmap {
            try DllCall("gdiplus\GdipDisposeImage", "Ptr", this.bgBitmap)
            this.bgBitmap := 0
        }

        ; Load new image
        bitmap := 0
        result := DllCall("gdiplus\GdipCreateBitmapFromFile", "WStr", imagePath, "Ptr*", &bitmap, "Int")
        if (result = 0)
            this.bgBitmap := bitmap
    }

    OnPaint(wParam, lParam, msg, hwnd) {
        if (hwnd != this.gui.Hwnd) || !this.gdipToken
            return

        ; Begin paint
        ps := Buffer(72, 0)
        hdc := DllCall("BeginPaint", "Ptr", hwnd, "Ptr", ps, "Ptr")
        if !hdc
            return

        ; Get client size
        this.gui.GetClientPos(,, &clientW, &clientH)

        ; Create GDI+ Graphics from HDC
        graphics := 0
        DllCall("gdiplus\GdipCreateFromHDC", "Ptr", hdc, "Ptr*", &graphics)

        if graphics {
            ; Set high quality rendering
            DllCall("gdiplus\GdipSetInterpolationMode", "Ptr", graphics, "Int", 7)
            DllCall("gdiplus\GdipSetSmoothingMode", "Ptr", graphics, "Int", 4)
            DllCall("gdiplus\GdipSetTextRenderingHint", "Ptr", graphics, "Int", 5) ; ClearType

            ; Check if window is active (edit mode)
            activeHwnd := DllCall("GetForegroundWindow", "Ptr")
            isActive := (activeHwnd = this.gui.Hwnd)

            ; Draw background only when inactive or in fullscreen (not during edit mode)
            if this.bgBitmap && (!isActive || this.isFullscreen) {
                DllCall("gdiplus\GdipDrawImageRectI", "Ptr", graphics, "Ptr", this.bgBitmap,
                    "Int", 0, "Int", 0, "Int", clientW, "Int", clientH)
            }

            ; Draw widgets using GDI+ only when inactive or fullscreen
            ; When active: use text controls instead
            ; Draw weather widgets first, then clock widgets (so clock is on top)
            for pass in [1, 2] {
                for i, w in this.widgets {
                    ; Skip GDI+ widget drawing when active (use text controls instead)
                    if isActive && !this.isFullscreen
                        continue

                    ; Pass 1: weather only, Pass 2: clock only
                    if (pass = 1 && w.type != "weather") || (pass = 2 && w.type != "clock")
                        continue

                    ; Get widget text
                    if w.type = "clock" {
                        timeStr := FormatTime(, w.format)
                        dateStr := w.HasOwnProp("dateFormat") && w.dateFormat != "" ? FormatTime(, w.dateFormat) : ""
                        text := dateStr != "" ? timeStr "`n" dateStr : timeStr
                    } else if w.type = "weather" {
                        text := this.weatherText
                    } else {
                        continue
                    }

                ; Draw background rectangle if not transparent
                if (w.bgColor != "" && w.bgColor != "transparent") {
                    bgBrush := 0
                    bgColor := 0xFF000000 | Integer("0x" w.bgColor)
                    DllCall("gdiplus\GdipCreateSolidFill", "UInt", bgColor, "Ptr*", &bgBrush)
                    if bgBrush {
                        DllCall("gdiplus\GdipFillRectangleI", "Ptr", graphics, "Ptr", bgBrush,
                            "Int", w.x, "Int", w.y, "Int", w.width, "Int", w.height)
                        DllCall("gdiplus\GdipDeleteBrush", "Ptr", bgBrush)
                    }
                }

                ; Create font
                fontFamily := 0
                fontName := w.HasOwnProp("fontName") ? w.fontName : "Segoe UI"
                DllCall("gdiplus\GdipCreateFontFamilyFromName", "WStr", fontName, "Ptr", 0, "Ptr*", &fontFamily)
                if !fontFamily
                    DllCall("gdiplus\GdipCreateFontFamilyFromName", "WStr", "Segoe UI", "Ptr", 0, "Ptr*", &fontFamily)

                font := 0
                if fontFamily {
                    DllCall("gdiplus\GdipCreateFont", "Ptr", fontFamily, "Float", w.fontSize, "Int", 0, "Int", 2, "Ptr*", &font)
                }

                ; Create text brush
                textBrush := 0
                fontColor := 0xFF000000 | Integer("0x" w.fontColor)
                DllCall("gdiplus\GdipCreateSolidFill", "UInt", fontColor, "Ptr*", &textBrush)

                ; Create string format
                format := 0
                DllCall("gdiplus\GdipCreateStringFormat", "Int", 0, "Int", 0, "Ptr*", &format)
                if format {
                    ; Weather: left-aligned, top; Clock: centered
                    horizAlign := (w.type = "weather") ? 0 : 1  ; 0=Near/Left, 1=Center
                    vertAlign := (w.type = "weather") ? 0 : 1   ; 0=Near/Top, 1=Center
                    DllCall("gdiplus\GdipSetStringFormatAlign", "Ptr", format, "Int", horizAlign)
                    DllCall("gdiplus\GdipSetStringFormatLineAlign", "Ptr", format, "Int", vertAlign)
                }

                ; Draw text
                if font && textBrush && format {
                    rect := Buffer(16, 0)
                    NumPut("Float", w.x, rect, 0)
                    NumPut("Float", w.y, rect, 4)
                    NumPut("Float", w.width, rect, 8)
                    NumPut("Float", w.height, rect, 12)
                    DllCall("gdiplus\GdipDrawString", "Ptr", graphics, "WStr", text, "Int", -1,
                        "Ptr", font, "Ptr", rect, "Ptr", format, "Ptr", textBrush)
                }

                ; Cleanup
                if format
                    DllCall("gdiplus\GdipDeleteStringFormat", "Ptr", format)
                if textBrush
                    DllCall("gdiplus\GdipDeleteBrush", "Ptr", textBrush)
                if font
                    DllCall("gdiplus\GdipDeleteFont", "Ptr", font)
                if fontFamily
                    DllCall("gdiplus\GdipDeleteFontFamily", "Ptr", fontFamily)
                }
            }

            DllCall("gdiplus\GdipDeleteGraphics", "Ptr", graphics)
        }

        DllCall("EndPaint", "Ptr", hwnd, "Ptr", ps)
        return 0
    }

    OnResize() {
        this.UpdateAllThumbnails()
        this.UpdateBackgroundSize()
    }

    UpdateBackgroundSize() {
        if !this.bgBitmap || this.bgImages.Length = 0
            return
        ; Trigger repaint
        DllCall("InvalidateRect", "Ptr", this.gui.Hwnd, "Ptr", 0, "Int", true)
        ; Force immediate update when window is inactive
        activeHwnd := DllCall("GetForegroundWindow", "Ptr")
        if (activeHwnd != this.gui.Hwnd)
            DllCall("UpdateWindow", "Ptr", this.gui.Hwnd)
    }

    SendBackgroundToBack() {
        ; No longer needed for GDI+ - background is drawn first in OnPaint
        ; Background thumbnail is registered first so it's behind everything
    }

    RefreshThumbnails() {
        static lastX := -99999, lastY := -99999, lastW := 0, lastH := 0

        if !this.initialized
            return

        ; Check if any region has a source
        hasAnySource := false
        for r in this.regions {
            if r.hSource {
                hasAnySource := true
                break
            }
        }
        if !hasAnySource
            return

        try {
            WinGetPos(&x, &y, &w, &h, this.gui.Hwnd)
            if (x != lastX || y != lastY || w != lastW || h != lastH) {
                lastX := x
                lastY := y
                lastW := w
                lastH := h

                ; Just update properties and flush - don't re-register
                this.UpdateAllThumbnails()
                DllCall("dwmapi\DwmFlush")

                ; Force window redraw
                DllCall("RedrawWindow", "Ptr", this.gui.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0100 | 0x0001)  ; RDW_UPDATENOW | RDW_INVALIDATE
            }
        }
    }

    PositionControls() {
        if !this.HasOwnProp("regionDropdown")
            return
        this.gui.GetClientPos(,, &w, &h)
        this.regionDropdown.Move(5, h - 30, 150)
        if this.HasOwnProp("widgetDropdown") && this.widgets.Length > 0 {
            this.widgetDropdown.Move(160, h - 30, 150)
            this.widgetDropdown.Visible := !this.isFullscreen && !this.isEditFullscreen
        }
    }

    OnWindowMove(wParam, lParam, msg, hwnd) {
        if !this.initialized
            return
        static lastUpdate := 0
        ; Only respond to our main window, throttle updates
        if (hwnd = this.gui.Hwnd) {
            now := A_TickCount
            if (now - lastUpdate > 16) {  ; ~60fps max
                lastUpdate := now
                this.UpdateAllThumbnails()
                DllCall("dwmapi\DwmFlush")  ; Force compositor update
            }
        }
    }

    ShowHelp() {
        help := "
        (
CONTROLS:

Mouse (in viewer):
  Left-click + drag = Move region
  Right-click + drag = Resize region

Keyboard - Region/Widget:
  Arrow keys = Move (10px)
  Shift + Arrow keys = Resize (10px)
  PgUp = Bring to front
  PgDn = Send to back
  S = Select source area
  A = Add new region
  D = Delete region

Widgets (from menu):
  Add Clock or Weather widget
  Select widget from dropdown
  Move/resize with arrow keys

Keyboard - App:
  W = Select source window
  Ctrl+S = Save configuration
  Ctrl+O = Load configuration
  E = Edit Fullscreen
  F11 = Fullscreen (locked)
  H = Toggle source visibility
  Escape = Exit current mode

Backgrounds:
  Put images in 'backgrounds' folder
  Formats: .jpg, .jpeg, .png, .bmp
  Images cycle every 30 minutes
        )"
        this.ShowMessage(help, 10000)
    }

    HandleEscape() {
        if this.selectingSource
            this.CancelSelection()
        else if this.isFullscreen
            this.ToggleFullscreen()
        else if this.isEditFullscreen
            this.ToggleEditFullscreen()
    }

    ShowWindowSelector() {
        if this.isFullscreen || this.isEditFullscreen
            return

        ; Get list of visible windows
        windows := this.GetWindowList()
        if windows.Length = 0 {
            this.ShowMessage("No windows found")
            return
        }

        ; Create selector GUI
        this.selectorGui := Gui("+AlwaysOnTop +ToolWindow", "Select Source Window")
        this.selectorGui.SetFont("s10")
        this.selectorGui.AddText("w400", "Select a window to capture:")

        ; Create listbox with window titles
        windowTitles := []
        this.windowHandles := []
        for w in windows {
            windowTitles.Push(w.title)
            this.windowHandles.Push(w.hwnd)
        }

        this.windowListBox := this.selectorGui.AddListBox("w400 h300", windowTitles)
        this.windowListBox.OnEvent("DoubleClick", (*) => this.SelectWindow())

        this.selectorGui.AddButton("w100", "Select").OnEvent("Click", (*) => this.SelectWindow())
        this.selectorGui.AddButton("x+10 w100", "Cancel").OnEvent("Click", (*) => this.selectorGui.Destroy())
        this.selectorGui.AddButton("x+10 w100", "Refresh").OnEvent("Click", (*) => this.RefreshWindowList())

        this.selectorGui.OnEvent("Close", (*) => this.selectorGui.Destroy())
        this.selectorGui.Show()
    }

    GetWindowList() {
        windows := []

        ; Enable detection of windows on other virtual desktops
        DetectHiddenWindows(true)

        ; Use WinGetList to get all windows
        ids := WinGetList()

        for hwnd in ids {
            try {
                ; Get window title
                title := WinGetTitle(hwnd)
                if title = ""
                    continue

                ; Skip our own windows
                if InStr(title, "Live View") || InStr(title, "Select Source")
                    continue

                ; Skip certain system windows
                if title = "Program Manager" || title = "Windows Input Experience" || title = "Settings"
                    continue

                ; Check window style - must have WS_VISIBLE
                style := WinGetStyle(hwnd)
                if (style & 0x10000000) = 0  ; WS_VISIBLE
                    continue

                ; Check if on another virtual desktop (cloaked)
                cloaked := 0
                DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 14, "UInt*", &cloaked, "UInt", 4)

                ; Include if visible OR cloaked (other desktop)
                isVisible := DllCall("IsWindowVisible", "Ptr", hwnd)
                if !isVisible && !cloaked
                    continue

                ; Check if on another virtual desktop (DWM_CLOAKED_SHELL = 0x2)
                isOnOtherDesktop := (cloaked & 0x2) != 0
                displayTitle := isOnOtherDesktop ? "[Other Desktop] " title : title
                windows.Push({hwnd: hwnd, title: displayTitle, isOtherDesktop: isOnOtherDesktop})
            }
        }

        DetectHiddenWindows(false)
        return windows
    }

    RefreshWindowList() {
        if !this.HasOwnProp("selectorGui")
            return

        windows := this.GetWindowList()
        windowTitles := []
        this.windowHandles := []

        for w in windows {
            windowTitles.Push(w.title)
            this.windowHandles.Push(w.hwnd)
        }

        this.windowListBox.Delete()
        this.windowListBox.Add(windowTitles)
    }

    SelectWindow() {
        if !this.HasOwnProp("windowListBox") || !this.windowListBox.Value
            return

        idx := this.windowListBox.Value
        if idx > this.windowHandles.Length
            return

        newHwnd := this.windowHandles[idx]
        newTitle := this.windowListBox.Text

        ; Close selector
        this.selectorGui.Destroy()

        ; Change source window
        this.ChangeSourceWindow(newHwnd, newTitle)
    }

    SaveConfig() {
        if this.isFullscreen || this.isEditFullscreen
            return

        ; Save directly to default config file
        selectedFile := A_ScriptDir "\LiveViewConfig.ini"

        ; Delete existing file
        try FileDelete(selectedFile)

        ; Write number of regions
        IniWrite(this.regions.Length, selectedFile, "Regions", "Count")

        ; Write each region with its source info
        for i, r in this.regions {
            section := "Region" i
            IniWrite(r.srcL, selectedFile, section, "srcL")
            IniWrite(r.srcT, selectedFile, section, "srcT")
            IniWrite(r.srcR, selectedFile, section, "srcR")
            IniWrite(r.srcB, selectedFile, section, "srcB")
            IniWrite(r.destL, selectedFile, section, "destL")
            IniWrite(r.destT, selectedFile, section, "destT")
            IniWrite(r.destR, selectedFile, section, "destR")
            IniWrite(r.destB, selectedFile, section, "destB")

            ; Save per-region source info
            sourceExe := ""
            sourceTitle := ""
            sourceClass := ""
            if r.hSource {
                try {
                    sourceExe := WinGetProcessName(r.hSource)
                    sourceTitle := WinGetTitle(r.hSource)
                    sourceClass := WinGetClass(r.hSource)
                }
            }
            IniWrite(sourceExe, selectedFile, section, "sourceExe")
            IniWrite(sourceTitle, selectedFile, section, "sourceTitle")
            IniWrite(sourceClass, selectedFile, section, "sourceClass")
        }

        ; Write widgets
        IniWrite(this.widgets.Length, selectedFile, "Widgets", "Count")
        IniWrite(this.weatherLocation, selectedFile, "Widgets", "WeatherLocation")
        IniWrite(this.weatherUnit, selectedFile, "Widgets", "WeatherUnit")
        IniWrite(this.weatherLat, selectedFile, "Widgets", "WeatherLat")
        IniWrite(this.weatherLon, selectedFile, "Widgets", "WeatherLon")
        IniWrite(this.weatherRefreshInterval, selectedFile, "Widgets", "WeatherRefreshInterval")

        for i, w in this.widgets {
            section := "Widget" i
            IniWrite(w.type, selectedFile, section, "type")
            IniWrite(w.x, selectedFile, section, "x")
            IniWrite(w.y, selectedFile, section, "y")
            IniWrite(w.width, selectedFile, section, "width")
            IniWrite(w.height, selectedFile, section, "height")
            IniWrite(w.fontSize, selectedFile, section, "fontSize")
            IniWrite(w.fontColor, selectedFile, section, "fontColor")
            IniWrite(w.bgColor, selectedFile, section, "bgColor")
            if w.type = "clock" {
                IniWrite(w.format, selectedFile, section, "format")
                IniWrite(w.HasOwnProp("dateFormat") ? w.dateFormat : "ddd, MMM d", selectedFile, section, "dateFormat")
                IniWrite(w.HasOwnProp("fontName") ? w.fontName : "Segoe UI", selectedFile, section, "fontName")
            }
        }

        ; Show GUI message
        this.ShowMessage("Configuration saved", 2000)
    }

    LoadConfig() {
        if this.isFullscreen || this.isEditFullscreen
            return

        ; Get load file path
        selectedFile := FileSelect(1, A_ScriptDir, "Load Configuration", "Config Files (*.ini)")
        if !selectedFile
            return

        if !FileExist(selectedFile) {
            this.ShowMessage("File not found: " selectedFile)
            return
        }

        ; Read number of regions
        regionCount := IniRead(selectedFile, "Regions", "Count", 0)
        if regionCount = 0 {
            this.ShowMessage("No regions found in config file")
            return
        }

        ; Unregister existing thumbnails
        for thumb in this.thumbnails {
            if thumb
                DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", thumb)
        }
        this.thumbnails := []

        ; Restore any hidden sources (restore position)
        for hwnd, info in this.hiddenSources {
            try WinMove(info.x, info.y, info.w, info.h, hwnd)
        }
        this.hiddenSources := Map()

        ; Clear existing regions
        this.regions := []

        ; Enumerate all windows once (including other virtual desktops)
        allWindows := []
        prevHidden := A_DetectHiddenWindows
        DetectHiddenWindows(true)
        try {
            for winHwnd in WinGetList() {
                try {
                    allWindows.Push({
                        hwnd: winHwnd,
                        title: WinGetTitle(winHwnd),
                        exe: WinGetProcessName(winHwnd),
                        class: WinGetClass(winHwnd)
                    })
                }
            }
        }
        DetectHiddenWindows(prevHidden)

        ; Load each region with its source
        Loop regionCount {
            section := "Region" A_Index
            r := {}
            r.srcL := Integer(IniRead(selectedFile, section, "srcL", 0))
            r.srcT := Integer(IniRead(selectedFile, section, "srcT", 0))
            r.srcR := Integer(IniRead(selectedFile, section, "srcR", 200))
            r.srcB := Integer(IniRead(selectedFile, section, "srcB", 200))
            r.destL := Integer(IniRead(selectedFile, section, "destL", 0))
            r.destT := Integer(IniRead(selectedFile, section, "destT", 0))
            r.destR := Integer(IniRead(selectedFile, section, "destR", 200))
            r.destB := Integer(IniRead(selectedFile, section, "destB", 200))

            r.sourceExe := IniRead(selectedFile, section, "sourceExe", "")
            r.sourceTitle := IniRead(selectedFile, section, "sourceTitle", "")
            r.sourceClass := IniRead(selectedFile, section, "sourceClass", "")
            r.hSource := 0

            ; Try to find the source window from pre-enumerated list
            if r.sourceTitle != "" {
                for win in allWindows {
                    if r.sourceExe != "" && win.exe != r.sourceExe
                        continue
                    if win.title == r.sourceTitle {
                        if r.sourceClass != "" {
                            if win.class == r.sourceClass {
                                r.hSource := win.hwnd
                                break
                            }
                        } else {
                            r.hSource := win.hwnd
                            break
                        }
                    }
                }
            }
            ; Fallback: any window from exe only if no specific title was saved
            if !r.hSource && r.sourceExe != "" && r.sourceTitle == "" {
                for win in allWindows {
                    if win.exe == r.sourceExe {
                        r.hSource := win.hwnd
                        break
                    }
                }
            }

            ; If window was found, check if it's on another desktop and bring it over
            if r.hSource
                this.BringWindowFromOtherDesktop(r.hSource)

            this.regions.Push(r)
        }

        ; Update region dropdown
        this.regionDropdown.Delete()
        this.regionDropdown.Add(this.GetRegionList())
        this.regionDropdown.Value := 1
        this.selectedRegion := 1

        ; Register thumbnails for regions with sources
        this.ReRegisterAllThumbnails()

        ; Load widgets
        ; First, remove existing widgets
        for ctrl in this.widgetControls
            ctrl.Visible := false
        this.widgets := []
        this.widgetControls := []

        this.weatherLocation := IniRead(selectedFile, "Widgets", "WeatherLocation", "New York")
        this.weatherUnit := IniRead(selectedFile, "Widgets", "WeatherUnit", "fahrenheit")
        this.weatherLat := Float(IniRead(selectedFile, "Widgets", "WeatherLat", "40.71"))
        this.weatherLon := Float(IniRead(selectedFile, "Widgets", "WeatherLon", "-74.01"))
        this.weatherRefreshInterval := Integer(IniRead(selectedFile, "Widgets", "WeatherRefreshInterval", "15"))
        widgetCount := IniRead(selectedFile, "Widgets", "Count", 0)

        Loop widgetCount {
            section := "Widget" A_Index
            w := {}
            w.type := IniRead(selectedFile, section, "type", "clock")
            w.x := Integer(IniRead(selectedFile, section, "x", 10))
            w.y := Integer(IniRead(selectedFile, section, "y", 10))
            w.width := Integer(IniRead(selectedFile, section, "width", 200))
            w.height := Integer(IniRead(selectedFile, section, "height", 50))
            w.fontSize := Integer(IniRead(selectedFile, section, "fontSize", 24))
            w.fontColor := IniRead(selectedFile, section, "fontColor", "White")
            w.bgColor := IniRead(selectedFile, section, "bgColor", "000000")

            if w.type = "clock" {
                w.format := IniRead(selectedFile, section, "format", "h:mm:ss tt")
                w.dateFormat := IniRead(selectedFile, section, "dateFormat", "ddd, MMM d")
                w.fontName := IniRead(selectedFile, section, "fontName", "Segoe UI")
            }

            this.widgets.Push(w)
            this.CreateWidgetControl(this.widgets.Length)
        }

        this.UpdateWidgetDropdown()
        if widgetCount > 0
            this.FetchWeather()

        this.ShowMessage("Configuration loaded")

        ; Trigger immediate check for missing sources
        SetTimer(() => this.CheckMissingSources(), -1)
    }

    LoadConfigSilent(configFile) {
        ; Silent config loader - no dialogs, handles all errors gracefully
        try {
            if !FileExist(configFile)
                return

            regionCount := IniRead(configFile, "Regions", "Count", 0)
            if regionCount = 0
                return

            ; Unregister existing thumbnails
            for thumb in this.thumbnails {
                try if thumb
                    DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", thumb)
            }
            this.thumbnails := []

            ; Restore any hidden sources (restore position)
            for hwnd, info in this.hiddenSources {
                try WinMove(info.x, info.y, info.w, info.h, hwnd)
            }
            this.hiddenSources := Map()

            ; Clear existing regions
            this.regions := []

            ; Enumerate all windows once (including other virtual desktops)
            allWindows := []
            prevHidden := A_DetectHiddenWindows
            DetectHiddenWindows(true)
            try {
                for winHwnd in WinGetList() {
                    try {
                        allWindows.Push({
                            hwnd: winHwnd,
                            title: WinGetTitle(winHwnd),
                            exe: WinGetProcessName(winHwnd),
                            class: WinGetClass(winHwnd)
                        })
                    }
                }
            }
            DetectHiddenWindows(prevHidden)

            ; Load each region with its source
            Loop regionCount {
                try {
                    section := "Region" A_Index
                    r := {}
                    r.srcL := Integer(IniRead(configFile, section, "srcL", 0))
                    r.srcT := Integer(IniRead(configFile, section, "srcT", 0))
                    r.srcR := Integer(IniRead(configFile, section, "srcR", 200))
                    r.srcB := Integer(IniRead(configFile, section, "srcB", 200))
                    r.destL := Integer(IniRead(configFile, section, "destL", 0))
                    r.destT := Integer(IniRead(configFile, section, "destT", 0))
                    r.destR := Integer(IniRead(configFile, section, "destR", 200))
                    r.destB := Integer(IniRead(configFile, section, "destB", 200))

                    r.sourceExe := IniRead(configFile, section, "sourceExe", "")
                    r.sourceTitle := IniRead(configFile, section, "sourceTitle", "")
                    r.sourceClass := IniRead(configFile, section, "sourceClass", "")
                    r.hSource := 0

                    ; Try to find the source window from pre-enumerated list
                    if r.sourceTitle != "" {
                        for win in allWindows {
                            ; Filter by exe if specified
                            if r.sourceExe != "" && win.exe != r.sourceExe
                                continue
                            ; Strict exact title match
                            if win.title == r.sourceTitle {
                                if r.sourceClass != "" {
                                    if win.class == r.sourceClass {
                                        r.hSource := win.hwnd
                                        break
                                    }
                                } else {
                                    r.hSource := win.hwnd
                                    break
                                }
                            }
                        }
                    }
                    ; Fallback: any window from exe only if no specific title was saved
                    if !r.hSource && r.sourceExe != "" && r.sourceTitle == "" {
                        for win in allWindows {
                            if win.exe == r.sourceExe {
                                r.hSource := win.hwnd
                                break
                            }
                        }
                    }

                    ; If window was found, check if it's on another desktop and bring it over
                    if r.hSource
                        this.BringWindowFromOtherDesktop(r.hSource)

                    this.regions.Push(r)
                }
            }

            ; Ensure at least one region exists
            if this.regions.Length = 0 {
                this.regions.Push({srcL: 0, srcT: 0, srcR: 200, srcB: 200, destL: 0, destT: 0, destR: 200, destB: 200, hSource: 0, sourceTitle: "", sourceExe: "", sourceClass: ""})
            }

            ; Update region dropdown
            this.regionDropdown.Delete()
            this.regionDropdown.Add(this.GetRegionList())
            this.regionDropdown.Value := 1
            this.selectedRegion := 1

            ; Register thumbnails for regions with sources
            this.ReRegisterAllThumbnails()

            ; Load widgets - remove existing first
            for ctrl in this.widgetControls {
                try ctrl.Visible := false
            }
            this.widgets := []
            this.widgetControls := []

            try this.weatherLocation := IniRead(configFile, "Widgets", "WeatherLocation", "New York")
            try this.weatherUnit := IniRead(configFile, "Widgets", "WeatherUnit", "fahrenheit")
            try this.weatherLat := Float(IniRead(configFile, "Widgets", "WeatherLat", "40.71"))
            try this.weatherLon := Float(IniRead(configFile, "Widgets", "WeatherLon", "-74.01"))
            try this.weatherRefreshInterval := Integer(IniRead(configFile, "Widgets", "WeatherRefreshInterval", "15"))
            widgetCount := IniRead(configFile, "Widgets", "Count", 0)

            Loop widgetCount {
                try {
                    section := "Widget" A_Index
                    w := {}
                    w.type := IniRead(configFile, section, "type", "clock")
                    w.x := Integer(IniRead(configFile, section, "x", 10))
                    w.y := Integer(IniRead(configFile, section, "y", 10))
                    w.width := Integer(IniRead(configFile, section, "width", 200))
                    w.height := Integer(IniRead(configFile, section, "height", 50))
                    w.fontSize := Integer(IniRead(configFile, section, "fontSize", 24))
                    w.fontName := IniRead(configFile, section, "fontName", "Segoe UI")
                    w.fontColor := IniRead(configFile, section, "fontColor", "00FFAA")
                    w.bgColor := IniRead(configFile, section, "bgColor", "000000")

                    if w.type = "clock" {
                        w.format := IniRead(configFile, section, "format", "h:mm:ss tt")
                        w.dateFormat := IniRead(configFile, section, "dateFormat", "ddd, MMM d")
                    }

                    this.widgets.Push(w)
                    this.CreateWidgetControl(this.widgets.Length)
                }
            }

            this.UpdateWidgetDropdown()
            this.SendBackgroundToBack()

            ; Fetch weather if we have weather widgets
            for w in this.widgets {
                if w.type = "weather" {
                    this.FetchWeather()
                    break
                }
            }
        }
    }

    LoadAPIConfig() {
        apiKeyFile := A_ScriptDir "\api_keys.ini"
        if FileExist(apiKeyFile) {
            try this.weatherApiKey := IniRead(apiKeyFile, "Weather", "APIKey", "")
        }
    }

    SaveAPIConfig() {
        apiKeyFile := A_ScriptDir "\api_keys.ini"
        try {
            IniWrite(this.weatherApiKey, apiKeyFile, "Weather", "APIKey")
            return true
        } catch {
            return false
        }
    }

    ToggleAPIKeyVisibility(editCtrl, btn) {
        ; Toggle password mask on the edit control
        ; Save current value first
        currentValue := editCtrl.Value

        style := DllCall("GetWindowLong", "Ptr", editCtrl.Hwnd, "Int", -16, "Int")
        ES_PASSWORD := 0x20
        if (style & ES_PASSWORD) {
            ; Remove password style to show text
            DllCall("SetWindowLong", "Ptr", editCtrl.Hwnd, "Int", -16, "Int", style & ~ES_PASSWORD)
            btn.Text := "Hide"
        } else {
            ; Add password style to hide text
            DllCall("SetWindowLong", "Ptr", editCtrl.Hwnd, "Int", -16, "Int", style | ES_PASSWORD)
            btn.Text := "Show"
        }

        ; Re-set the value to force redraw with new style
        editCtrl.Value := ""
        editCtrl.Value := currentValue
    }

    SaveAPIKeyFromConfig(editCtrl) {
        this.weatherApiKey := editCtrl.Value
        if this.SaveAPIConfig() {
            this.ShowMessage("API Key saved", 2000)
            ; Refresh weather with new key
            this.FetchWeather()
        } else {
            this.ShowMessage("Failed to save API Key", 2000)
        }
    }

    ChangeSourceWindow(newHwnd, newTitle) {
        if this.selectedRegion > this.regions.Length
            return

        r := this.regions[this.selectedRegion]

        ; Restore old source window if it was hidden (restore position)
        if r.hSource && this.hiddenSources.Has(r.hSource) {
            try {
                info := this.hiddenSources[r.hSource]
                WinMove(info.x, info.y, info.w, info.h, r.hSource)
            }
            this.hiddenSources.Delete(r.hSource)
        }

        ; Check if window is from another desktop
        isFromOtherDesktop := InStr(newTitle, "[Other Desktop] ") = 1
        actualTitle := isFromOtherDesktop ? SubStr(newTitle, 18) : newTitle  ; Remove prefix

        ; If from other desktop and not already moved, bring it to current desktop
        if isFromOtherDesktop && !this.movedFromOtherDesktop.Has(newHwnd) {
            try {
                ; Store original position before moving
                WinGetPos(&origX, &origY, &origW, &origH, newHwnd)
                this.movedFromOtherDesktop[newHwnd] := {x: origX, y: origY, w: origW, h: origH}

                ; Move window to current desktop using COM interface
                if this.MoveWindowToCurrentDesktop(newHwnd) {
                    Sleep(100)
                    ; Move mostly off-screen (1 pixel visible to keep DWM live)
                    WinMove(this.GetBarelyOffScreenX(origW), origY,,, newHwnd)
                    this.ShowMessage("Window moved from other desktop (will restore on exit)")
                } else {
                    this.movedFromOtherDesktop.Delete(newHwnd)
                    this.ShowMessage("Could not move window from other desktop")
                }
            } catch as e {
                this.ShowMessage("Could not move window: " e.Message)
            }
        }

        ; Update region's source
        r.hSource := newHwnd
        r.sourceTitle := actualTitle  ; Store without prefix
        try r.sourceExe := WinGetProcessName(newHwnd)
        try r.sourceClass := WinGetClass(newHwnd)

        ; Get client rect of new source and set source rectangle to full window
        clientRect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", newHwnd, "Ptr", clientRect)
        r.srcL := 0
        r.srcT := 0
        r.srcR := NumGet(clientRect, 8, "Int")
        r.srcB := NumGet(clientRect, 12, "Int")

        ; Update region dropdown to show new source name
        currentValue := this.regionDropdown.Value
        this.regionDropdown.Delete()
        this.regionDropdown.Add(this.GetRegionList())
        this.regionDropdown.Value := currentValue

        ; Re-register thumbnail for this region
        this.ReRegisterThumbnailForRegion(this.selectedRegion)
        this.UpdateThumbnail(this.selectedRegion)
    }

    ; Move a window from another virtual desktop to the current one using COM
    MoveWindowToCurrentDesktop(targetHwnd) {
        ; IVirtualDesktopManager COM interface
        CLSID := Buffer(16)
        IID := Buffer(16)
        DllCall("ole32\CLSIDFromString", "WStr", "{aa509086-5ca9-4c25-8f95-589d3c07b48a}", "Ptr", CLSID)
        DllCall("ole32\CLSIDFromString", "WStr", "{a5cd92ff-29be-454c-8d04-d82879fb3f1b}", "Ptr", IID)

        pVDM := 0
        hr := DllCall("ole32\CoCreateInstance", "Ptr", CLSID, "Ptr", 0, "UInt", 1, "Ptr", IID, "Ptr*", &pVDM)
        if (hr != 0 || !pVDM)
            return false

        ; Get current desktop ID from our own window
        desktopId := Buffer(16)
        hr := ComCall(4, pVDM, "Ptr", this.gui.Hwnd, "Ptr", desktopId)  ; GetWindowDesktopId
        if (hr != 0) {
            ObjRelease(pVDM)
            return false
        }

        ; Move target window to current desktop
        hr := ComCall(5, pVDM, "Ptr", targetHwnd, "Ptr", desktopId)  ; MoveWindowToDesktop
        ObjRelease(pVDM)
        return (hr = 0)
    }

    ; Check if window is on another desktop and bring it to current if so
    BringWindowFromOtherDesktop(hwnd) {
        if this.movedFromOtherDesktop.Has(hwnd)
            return  ; Already moved

        ; Check if window is cloaked (on another virtual desktop)
        cloaked := 0
        DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 14, "UInt*", &cloaked, "UInt", 4)

        if (cloaked & 0x2) {  ; DWM_CLOAKED_SHELL = on another desktop
            try {
                ; Store original position
                WinGetPos(&origX, &origY, &origW, &origH, hwnd)
                this.movedFromOtherDesktop[hwnd] := {x: origX, y: origY, w: origW, h: origH}

                ; Move window to current desktop using COM interface
                if this.MoveWindowToCurrentDesktop(hwnd) {
                    Sleep(100)
                    ; Move mostly off-screen (1 pixel visible to keep DWM live)
                    WinMove(this.GetBarelyOffScreenX(origW), origY,,, hwnd)
                } else {
                    this.movedFromOtherDesktop.Delete(hwnd)
                }
            }
        }
    }

    ReRegisterThumbnailForRegion(index) {
        if index > this.regions.Length
            return

        ; Unregister existing thumbnail if any
        if index <= this.thumbnails.Length && this.thumbnails[index] {
            DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", this.thumbnails[index])
            this.thumbnails[index] := 0
        }

        ; Register new thumbnail
        this.RegisterThumbnailForRegion(index)
    }

    ToggleEditFullscreen() {
        if this.isFullscreen
            return

        if this.isEditFullscreen {
            ; Exit edit fullscreen
            this.gui.Opt("+Caption +Border")
            WinSetAlwaysOnTop(false, this.gui.Hwnd)
            this.gui.Move(this.savedGuiX, this.savedGuiY, this.savedGuiW, this.savedGuiH)
            this.regionDropdown.Visible := true
            if this.widgets.Length > 0
                this.widgetDropdown.Visible := true
            this.exitButton.Visible := false
            this.isEditFullscreen := false
        } else {
            ; Save current position
            WinGetPos(&x, &y, &w, &h, this.gui.Hwnd)
            this.savedGuiX := x
            this.savedGuiY := y
            this.savedGuiW := w
            this.savedGuiH := h

            ; Get monitor dimensions
            MonitorGet(MonitorGetPrimary(), &mLeft, &mTop, &mRight, &mBottom)

            ; Enter edit fullscreen - fullscreen but editable
            this.gui.Opt("-Caption -Border")
            this.regionDropdown.Visible := false
            this.widgetDropdown.Visible := false
            WinSetAlwaysOnTop(true, this.gui.Hwnd)
            this.gui.Move(mLeft, mTop, mRight - mLeft, mBottom - mTop)

            ; Show exit button
            this.exitButton.Visible := true
            this.PositionExitButton()

            this.isEditFullscreen := true
        }
        this.UpdateAllThumbnails()
        this.UpdateBackgroundSize()
        this.SendBackgroundToBack()
    }

    PositionExitButton() {
        this.gui.GetClientPos(,, &w, &h)
        this.exitButton.Move(w - 110, 10, 100, 30)
    }

    ToggleFullscreen() {
        if this.isFullscreen {
            ; Restore windowed mode
            this.gui.Opt("+Caption +Border")
            WinSetAlwaysOnTop(false, this.gui.Hwnd)
            this.gui.Move(this.savedGuiX, this.savedGuiY, this.savedGuiW, this.savedGuiH)
            this.regionDropdown.Visible := true
            if this.widgets.Length > 0
                this.widgetDropdown.Visible := true
            this.isFullscreen := false

            ; Remove click-through styles
            WinSetExStyle("-0x80020", this.gui.Hwnd)
        } else {
            ; Save current position
            WinGetPos(&x, &y, &w, &h, this.gui.Hwnd)
            this.savedGuiX := x
            this.savedGuiY := y
            this.savedGuiW := w
            this.savedGuiH := h

            ; Get monitor work area (full screen including taskbar area)
            MonitorGet(MonitorGetPrimary(), &mLeft, &mTop, &mRight, &mBottom)

            ; Go true fullscreen
            this.gui.Opt("-Caption -Border")
            this.regionDropdown.Visible := false
            this.widgetDropdown.Visible := false

            ; Set topmost and cover entire screen
            WinSetAlwaysOnTop(true, this.gui.Hwnd)
            this.gui.Move(mLeft, mTop, mRight - mLeft, mBottom - mTop)

            this.isFullscreen := true

            ; Add click-through styles (mouse passes to apps behind)
            WinSetExStyle("+0x80020", this.gui.Hwnd)
        }
        this.UpdateAllThumbnails()
        this.UpdateBackgroundSize()
        this.SendBackgroundToBack()
    }

    SetupMouseTracking() {
        SetTimer(() => this.CheckViewerMouse(), 16)
    }

    CheckViewerMouse() {
        static wasLDown := false
        static wasRDown := false

        if !this.initialized
            return

        lDown := GetKeyState("LButton", "P")

        ; Track clicks in locked fullscreen mode
        if this.isFullscreen {
            if (lDown && !wasLDown) {
                ; Reset count if too much time passed (1 second)
                if (A_TickCount - this.fsLastClickTime > 1000)
                    this.fsClickCount := 0

                this.fsClickCount++
                this.fsLastClickTime := A_TickCount

                ; Show message after 3 clicks
                if (this.fsClickCount >= 3) {
                    this.ShowMessage("Press F11 to exit fullscreen", 2000)
                    this.fsClickCount := 0
                }
            }
            wasLDown := lDown
            return
        }

        rDown := GetKeyState("RButton", "P")

        ; Only track mouse when LiveView is the active window (editing mode)
        if !this.isDragging && !WinActive("ahk_id " this.gui.Hwnd) {
            wasLDown := lDown
            wasRDown := rDown
            return
        }

        ; Get mouse position in client coordinates
        CoordMode("Mouse", "Client")
        MouseGetPos(&localX, &localY, &win)

        ; If currently dragging, continue the drag
        if this.isDragging {
            if (lDown || rDown) {
                dx := localX - this.dragStartX
                dy := localY - this.dragStartY
                this.UpdateDrag(dx, dy)
            } else {
                this.isDragging := false
            }
            wasLDown := lDown
            wasRDown := rDown
            return
        }

        ; Only start new drag if mouse is over our window
        if (win != this.gui.Hwnd) {
            wasLDown := lDown
            wasRDown := rDown
            return
        }

        ; Get client size and check bounds
        this.gui.GetClientPos(,, &cw, &ch)

        ; Must be inside client area and not on dropdown at bottom
        if (localX < 0 || localY < 0 || localX >= cw || localY >= ch - 40) {
            wasLDown := lDown
            wasRDown := rDown
            return
        }

        ; Only select and start drag on actual click, not hover
        if (lDown && !wasLDown) {
            ; Check if clicking on a widget first (widgets are on top)
            clickedWidget := this.GetWidgetAtPoint(localX, localY)
            if clickedWidget > 0 {
                ; Select this widget
                this.selectedWidget := clickedWidget
                if this.widgets.Length > 0
                    this.widgetDropdown.Value := clickedWidget
                this.editingWidgets := true
            } else {
                ; Check if clicking on a region
                clickedRegion := this.GetRegionAtPoint(localX, localY)
                if clickedRegion > 0 {
                    ; Select this region
                    this.selectedRegion := clickedRegion
                    this.regionDropdown.Value := clickedRegion
                    this.editingWidgets := false

                }
            }

            ; Start move drag
            this.isDragging := true
            this.dragType := "move"
            this.dragStartX := localX
            this.dragStartY := localY
            this.SaveDragState()
        }
        else if (rDown && !wasRDown) {
            ; Check if clicking on empty space - show context menu
            clickedWidget := this.GetWidgetAtPoint(localX, localY)
            clickedRegion := this.GetRegionAtPoint(localX, localY)
            if (clickedWidget = 0 && clickedRegion = 0) {
                ; Right-click on empty space - show context menu
                this.contextMenu.Show()
            } else {
                ; Start resize drag on region/widget
                this.isDragging := true
                this.dragType := "resize"
                this.dragStartX := localX
                this.dragStartY := localY
                this.SaveDragState()
            }
        }

        ; Middle-click shows context menu anywhere
        mDown := GetKeyState("MButton", "P")
        static wasMDown := false
        if (mDown && !wasMDown) {
            this.contextMenu.Show()
        }
        wasMDown := mDown

        wasLDown := lDown
        wasRDown := rDown
    }

    GetWidgetAtPoint(x, y) {
        ; Check widgets in reverse order (top to bottom)
        i := this.widgets.Length
        while i >= 1 {
            w := this.widgets[i]
            if (x >= w.x && x <= w.x + w.width && y >= w.y && y <= w.y + w.height)
                return i
            i--
        }
        return 0
    }

    GetRegionAtPoint(x, y) {
        ; Check regions in reverse order (top-most region first)
        ; Regions later in array are rendered on top
        i := this.regions.Length
        while i >= 1 {
            r := this.regions[i]
            ; Check if point is inside this region's destination rectangle
            if (x >= r.destL && x <= r.destR && y >= r.destT && y <= r.destB)
                return i
            i--
        }
        return 0
    }

    SaveDragState() {
        ; Save widget state if editing widgets
        if this.editingWidgets && this.selectedWidget > 0 && this.selectedWidget <= this.widgets.Length {
            w := this.widgets[this.selectedWidget]
            this.savedWidgetX := w.x
            this.savedWidgetY := w.y
            this.savedWidgetW := w.width
            this.savedWidgetH := w.height
            this.savedWidgetFontSize := w.HasOwnProp("fontSize") ? w.fontSize : 14
            return
        }

        ; Save region state
        if (this.selectedRegion > this.regions.Length)
            return
        r := this.regions[this.selectedRegion]
        this.savedDestL := r.destL
        this.savedDestT := r.destT
        this.savedDestR := r.destR
        this.savedDestB := r.destB
    }

    UpdateDrag(dx, dy) {
        ; Handle widget drag
        if this.editingWidgets && this.selectedWidget > 0 && this.selectedWidget <= this.widgets.Length {
            w := this.widgets[this.selectedWidget]
            ctrl := this.widgetControls[this.selectedWidget]

            if (this.dragType = "move") {
                w.x := this.savedWidgetX + dx
                w.y := this.savedWidgetY + dy
                ctrl.Move(w.x, w.y)
            }
            else if (this.dragType = "resize") {
                newW := Max(100, this.savedWidgetW + dx)
                newH := Max(40, this.savedWidgetH + dy)

                ; Scale font size proportionally from original
                if this.savedWidgetH > 0 {
                    scaleFactor := newH / this.savedWidgetH
                    w.fontSize := Max(10, Round(this.savedWidgetFontSize * scaleFactor))
                }

                w.width := newW
                w.height := newH

                ; Recreate control with new font size
                this.RecreateWidgetControl(this.selectedWidget)
            }
            return
        }

        ; Handle region drag
        if (this.selectedRegion > this.regions.Length)
            return
        r := this.regions[this.selectedRegion]

        ; Convert delta from window coords to canvas coords
        scale := this.GetScaleFactor()
        if (scale = 0)
            scale := 1
        canvasDX := Round(dx / scale)
        canvasDY := Round(dy / scale)

        if (this.dragType = "move") {
            w := this.savedDestR - this.savedDestL
            h := this.savedDestB - this.savedDestT
            r.destL := this.savedDestL + canvasDX
            r.destT := this.savedDestT + canvasDY
            r.destR := r.destL + w
            r.destB := r.destT + h
        }
        else if (this.dragType = "resize") {
            r.destR := Max(this.savedDestL + 20, this.savedDestR + canvasDX)
            r.destB := Max(this.savedDestT + 20, this.savedDestB + canvasDY)
        }

        this.UpdateThumbnail(this.selectedRegion)
    }

    StartSourceSelection() {
        if this.selectedRegion > this.regions.Length
            return

        r := this.regions[this.selectedRegion]

        if !r.hSource || !WinExist(r.hSource) {
            this.ShowMessage("No source window selected. Press W first.")
            return
        }

        if this.isFullscreen
            this.ToggleFullscreen()

        ; Use thumbnail preview for source selection
        this.StartSourceSelectionOnThumbnail()
    }

    StartSourceSelectionOnThumbnail() {
        r := this.regions[this.selectedRegion]

        ; Get source window client size
        clientRect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", r.hSource, "Ptr", clientRect)
        srcW := NumGet(clientRect, 8, "Int")
        srcH := NumGet(clientRect, 12, "Int")

        if srcW < 10 || srcH < 10 {
            this.ShowMessage("Cannot get source window size")
            return
        }

        ; Store source dimensions
        this.thumbPreviewSrcW := srcW
        this.thumbPreviewSrcH := srcH

        ; Initialize zoom and pan state
        this.thumbZoomLevel := 1.0
        this.thumbPanX := 0
        this.thumbPanY := 0
        this.isPanning := false
        this.panStartX := 0
        this.panStartY := 0
        this.panStartPanX := 0
        this.panStartPanY := 0

        ; Calculate initial preview size (70% of screen, maintain aspect ratio)
        screenW := A_ScreenWidth
        screenH := A_ScreenHeight
        maxW := Round(screenW * 0.7)
        maxH := Round(screenH * 0.7) - 50  ; Reserve space for instructions
        this.thumbBaseScale := Min(maxW / srcW, maxH / srcH, 1.0)
        previewW := Round(srcW * this.thumbBaseScale)
        previewH := Round(srcH * this.thumbBaseScale)

        ; Store effective scale for coordinate conversion
        this.thumbPreviewScale := this.thumbBaseScale * this.thumbZoomLevel

        ; Create preview window (resizable and maximizable)
        this.thumbPreviewGui := Gui("+AlwaysOnTop +Resize", "Crop Preview - Draw to select, Wheel to zoom, Right-drag to pan")
        this.thumbPreviewGui.BackColor := "222222"
        this.thumbPreviewGui.OnEvent("Close", (*) => this.CancelThumbnailSelection())
        this.thumbPreviewGui.OnEvent("Size", (*) => this.OnThumbnailPreviewResize())

        ; Add instruction text
        this.thumbPreviewGui.SetFont("s10 cWhite")
        this.thumbPreviewGui.AddText("w" previewW " Center vInstructionText", "Left-drag: Select | Wheel: Zoom | Right-drag: Pan | Esc: Cancel")

        ; Create a child window/control area for the thumbnail
        this.thumbPreviewGui.AddText("w" previewW " h" previewH " Background000000 vThumbArea")

        ; Add zoom indicator
        this.thumbPreviewGui.AddText("x10 y+5 w100 cWhite vZoomText", "Zoom: 100%")

        this.thumbPreviewGui.Show("AutoSize")

        ; Get the control and its screen position
        this.UpdateThumbnailAreaPosition()

        ; Register a DWM thumbnail for the preview
        previewThumb := 0
        result := DllCall("dwmapi\DwmRegisterThumbnail",
            "Ptr", this.thumbPreviewGui.Hwnd,
            "Ptr", r.hSource,
            "Ptr*", &previewThumb)
        this.previewThumb := previewThumb

        if result != 0 || !this.previewThumb {
            this.ShowMessage("Failed to create preview thumbnail")
            this.thumbPreviewGui.Destroy()
            return
        }

        ; Update thumbnail properties
        this.UpdateThumbnailPreviewProps()

        ; Selection rectangle (green border)
        this.selRect := Gui("+AlwaysOnTop -Caption +ToolWindow")
        this.selRect.BackColor := "00FF00"
        WinSetTransparent(150, this.selRect.Hwnd)

        this.selectingOnThumbnail := true
        this.isDrawing := false
        this.selStartX := 0
        this.selStartY := 0

        ; Register mouse wheel handler for zoom
        this.wheelHandler := ObjBindMethod(this, "OnThumbnailWheel")
        OnMessage(0x020A, this.wheelHandler)  ; WM_MOUSEWHEEL

        SetTimer(() => this.CheckThumbnailSelection(), 16)
    }

    OnThumbnailWheel(wParam, lParam, msg, hwnd) {
        if !this.selectingOnThumbnail
            return

        ; Get wheel delta (positive = up/zoom in, negative = down/zoom out)
        delta := (wParam >> 16) & 0xFFFF
        if (delta > 0x7FFF)
            delta := delta - 0x10000

        ; Get mouse position from lParam
        mouseX := lParam & 0xFFFF
        mouseY := (lParam >> 16) & 0xFFFF
        if (mouseX > 0x7FFF)
            mouseX := mouseX - 0x10000
        if (mouseY > 0x7FFF)
            mouseY := mouseY - 0x10000

        ; Call zoom handler
        direction := (delta > 0) ? 1 : -1
        this.HandleThumbnailZoom(direction, mouseX, mouseY)

        return 0  ; Prevent default handling
    }

    UpdateThumbnailAreaPosition() {
        if !this.thumbPreviewGui
            return

        thumbAreaCtrl := this.thumbPreviewGui["ThumbArea"]
        thumbAreaCtrl.GetPos(&areaX, &areaY, &areaW, &areaH)

        ; Convert control position to screen coordinates
        pt := Buffer(8, 0)
        NumPut("Int", areaX, pt, 0)
        NumPut("Int", areaY, pt, 4)
        DllCall("ClientToScreen", "Ptr", this.thumbPreviewGui.Hwnd, "Ptr", pt)

        this.thumbAreaScreenX := NumGet(pt, 0, "Int")
        this.thumbAreaScreenY := NumGet(pt, 4, "Int")
        this.thumbAreaW := areaW
        this.thumbAreaH := areaH
        this.thumbAreaClientX := areaX
        this.thumbAreaClientY := areaY
    }

    UpdateThumbnailPreviewProps() {
        if !this.previewThumb || !this.thumbPreviewGui
            return

        thumbAreaCtrl := this.thumbPreviewGui["ThumbArea"]
        thumbAreaCtrl.GetPos(&areaX, &areaY, &areaW, &areaH)

        ; Calculate the effective scale
        this.thumbPreviewScale := this.thumbBaseScale * this.thumbZoomLevel

        ; Calculate source region based on zoom and pan
        srcW := this.thumbPreviewSrcW
        srcH := this.thumbPreviewSrcH

        ; When zoomed, we show a portion of the source
        visibleSrcW := Round(areaW / this.thumbPreviewScale)
        visibleSrcH := Round(areaH / this.thumbPreviewScale)

        ; Clamp pan to valid range
        maxPanX := Max(0, srcW - visibleSrcW)
        maxPanY := Max(0, srcH - visibleSrcH)
        this.thumbPanX := Max(0, Min(this.thumbPanX, maxPanX))
        this.thumbPanY := Max(0, Min(this.thumbPanY, maxPanY))

        ; Source rectangle (what part of the source window to show)
        srcL := Round(this.thumbPanX)
        srcT := Round(this.thumbPanY)
        srcR := Min(srcW, srcL + visibleSrcW)
        srcB := Min(srcH, srcT + visibleSrcH)

        ; Store visible source bounds for coordinate conversion
        this.thumbVisibleSrcL := srcL
        this.thumbVisibleSrcT := srcT
        this.thumbVisibleSrcR := srcR
        this.thumbVisibleSrcB := srcB

        ; Update thumbnail properties
        props := Buffer(48, 0)
        NumPut("UInt", 0x1F, props, 0)  ; flags
        NumPut("Int", areaX, props, 4)   ; dest left
        NumPut("Int", areaY, props, 8)   ; dest top
        NumPut("Int", areaX + areaW, props, 12)  ; dest right
        NumPut("Int", areaY + areaH, props, 16)  ; dest bottom
        NumPut("Int", srcL, props, 20)   ; src left
        NumPut("Int", srcT, props, 24)   ; src top
        NumPut("Int", srcR, props, 28)   ; src right
        NumPut("Int", srcB, props, 32)   ; src bottom
        NumPut("UChar", 255, props, 36)  ; opacity
        NumPut("Int", 1, props, 40)      ; visible
        NumPut("Int", 1, props, 44)      ; client area only

        DllCall("dwmapi\DwmUpdateThumbnailProperties", "Ptr", this.previewThumb, "Ptr", props)

        ; Update zoom indicator
        try {
            zoomPct := Round(this.thumbZoomLevel * 100)
            this.thumbPreviewGui["ZoomText"].Value := "Zoom: " zoomPct "%"
        }
    }

    OnThumbnailPreviewResize() {
        if !this.thumbPreviewGui || !this.selectingOnThumbnail
            return

        ; Get new client size
        this.thumbPreviewGui.GetClientPos(,, &clientW, &clientH)

        ; Resize the thumbnail area control to fill available space
        thumbAreaCtrl := this.thumbPreviewGui["ThumbArea"]
        instructionCtrl := this.thumbPreviewGui["InstructionText"]
        zoomCtrl := this.thumbPreviewGui["ZoomText"]

        ; Reserve space for instruction text (top) and zoom text (bottom)
        instructionCtrl.Move(5, 5, clientW - 10)
        instructionCtrl.GetPos(,, &instrW, &instrH)

        areaTop := instrH + 10
        areaH := clientH - areaTop - 30  ; Reserve space for zoom text
        if areaH < 50
            areaH := 50

        thumbAreaCtrl.Move(0, areaTop, clientW, areaH)
        zoomCtrl.Move(10, areaTop + areaH + 5)

        ; Update positions and thumbnail
        this.UpdateThumbnailAreaPosition()
        this.UpdateThumbnailPreviewProps()
    }

    HandleThumbnailZoom(direction, mouseX, mouseY) {
        if !this.selectingOnThumbnail || !this.thumbPreviewGui
            return

        ; Calculate zoom factor
        zoomFactor := (direction > 0) ? 1.25 : 0.8
        oldZoom := this.thumbZoomLevel
        newZoom := oldZoom * zoomFactor

        ; Clamp zoom level (0.5x to 10x)
        newZoom := Max(0.5, Min(newZoom, 10.0))

        if newZoom = oldZoom
            return

        ; Calculate mouse position relative to thumbnail area
        relX := mouseX - this.thumbAreaScreenX
        relY := mouseY - this.thumbAreaScreenY

        ; Only adjust pan if mouse is inside the thumbnail area
        if (relX >= 0 && relX <= this.thumbAreaW && relY >= 0 && relY <= this.thumbAreaH) {
            ; Calculate the source position under the mouse before zoom
            oldScale := this.thumbBaseScale * oldZoom
            srcXUnderMouse := this.thumbPanX + (relX / oldScale)
            srcYUnderMouse := this.thumbPanY + (relY / oldScale)

            ; Update zoom
            this.thumbZoomLevel := newZoom
            newScale := this.thumbBaseScale * newZoom

            ; Adjust pan so the same source point stays under the mouse
            this.thumbPanX := srcXUnderMouse - (relX / newScale)
            this.thumbPanY := srcYUnderMouse - (relY / newScale)
        } else {
            this.thumbZoomLevel := newZoom
        }

        this.UpdateThumbnailPreviewProps()
    }

    CheckThumbnailSelection() {
        if !this.selectingOnThumbnail
            return

        CoordMode("Mouse", "Screen")
        MouseGetPos(&mx, &my)

        lDown := GetKeyState("LButton", "P")
        rDown := GetKeyState("RButton", "P")

        ; Check if mouse is within the thumbnail area
        inArea := (mx >= this.thumbAreaScreenX && mx <= this.thumbAreaScreenX + this.thumbAreaW
                && my >= this.thumbAreaScreenY && my <= this.thumbAreaScreenY + this.thumbAreaH)

        ; Handle right-click panning
        if (rDown && !this.isPanning && inArea && !this.isDrawing) {
            ; Start panning
            this.isPanning := true
            this.panStartX := mx
            this.panStartY := my
            this.panStartPanX := this.thumbPanX
            this.panStartPanY := this.thumbPanY
        }
        else if (rDown && this.isPanning) {
            ; Update pan position
            deltaX := mx - this.panStartX
            deltaY := my - this.panStartY

            ; Convert screen delta to source coordinates
            this.thumbPanX := this.panStartPanX - (deltaX / this.thumbPreviewScale)
            this.thumbPanY := this.panStartPanY - (deltaY / this.thumbPreviewScale)

            this.UpdateThumbnailPreviewProps()
        }
        else if (!rDown && this.isPanning) {
            ; Finish panning
            this.isPanning := false
        }

        ; Handle left-click drawing (only if not panning)
        if (!this.isPanning) {
            if (lDown && !this.isDrawing && inArea) {
                ; Start drawing - store position in SOURCE coordinates so it stays anchored when zooming
                this.isDrawing := true
                relX := mx - this.thumbAreaScreenX
                relY := my - this.thumbAreaScreenY
                this.selStartSrcX := this.thumbVisibleSrcL + (relX / this.thumbPreviewScale)
                this.selStartSrcY := this.thumbVisibleSrcT + (relY / this.thumbPreviewScale)
                this.selRect.Show("x" mx " y" my " w1 h1 NoActivate")
            }
            else if (lDown && this.isDrawing) {
                ; Update rectangle while drawing
                ; Convert source start position back to current screen position
                startScreenX := this.thumbAreaScreenX + Round((this.selStartSrcX - this.thumbVisibleSrcL) * this.thumbPreviewScale)
                startScreenY := this.thumbAreaScreenY + Round((this.selStartSrcY - this.thumbVisibleSrcT) * this.thumbPreviewScale)

                ; Clamp current mouse position to thumbnail area
                clampedX := Max(this.thumbAreaScreenX, Min(mx, this.thumbAreaScreenX + this.thumbAreaW))
                clampedY := Max(this.thumbAreaScreenY, Min(my, this.thumbAreaScreenY + this.thumbAreaH))

                ; Clamp start position too (in case zoomed out past original click)
                startScreenX := Max(this.thumbAreaScreenX, Min(startScreenX, this.thumbAreaScreenX + this.thumbAreaW))
                startScreenY := Max(this.thumbAreaScreenY, Min(startScreenY, this.thumbAreaScreenY + this.thumbAreaH))

                x1 := Min(startScreenX, clampedX)
                y1 := Min(startScreenY, clampedY)
                w := Abs(clampedX - startScreenX)
                h := Abs(clampedY - startScreenY)
                if (w < 1)
                    w := 1
                if (h < 1)
                    h := 1
                this.selRect.Move(x1, y1, w, h)
            }
            else if (!lDown && this.isDrawing) {
                ; Finished drawing
                this.FinishThumbnailSelection(mx, my)
            }
        }
    }

    FinishThumbnailSelection(endX, endY) {
        SetTimer(() => this.CheckThumbnailSelection(), 0)
        this.selectingOnThumbnail := false
        this.isDrawing := false

        ; Clamp end coordinates to thumbnail area
        endX := Max(this.thumbAreaScreenX, Min(endX, this.thumbAreaScreenX + this.thumbAreaW))
        endY := Max(this.thumbAreaScreenY, Min(endY, this.thumbAreaScreenY + this.thumbAreaH))

        ; Convert end position to source coordinates
        relEndX := endX - this.thumbAreaScreenX
        relEndY := endY - this.thumbAreaScreenY
        endSrcX := this.thumbVisibleSrcL + (relEndX / this.thumbPreviewScale)
        endSrcY := this.thumbVisibleSrcT + (relEndY / this.thumbPreviewScale)

        ; Use the source coordinates directly (selStartSrcX/Y were stored when drawing started)
        x1 := Round(Min(this.selStartSrcX, endSrcX))
        y1 := Round(Min(this.selStartSrcY, endSrcY))
        x2 := Round(Max(this.selStartSrcX, endSrcX))
        y2 := Round(Max(this.selStartSrcY, endSrcY))

        ; Clamp to source dimensions
        x1 := Max(0, Min(x1, this.thumbPreviewSrcW))
        y1 := Max(0, Min(y1, this.thumbPreviewSrcH))
        x2 := Max(0, Min(x2, this.thumbPreviewSrcW))
        y2 := Max(0, Min(y2, this.thumbPreviewSrcH))

        ; Ensure minimum size
        if (x2 - x1 < 10)
            x2 := x1 + 10
        if (y2 - y1 < 10)
            y2 := y1 + 10

        ; Update region
        if (this.selectedRegion <= this.regions.Length) {
            r := this.regions[this.selectedRegion]

            ; Set source crop
            r.srcL := x1
            r.srcT := y1
            r.srcR := x2
            r.srcB := y2

            ; Set destination to same size as source (1:1 scale)
            srcWidth := r.srcR - r.srcL
            srcHeight := r.srcB - r.srcT
            r.destR := r.destL + srcWidth
            r.destB := r.destT + srcHeight

            this.UpdateThumbnail(this.selectedRegion)
        }

        ; Cleanup
        try this.selRect.Destroy()
        if this.previewThumb {
            DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", this.previewThumb)
            this.previewThumb := 0
        }
        ; Unregister wheel handler
        if this.wheelHandler
            OnMessage(0x020A, this.wheelHandler, 0)
        try this.thumbPreviewGui.Destroy()
    }

    CancelThumbnailSelection() {
        if this.selectingOnThumbnail {
            SetTimer(() => this.CheckThumbnailSelection(), 0)
            this.selectingOnThumbnail := false
            this.isDrawing := false
            this.isPanning := false
            try this.selRect.Destroy()
            if this.previewThumb {
                DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", this.previewThumb)
                this.previewThumb := 0
            }
            ; Unregister wheel handler
            if this.wheelHandler
                OnMessage(0x020A, this.wheelHandler, 0)
            try this.thumbPreviewGui.Destroy()
        }
    }

    CancelSelection() {
        this.CancelThumbnailSelection()
    }

    GetRegionList() {
        list := []
        for i, r in this.regions {
            if r.sourceTitle != ""
                list.Push("Region " i " (" SubStr(r.sourceTitle, 1, 20) ")")
            else
                list.Push("Region " i " (no source)")
        }
        return list
    }

    OnRegionSelect() {
        this.selectedRegion := this.regionDropdown.Value
        this.editingWidgets := false

    }

    AddRegion() {
        if this.isFullscreen
            return

        ; Create new region with source fields initialized
        newRegion := {
            srcL: 0, srcT: 0, srcR: 200, srcB: 200,
            destL: 0, destT: 0, destR: 200, destB: 200,
            hSource: 0, sourceTitle: "", sourceExe: "", sourceClass: ""
        }
        this.regions.Push(newRegion)

        ; Add placeholder for thumbnail (no source yet)
        this.thumbnails.Push(0)

        this.regionDropdown.Delete()
        this.regionDropdown.Add(this.GetRegionList())
        this.regionDropdown.Value := this.regions.Length
        this.selectedRegion := this.regions.Length

        ; Prompt to select source for new region
        this.ShowMessage("New region added. Press W to select source.")
    }

    DeleteRegion() {
        if this.isFullscreen
            return
        if (this.regions.Length <= 1) {
            this.ShowMessage("Cannot delete the last region")
            return
        }
        idx := this.selectedRegion

        ; Unregister thumbnail if it exists
        if idx <= this.thumbnails.Length && this.thumbnails[idx]
            DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", this.thumbnails[idx])

        if idx <= this.thumbnails.Length
            this.thumbnails.RemoveAt(idx)
        this.regions.RemoveAt(idx)
        this.regionDropdown.Delete()
        this.regionDropdown.Add(this.GetRegionList())
        this.selectedRegion := Min(idx, this.regions.Length)
        this.regionDropdown.Value := this.selectedRegion
    }

    BringToFront() {
        if (this.isFullscreen || this.regions.Length <= 1 || this.selectedRegion = this.regions.Length)
            return

        idx := this.selectedRegion

        ; Move region to end of array (renders on top)
        region := this.regions.RemoveAt(idx)
        this.regions.Push(region)

        ; Re-register all thumbnails in new order
        this.ReRegisterAllThumbnails()

        ; Update dropdown to reflect new order
        this.regionDropdown.Delete()
        this.regionDropdown.Add(this.GetRegionList())

        ; Update selection to follow the moved region
        this.selectedRegion := this.regions.Length
        this.regionDropdown.Value := this.selectedRegion
    }

    SendToBack() {
        if (this.isFullscreen || this.regions.Length <= 1 || this.selectedRegion = 1)
            return

        idx := this.selectedRegion

        ; Move region to beginning of array (renders on bottom)
        region := this.regions.RemoveAt(idx)
        this.regions.InsertAt(1, region)

        ; Re-register all thumbnails in new order
        this.ReRegisterAllThumbnails()

        ; Update dropdown to reflect new order
        this.regionDropdown.Delete()
        this.regionDropdown.Add(this.GetRegionList())

        ; Update selection to follow the moved region
        this.selectedRegion := 1
        this.regionDropdown.Value := this.selectedRegion
    }

    ReRegisterAllThumbnails() {
        ; Unregister all existing thumbnails
        for thumb in this.thumbnails {
            if thumb
                DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", thumb)
        }

        this.thumbnails := []

        ; Re-register using each region's source
        for i, region in this.regions {
            if region.hSource {
                thumb := 0
                result := DllCall("dwmapi\DwmRegisterThumbnail",
                    "Ptr", this.gui.Hwnd,
                    "Ptr", region.hSource,
                    "Ptr*", &thumb)
                this.thumbnails.Push(result = 0 ? thumb : 0)
            } else {
                this.thumbnails.Push(0)
            }
        }

        this.UpdateAllThumbnails()
    }

    CheckMissingSources() {
        elapsedTime := A_TickCount - this.appStartTime

        ; Check if any sources are still missing
        hasMissing := false
        foundNew := false

        ; Enumerate all windows once
        allWindows := []
        try {
            for winHwnd in WinGetList() {
                try {
                    allWindows.Push({
                        hwnd: winHwnd,
                        title: WinGetTitle(winHwnd),
                        exe: WinGetProcessName(winHwnd),
                        class: WinGetClass(winHwnd)
                    })
                }
            }
        }

        for i, r in this.regions {
            ; Check if source is still valid (window exists AND title matches)
            if r.hSource {
                isValid := false
                if WinExist(r.hSource) {
                    ; Also verify the title still matches (if we have a saved title)
                    if r.sourceTitle != "" {
                        try {
                            currentTitle := WinGetTitle(r.hSource)
                            if currentTitle == r.sourceTitle
                                isValid := true
                        }
                    } else {
                        isValid := true  ; No title to verify, just existence
                    }
                }
                if isValid
                    continue  ; Source is valid, skip
                ; Window closed or title changed - clear the invalid handle and unregister thumbnail
                r.hSource := 0
                if this.thumbnails.Length >= i && this.thumbnails[i] {
                    DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", this.thumbnails[i])
                    this.thumbnails[i] := 0
                }
                ; Fall through to re-find the window
            }

            ; Skip if no source info saved
            if r.sourceExe = "" && r.sourceTitle = ""
                continue

            ; This region has saved info but no source - it's missing
            hasMissing := true

            ; Try to find the window from pre-enumerated list
            hwnd := 0

            ; If we have an exe, search windows for that exe
            if r.sourceExe != "" {
                for win in allWindows {
                    if win.exe != r.sourceExe
                        continue
                    ; Strict exact title + class match (most precise)
                    if r.sourceTitle != "" && r.sourceClass != "" && win.title == r.sourceTitle && win.class == r.sourceClass {
                        hwnd := win.hwnd
                        break
                    }
                    ; Strict exact title match
                    if r.sourceTitle != "" && win.title == r.sourceTitle {
                        hwnd := win.hwnd
                        break
                    }
                    ; If no title saved, take first window with this exe
                    if r.sourceTitle == "" {
                        hwnd := win.hwnd
                        break
                    }
                }
            }

            ; If no exe or not found by exe, try by exact title + class
            if !hwnd && r.sourceTitle != "" {
                for win in allWindows {
                    ; Strict exact title + class match first
                    if r.sourceClass != "" && win.title == r.sourceTitle && win.class == r.sourceClass {
                        hwnd := win.hwnd
                        break
                    }
                    ; Strict exact title only
                    if win.title == r.sourceTitle {
                        hwnd := win.hwnd
                        break
                    }
                }
            }

            ; If found, connect it
            if hwnd {
                r.hSource := hwnd
                foundNew := true
            }
        }

        ; Recheck if anything still missing
        hasMissing := false
        for i, r in this.regions {
            if !r.hSource && (r.sourceExe != "" || r.sourceTitle != "") {
                hasMissing := true
                break
            }
        }

        ; If ANY source is missing, unregister ALL thumbnails (don't display partial)
        if hasMissing {
            for i, thumb in this.thumbnails {
                if thumb {
                    DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", thumb)
                    this.thumbnails[i] := 0
                }
            }
        } else {
            ; All sources found - register thumbnails if we found new ones
            if foundNew {
                this.ReRegisterAllThumbnails()
            }
        }

        ; If no sources are missing, slow down to periodic checks (every 5 minutes)
        if !hasMissing {
            if this.missingSourceCheckInterval != 300000 {
                this.missingSourceCheckInterval := 300000
                SetTimer(() => this.CheckMissingSources(), 300000)
            }
            return
        }

        ; Sources still missing - check every 20 seconds for up to 5 minutes
        if elapsedTime < 300000 {
            if this.missingSourceCheckInterval != 20000 {
                this.missingSourceCheckInterval := 20000
                SetTimer(() => this.CheckMissingSources(), 20000)
            }
            return
        }

        ; After 5 minutes with missing sources, switch to 5-minute interval
        if this.missingSourceCheckInterval != 300000 {
            this.missingSourceCheckInterval := 300000
            SetTimer(() => this.CheckMissingSources(), 300000)
        }

        ; After 1 hour with missing sources, close the app
        if elapsedTime >= 3600000 {
            this.Cleanup()
            return
        }
    }

    NudgeRegion(dx, dy) {
        if this.isFullscreen
            return

        ; Check if editing widgets
        if this.editingWidgets && this.selectedWidget > 0 {
            this.NudgeWidget(dx, dy)
            return
        }

        if this.selectedRegion > this.regions.Length
            return
        r := this.regions[this.selectedRegion]
        w := r.destR - r.destL
        h := r.destB - r.destT
        r.destL += dx
        r.destT += dy
        r.destR := r.destL + w
        r.destB := r.destT + h
        this.UpdateThumbnail(this.selectedRegion)
    }

    ResizeRegion(dw, dh) {
        if this.isFullscreen
            return

        ; Check if editing widgets
        if this.editingWidgets && this.selectedWidget > 0 {
            this.ResizeWidget(dw, dh)
            return
        }

        if this.selectedRegion > this.regions.Length
            return
        r := this.regions[this.selectedRegion]
        newW := (r.destR - r.destL) + dw
        newH := (r.destB - r.destT) + dh
        if (newW >= 20)
            r.destR := r.destL + newW
        if (newH >= 20)
            r.destB := r.destT + newH
        this.UpdateThumbnail(this.selectedRegion)
    }

    RegisterThumbnailForRegion(index) {
        if index > this.regions.Length
            return

        r := this.regions[index]
        if !r.hSource
            return

        thumb := 0
        result := DllCall("dwmapi\DwmRegisterThumbnail",
            "Ptr", this.gui.Hwnd,
            "Ptr", r.hSource,
            "Ptr*", &thumb)

        if (result = 0) {
            ; Store thumbnail at the correct index
            while this.thumbnails.Length < index
                this.thumbnails.Push(0)
            if index <= this.thumbnails.Length
                this.thumbnails[index] := thumb
            else
                this.thumbnails.Push(thumb)
        }
    }

    ; ===== WIDGET METHODS =====

    AddClockWidget() {
        if this.isFullscreen
            return

        widget := {
            type: "clock",
            x: 10,
            y: 10,
            width: 350,
            height: 100,
            fontSize: 24,
            fontName: "Segoe UI",
            fontColor: "00FFAA",
            bgColor: "000000",
            format: "h:mm:ss tt",
            dateFormat: "ddd, MMM d"
        }

        this.widgets.Push(widget)
        this.CreateWidgetControl(this.widgets.Length)
        this.UpdateWidgetDropdown()
        this.selectedWidget := this.widgets.Length
        this.widgetDropdown.Value := this.selectedWidget
        this.editingWidgets := true
        this.SendBackgroundToBack()
    }

    AddWeatherWidget() {
        if this.isFullscreen
            return

        widget := {
            type: "weather",
            x: 10,
            y: 100,
            width: 320,
            height: 150,
            fontSize: 14,
            fontName: "Segoe UI",
            fontColor: "00FFAA",
            bgColor: "000000"
        }

        this.widgets.Push(widget)
        this.CreateWidgetControl(this.widgets.Length)
        this.UpdateWidgetDropdown()
        this.selectedWidget := this.widgets.Length
        this.widgetDropdown.Value := this.selectedWidget
        this.SendBackgroundToBack()

        ; Fetch weather
        this.FetchWeather()
    }

    CreateWidgetControl(index) {
        w := this.widgets[index]

        fontName := w.HasOwnProp("fontName") ? w.fontName : "Segoe UI"
        this.gui.SetFont("s" w.fontSize " c" w.fontColor, fontName)

        bgOpt := (w.bgColor != "" && w.bgColor != "transparent") ? " Background" w.bgColor : " BackgroundTrans"

        if w.type = "clock" {
            timeStr := FormatTime(, w.format)
            dateStr := w.HasOwnProp("dateFormat") && w.dateFormat != "" ? FormatTime(, w.dateFormat) : ""
            displayText := dateStr != "" ? timeStr "`n" dateStr : timeStr
            ctrl := this.gui.AddText("x" w.x " y" w.y " w" w.width " h" w.height bgOpt " Center", displayText)
        } else if w.type = "weather" {
            ctrl := this.gui.AddText("x" w.x " y" w.y " w" w.width " h" w.height bgOpt " Left Wrap", this.weatherText)
        }

        this.widgetControls.Push(ctrl)
        this.gui.SetFont()
    }

    ConfigureClockWidget() {
        if this.selectedWidget = 0 || this.selectedWidget > this.widgets.Length
            return
        w := this.widgets[this.selectedWidget]
        if w.type != "clock"
            return

        configGui := Gui("+AlwaysOnTop +ToolWindow", "Clock Settings")
        configGui.SetFont("s10")

        ; Font selection
        configGui.AddText("w320", "Font:")
        fontList := configGui.AddDropDownList("w320", ["Segoe UI", "Arial", "Consolas", "Courier New", "Verdana", "Tahoma", "Georgia", "Times New Roman"])
        fontList.Text := w.fontName

        ; Color presets
        configGui.AddText("w320", "Text Color:")
        colors := ["Cyan (00FFAA)", "White (FFFFFF)", "Yellow (FFFF00)", "Orange (FF8800)", "Red (FF0000)", "Pink (FF88FF)", "Blue (0088FF)", "Green (00FF00)", "Purple (AA00FF)"]
        colorList := configGui.AddDropDownList("w320", colors)
        currentColor := w.fontColor
        colorList.Choose(1)
        Loop colors.Length {
            if InStr(colors[A_Index], currentColor)
                colorList.Choose(A_Index)
        }

        ; Background color presets
        configGui.AddText("w320", "Background Color:")
        bgColors := ["Black (000000)", "Dark Gray (222222)", "Dark Blue (000033)", "Dark Green (003300)", "Dark Red (330000)", "Transparent"]
        bgList := configGui.AddDropDownList("w320", bgColors)
        if w.bgColor = "" || w.bgColor = "transparent"
            bgList.Choose(6)
        else {
            bgList.Choose(1)
            Loop bgColors.Length {
                if InStr(bgColors[A_Index], w.bgColor)
                    bgList.Choose(A_Index)
            }
        }

        ; Time format presets
        configGui.AddText("w320", "Time Format:")
        timeFormats := ["12-hour with seconds (h:mm:ss tt)", "12-hour no seconds (h:mm tt)", "24-hour with seconds (HH:mm:ss)", "24-hour no seconds (HH:mm)"]
        timeList := configGui.AddDropDownList("w320", timeFormats)
        if InStr(w.format, "tt")
            timeList.Choose(InStr(w.format, "ss") ? 1 : 2)
        else
            timeList.Choose(InStr(w.format, "ss") ? 3 : 4)

        ; Date format presets
        configGui.AddText("w320", "Date Format:")
        dateFormats := ["Short (ddd, MMM d)", "Medium (dddd, MMM d)", "Long (dddd, MMMM d, yyyy)", "Numeric (M/d/yyyy)", "No Date"]
        dateList := configGui.AddDropDownList("w320", dateFormats)
        if w.dateFormat = "" || w.dateFormat = "none"
            dateList.Choose(5)
        else if InStr(w.dateFormat, "yyyy") && InStr(w.dateFormat, "MMMM")
            dateList.Choose(3)
        else if InStr(w.dateFormat, "dddd")
            dateList.Choose(2)
        else if InStr(w.dateFormat, "/")
            dateList.Choose(4)
        else
            dateList.Choose(1)

        ; Preview section
        configGui.AddText("w320 h1 0x10")  ; Horizontal line (SS_ETCHEDHORZ)
        configGui.AddText("w320", "Preview:")
        previewBg := configGui.AddText("w320 h60 Background000000")
        configGui.SetFont("s16 c00FFAA", w.fontName)
        previewText := configGui.AddText("xp yp wp hp Center +0x200", FormatTime(, w.format) "`n" FormatTime(, w.dateFormat))
        configGui.SetFont("s10")

        ; Store references
        this.clockConfigGui := configGui
        this.clockConfigControls := {font: fontList, color: colorList, bg: bgList, time: timeList, date: dateList, preview: previewText, previewBg: previewBg, colors: colors, bgColors: bgColors, timeFormats: timeFormats, dateFormats: dateFormats}

        ; Add change handlers to update preview
        fontList.OnEvent("Change", (*) => this.UpdateClockPreview())
        colorList.OnEvent("Change", (*) => this.UpdateClockPreview())
        bgList.OnEvent("Change", (*) => this.UpdateClockPreview())
        timeList.OnEvent("Change", (*) => this.UpdateClockPreview())
        dateList.OnEvent("Change", (*) => this.UpdateClockPreview())

        configGui.AddButton("w155", "Apply").OnEvent("Click", (*) => this.ApplyClockConfig())
        configGui.AddButton("x+10 w155", "Cancel").OnEvent("Click", (*) => configGui.Destroy())

        configGui.Show()
    }

    ApplyClockConfig() {
        if this.selectedWidget = 0 || this.selectedWidget > this.widgets.Length
            return
        w := this.widgets[this.selectedWidget]
        c := this.clockConfigControls

        ; Get font
        w.fontName := c.font.Text

        ; Extract color from selection (e.g., "Cyan (00FFAA)" -> "00FFAA")
        colorText := c.color.Text
        if RegExMatch(colorText, "\(([A-Fa-f0-9]{6})\)", &m)
            w.fontColor := m[1]

        ; Extract bg color - empty string for transparent
        bgText := c.bg.Text
        if InStr(bgText, "none") || InStr(bgText, "Transparent")
            w.bgColor := ""
        else if RegExMatch(bgText, "\(([A-Fa-f0-9]{6})\)", &m)
            w.bgColor := m[1]

        ; Get time format
        timeFormats := Map(1, "h:mm:ss tt", 2, "h:mm tt", 3, "HH:mm:ss", 4, "HH:mm")
        w.format := timeFormats.Has(c.time.Value) ? timeFormats[c.time.Value] : "h:mm:ss tt"

        ; Get date format
        dateFormats := Map(1, "ddd, MMM d", 2, "dddd, MMM d", 3, "dddd, MMMM d, yyyy", 4, "M/d/yyyy", 5, "")
        w.dateFormat := dateFormats.Has(c.date.Value) ? dateFormats[c.date.Value] : "ddd, MMM d"

        ; Recreate the control with new settings
        this.RecreateWidgetControl(this.selectedWidget)

        this.clockConfigGui.Destroy()
    }

    UpdateClockPreview() {
        if !this.HasOwnProp("clockConfigControls")
            return

        c := this.clockConfigControls

        ; Get selected font color
        fontColor := "00FFAA"
        colorText := c.color.Text
        if RegExMatch(colorText, "\(([A-Fa-f0-9]{6})\)", &m)
            fontColor := m[1]

        ; Get selected background color
        bgColor := "000000"
        bgText := c.bg.Text
        if InStr(bgText, "Transparent")
            bgColor := "333333"  ; Show gray for transparent preview
        else if RegExMatch(bgText, "\(([A-Fa-f0-9]{6})\)", &m)
            bgColor := m[1]

        ; Get time format
        timeFormatsMap := Map(1, "h:mm:ss tt", 2, "h:mm tt", 3, "HH:mm:ss", 4, "HH:mm")
        timeFormat := timeFormatsMap.Has(c.time.Value) ? timeFormatsMap[c.time.Value] : "h:mm:ss tt"

        ; Get date format
        dateFormatsMap := Map(1, "ddd, MMM d", 2, "dddd, MMM d", 3, "dddd, MMMM d, yyyy", 4, "M/d/yyyy", 5, "")
        dateFormat := dateFormatsMap.Has(c.date.Value) ? dateFormatsMap[c.date.Value] : ""

        ; Update preview text
        timeStr := FormatTime(, timeFormat)
        dateStr := dateFormat != "" ? FormatTime(, dateFormat) : ""
        displayText := dateStr != "" ? timeStr "`n" dateStr : timeStr

        ; Update preview controls
        try {
            c.previewBg.Opt("Background" bgColor)
            this.clockConfigGui.SetFont("s16 c" fontColor, c.font.Text)
            c.preview.Text := displayText
            this.clockConfigGui.SetFont("s10")
        }
    }

    RecreateWidgetControl(index) {
        if index > this.widgets.Length || index > this.widgetControls.Length
            return

        w := this.widgets[index]
        oldCtrl := this.widgetControls[index]

        ; Hide old control
        oldCtrl.Visible := false

        ; Create new control with proper styling
        fontName := w.HasOwnProp("fontName") ? w.fontName : "Segoe UI"
        this.gui.SetFont("s" w.fontSize " c" w.fontColor, fontName)

        bgOpt := (w.bgColor != "" && w.bgColor != "transparent") ? " Background" w.bgColor : " BackgroundTrans"

        if w.type = "clock" {
            timeStr := FormatTime(, w.format)
            dateStr := w.HasOwnProp("dateFormat") && w.dateFormat != "" ? FormatTime(, w.dateFormat) : ""
            displayText := dateStr != "" ? timeStr "`n" dateStr : timeStr
            ctrl := this.gui.AddText("x" w.x " y" w.y " w" w.width " h" w.height bgOpt " Center", displayText)
        } else if w.type = "weather" {
            ctrl := this.gui.AddText("x" w.x " y" w.y " w" w.width " h" w.height bgOpt " Left Wrap", this.weatherText)
        }

        this.widgetControls[index] := ctrl
        this.gui.SetFont()
    }

    UpdateWidgetDropdown() {
        list := []
        for i, w in this.widgets {
            list.Push(w.type " " i)
        }

        ; Remember current selection
        prevSelection := this.selectedWidget

        this.widgetDropdown.Delete()
        if list.Length > 0 {
            this.widgetDropdown.Add(list)
            this.widgetDropdown.Visible := !this.isFullscreen && !this.isEditFullscreen

            ; Restore selection if valid, otherwise select first widget
            if prevSelection > 0 && prevSelection <= list.Length {
                this.widgetDropdown.Value := prevSelection
                this.selectedWidget := prevSelection
            } else if list.Length > 0 {
                this.widgetDropdown.Value := 1
                this.selectedWidget := 1
            }
            this.editingWidgets := true
        } else {
            this.widgetDropdown.Visible := false
            this.selectedWidget := 0
            this.editingWidgets := false
        }
    }

    OnWidgetSelect() {
        this.selectedWidget := this.widgetDropdown.Value
        this.editingWidgets := this.selectedWidget > 0
    }

    DeleteWidget() {
        if this.isFullscreen || this.selectedWidget = 0 || this.selectedWidget > this.widgets.Length
            return

        ; Destroy control
        this.widgetControls[this.selectedWidget].Visible := false

        ; Remove from arrays
        this.widgets.RemoveAt(this.selectedWidget)
        this.widgetControls.RemoveAt(this.selectedWidget)

        ; Update dropdown
        this.UpdateWidgetDropdown()
        this.selectedWidget := Min(this.selectedWidget, this.widgets.Length)
        if this.selectedWidget > 0
            this.widgetDropdown.Value := this.selectedWidget
        this.editingWidgets := this.selectedWidget > 0
    }

    UpdateWidgets() {
        if !this.initialized || this.widgets.Length = 0
            return

        ; Check if window is active
        activeHwnd := DllCall("GetForegroundWindow", "Ptr")
        isActive := (activeHwnd = this.gui.Hwnd)

        ; When active: use text controls for non-transparent widgets
        ; When active: use text controls for all widgets
        if isActive {
            for i, w in this.widgets {
                if i > this.widgetControls.Length
                    continue
                ctrl := this.widgetControls[i]
                ctrl.Visible := true

                if w.type = "clock" {
                    timeStr := FormatTime(, w.format)
                    dateStr := w.HasOwnProp("dateFormat") && w.dateFormat != "" ? FormatTime(, w.dateFormat) : ""
                    ctrl.Text := dateStr != "" ? timeStr "`n" dateStr : timeStr
                } else if w.type = "weather" {
                    ctrl.Text := this.weatherText
                }
            }
            ; Bring clock controls to top (above weather)
            for i, w in this.widgets {
                if w.type = "clock" && i <= this.widgetControls.Length {
                    DllCall("SetWindowPos", "Ptr", this.widgetControls[i].Hwnd, "Ptr", 0,
                        "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)  ; HWND_TOP, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE
                }
            }
            return
        }

        ; When inactive: hide controls, use GDI+ drawing
        for i, w in this.widgets {
            if i > this.widgetControls.Length
                continue
            this.widgetControls[i].Visible := false
        }

        ; Only invalidate specific widget rectangles that changed
        for w in this.widgets {
            needsUpdate := false

            if w.type = "clock" {
                if InStr(w.format, "s") {
                    needsUpdate := true
                } else {
                    currentMinute := FormatTime(, "mm")
                    if (currentMinute != this.lastMinute) {
                        this.lastMinute := currentMinute
                        needsUpdate := true
                    }
                }
            } else if w.type = "weather" {
                if (this.weatherText != this.lastWeatherText) {
                    this.lastWeatherText := this.weatherText
                    needsUpdate := true
                }
            }

            if needsUpdate {
                rect := Buffer(16, 0)
                NumPut("Int", w.x, rect, 0)
                NumPut("Int", w.y, rect, 4)
                NumPut("Int", w.x + w.width, rect, 8)
                NumPut("Int", w.y + w.height, rect, 12)
                DllCall("InvalidateRect", "Ptr", this.gui.Hwnd, "Ptr", rect, "Int", false)
            }
        }

        DllCall("UpdateWindow", "Ptr", this.gui.Hwnd)
    }

    AnimateBackground() {
        if !this.initialized || this.bgImages.Length = 0
            return

        ; Check if it's time to change the image
        if (A_TickCount - this.bgLastChange) >= this.bgCycleInterval {
            this.bgLastChange := A_TickCount
            this.bgImageIndex++
            if this.bgImageIndex > this.bgImages.Length
                this.bgImageIndex := 1

            ; Load new background image and trigger repaint
            this.LoadBackgroundBitmap(this.bgImages[this.bgImageIndex])
            this.UpdateBackgroundSize()
        }
    }

    ConfigureWeather() {
        if this.isFullscreen
            return

        w := (this.selectedWidget > 0 && this.selectedWidget <= this.widgets.Length && this.widgets[this.selectedWidget].type = "weather")
            ? this.widgets[this.selectedWidget] : ""

        configGui := Gui("+AlwaysOnTop +ToolWindow", "Weather Settings")
        configGui.SetFont("s10")

        ; Location search
        configGui.AddText("w320", "Search for your city:")
        searchEdit := configGui.AddEdit("w240", this.weatherLocation)
        searchBtn := configGui.AddButton("x+5 w75", "Search")

        configGui.AddText("xm w320", "Search Results (double-click to select):")
        resultsList := configGui.AddListBox("w320 h100", ["Enter a city name and click Search"])

        ; Temperature unit
        configGui.AddText("w320", "Temperature Unit:")
        unitList := configGui.AddDropDownList("w320", ["Fahrenheit (°F)", "Celsius (°C)"])
        unitList.Choose(this.weatherUnit = "celsius" ? 2 : 1)

        ; Font selection
        configGui.AddText("w320", "Font:")
        fonts := ["Segoe UI", "Arial", "Consolas", "Courier New", "Verdana", "Tahoma"]
        fontList := configGui.AddDropDownList("w320", fonts)
        fontList.Choose(1)
        if w && w.HasOwnProp("fontName") {
            Loop fonts.Length {
                if fonts[A_Index] = w.fontName
                    fontList.Choose(A_Index)
            }
        }

        ; Color presets
        configGui.AddText("w320", "Text Color:")
        colors := ["Cyan (00FFAA)", "White (FFFFFF)", "Yellow (FFFF00)", "Orange (FF8800)", "Red (FF0000)", "Pink (FF88FF)", "Blue (0088FF)", "Green (00FF00)"]
        colorList := configGui.AddDropDownList("w320", colors)
        colorList.Choose(1)
        if w {
            Loop colors.Length {
                if InStr(colors[A_Index], w.fontColor)
                    colorList.Choose(A_Index)
            }
        }

        ; Background color
        configGui.AddText("w320", "Background:")
        bgColors := ["Black (000000)", "Dark Gray (222222)", "Dark Blue (000033)", "Transparent"]
        bgList := configGui.AddDropDownList("w320", bgColors)
        bgList.Choose(1)
        if w {
            if w.bgColor = "" || w.bgColor = "transparent"
                bgList.Choose(4)
            else {
                Loop bgColors.Length {
                    if InStr(bgColors[A_Index], w.bgColor)
                        bgList.Choose(A_Index)
                }
            }
        }

        ; Refresh interval
        configGui.AddText("w320", "Auto-refresh interval:")
        refreshIntervals := ["Never", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "60 minutes"]
        refreshList := configGui.AddDropDownList("w320", refreshIntervals)
        ; Map current interval to dropdown index
        intervalMap := Map(0, 1, 5, 2, 10, 3, 15, 4, 30, 5, 60, 6)
        refreshList.Choose(intervalMap.Has(this.weatherRefreshInterval) ? intervalMap[this.weatherRefreshInterval] : 4)

        ; API Key section
        configGui.AddText("w320", "")  ; Spacer
        configGui.AddText("w320 cGray", "━━━━━━━━━━ API Key Settings ━━━━━━━━━━")
        configGui.AddText("w320", "Get a free API key from weatherapi.com:")
        apiLinkBtn := configGui.AddButton("w320", "Open weatherapi.com (Free signup)")
        apiLinkBtn.OnEvent("Click", (*) => Run("https://www.weatherapi.com/signup.aspx"))

        configGui.AddText("w320", "Enter your API Key:")
        apiKeyEdit := configGui.AddEdit("w240 Password", this.weatherApiKey)
        showKeyBtn := configGui.AddButton("x+5 w75", "Show")
        showKeyBtn.OnEvent("Click", (*) => this.ToggleAPIKeyVisibility(apiKeyEdit, showKeyBtn))

        saveKeyBtn := configGui.AddButton("xm w320", "Save API Key")
        saveKeyBtn.OnEvent("Click", (*) => this.SaveAPIKeyFromConfig(apiKeyEdit))

        this.weatherConfigGui := configGui
        this.weatherConfigControls := {search: searchEdit, results: resultsList, unit: unitList, font: fontList, color: colorList, bg: bgList, refresh: refreshList, apiKey: apiKeyEdit}
        this.weatherSearchResults := []

        searchBtn.OnEvent("Click", (*) => this.SearchWeatherLocation())
        resultsList.OnEvent("DoubleClick", (*) => this.SelectWeatherLocation())

        configGui.AddButton("w155", "Apply").OnEvent("Click", (*) => this.ApplyWeatherConfig())
        configGui.AddButton("x+10 w155", "Cancel").OnEvent("Click", (*) => configGui.Destroy())

        configGui.Show()
    }

    SearchWeatherLocation() {
        query := this.weatherConfigControls.search.Value
        if query = ""
            return

        this.weatherConfigControls.results.Delete()
        this.weatherConfigControls.results.Add(["Searching..."])
        this.weatherSearchResults := []

        ; Use Open-Meteo geocoding API
        tempFile := A_Temp "\liveview_geocode.json"
        try FileDelete(tempFile)

        curlPath := A_ScriptDir "\curl.exe"
        if !FileExist(curlPath) {
            this.weatherConfigControls.results.Delete()
            this.weatherConfigControls.results.Add(["curl.exe not found"])
            this.ShowMessage("curl.exe not found - download from curl.se/windows")
            return
        }
        url := "https://geocoding-api.open-meteo.com/v1/search?name=" query "&count=10&language=en&format=json"
        RunWait('"' curlPath '" -s -o "' tempFile '" "' url '"',, "Hide")

        try {
            response := FileRead(tempFile)
            this.weatherConfigControls.results.Delete()

            ; Parse each result block individually
            results := []
            pos := 1
            while RegExMatch(response, '\{[^{}]*"name":"([^"]+)"[^{}]*"latitude":([\d.-]+)[^{}]*"longitude":([\d.-]+)[^{}]*"country":"([^"]+)"[^{}]*\}', &m, pos) {
                city := m[1]
                lat := m[2]
                lon := m[3]
                country := m[4]

                ; Try to extract admin1 (state/province) from this block
                block := m[0]
                state := ""
                if RegExMatch(block, '"admin1":"([^"]+)"', &sm)
                    state := sm[1]

                ; Build display name: City, State, Country
                displayName := city
                if state != ""
                    displayName .= ", " state
                displayName .= ", " country

                results.Push(displayName)
                this.weatherSearchResults.Push({name: city, display: displayName, lat: Float(lat), lon: Float(lon)})
                pos := m.Pos + m.Len
            }

            if results.Length > 0
                this.weatherConfigControls.results.Add(results)
            else
                this.weatherConfigControls.results.Add(["No results found"])
        } catch {
            this.weatherConfigControls.results.Delete()
            this.weatherConfigControls.results.Add(["Search failed"])
        }
    }

    SelectWeatherLocation() {
        idx := this.weatherConfigControls.results.Value
        if idx = 0 || idx > this.weatherSearchResults.Length
            return

        result := this.weatherSearchResults[idx]
        this.weatherLocation := result.name
        this.weatherLat := result.lat
        this.weatherLon := result.lon
        this.weatherConfigControls.search.Value := result.display
    }

    ApplyWeatherConfig() {
        c := this.weatherConfigControls

        ; Update location from search field if no result was selected
        if this.weatherSearchResults.Length = 0
            this.weatherLocation := c.search.Value

        ; Update temperature unit
        this.weatherUnit := c.unit.Value = 2 ? "celsius" : "fahrenheit"

        ; Update refresh interval
        refreshValues := [0, 5, 10, 15, 30, 60]
        this.weatherRefreshInterval := refreshValues[c.refresh.Value]

        ; Update widget appearance if one is selected
        if this.selectedWidget > 0 && this.selectedWidget <= this.widgets.Length {
            w := this.widgets[this.selectedWidget]
            if w.type = "weather" {
                w.fontName := c.font.Text

                colorText := c.color.Text
                if RegExMatch(colorText, "\(([A-Fa-f0-9]{6})\)", &m)
                    w.fontColor := m[1]

                bgText := c.bg.Text
                if InStr(bgText, "Transparent") || InStr(bgText, "none")
                    w.bgColor := ""
                else if RegExMatch(bgText, "\(([A-Fa-f0-9]{6})\)", &m)
                    w.bgColor := m[1]

                this.RecreateWidgetControl(this.selectedWidget)
            }
        }

        this.weatherConfigGui.Destroy()
        this.FetchWeather()
    }

    FetchWeather() {
        this.lastWeatherFetch := A_TickCount
        ; Use Open-Meteo - fast and free
        coords := Map("New York", [40.71,-74.01], "Los Angeles", [34.05,-118.24], "Chicago", [41.88,-87.63], "London", [51.51,-0.13], "Tokyo", [35.68,139.69], "Miami", [25.76,-80.19], "Seattle", [47.61,-122.33], "Denver", [39.74,-104.99], "San Francisco", [37.77,-122.42])
        loc := this.weatherLocation
        if coords.Has(loc) {
            this.weatherLat := coords[loc][1]
            this.weatherLon := coords[loc][2]
        } else {
            this.weatherLat := 40.71
            this.weatherLon := -74.01
        }
        this.weatherText := loc "`nLoading..."
        SetTimer(() => this.DoFetchWeather(), -1)
    }

    DoFetchWeather() {
        try {
            if this.weatherApiKey = "" {
                this.weatherText := this.weatherLocation "`nNo API Key"
                return
            }

            ; Use WeatherAPI.com - supports city name search directly
            url := "https://api.weatherapi.com/v1/current.json?key=" this.weatherApiKey "&q=" this.weatherLocation "&aqi=no"
            this.weatherTempFile := A_Temp "\liveview_weather.json"

            ; Delete old file if exists
            try FileDelete(this.weatherTempFile)

            ; Use local 32-bit curl.exe
            curlPath := A_ScriptDir "\curl.exe"
            if !FileExist(curlPath) {
                this.weatherText := this.weatherLocation "`nNo curl.exe"
                return
            }
            Run('"' curlPath '" -s -o "' this.weatherTempFile '" "' url '"',, "Hide")

            ; Poll for file to appear
            SetTimer(() => this.CheckWeatherFile(), 500)
        } catch as e {
            this.weatherText := this.weatherLocation "`nOffline"
        }
    }

    CheckWeatherFile() {
        ; Check if download is complete
        if !FileExist(this.weatherTempFile)
            return  ; Still downloading

        ; Stop polling
        SetTimer(() => this.CheckWeatherFile(), 0)

        try {
            response := FileRead(this.weatherTempFile)

            ; Parse WeatherAPI.com response
            isCelsius := this.weatherUnit = "celsius"
            unitSymbol := isCelsius ? "C" : "F"

            ; Temperature
            tempKey := isCelsius ? "temp_c" : "temp_f"
            temp := RegExMatch(response, '"' tempKey '":\s*([\d.-]+)', &m) ? Round(m[1]) : "?"

            ; Feels like
            feelsKey := isCelsius ? "feelslike_c" : "feelslike_f"
            feels := RegExMatch(response, '"' feelsKey '":\s*([\d.-]+)', &f) ? Round(f[1]) : "?"

            ; Condition text
            condition := RegExMatch(response, '"condition":\s*\{[^}]*"text":\s*"([^"]+)"', &c) ? c[1] : ""

            ; Wind
            windKey := isCelsius ? "wind_kph" : "wind_mph"
            windUnit := isCelsius ? "km/h" : "mph"
            wind := RegExMatch(response, '"' windKey '":\s*([\d.-]+)', &w) ? Round(w[1]) : "?"
            windDir := RegExMatch(response, '"wind_dir":\s*"([^"]+)"', &wd) ? wd[1] : ""

            ; Humidity
            humidity := RegExMatch(response, '"humidity":\s*(\d+)', &h) ? h[1] : "?"

            ; Build display text
            this.weatherText := this.weatherLocation "`n"
            this.weatherText .= temp "°" unitSymbol " " condition "`n"
            this.weatherText .= "Feels: " feels "°" unitSymbol " | Wind: " wind windUnit " " windDir "`n"
            this.weatherText .= "Humidity: " humidity "%"
        } catch as e {
            this.weatherText := this.weatherLocation "`nError"
        }
    }

    CheckWeatherRefresh() {
        ; Skip if refresh is disabled (0 = never)
        if this.weatherRefreshInterval <= 0
            return

        ; Skip if no weather widgets exist
        hasWeatherWidget := false
        for w in this.widgets {
            if w.type = "weather" {
                hasWeatherWidget := true
                break
            }
        }
        if !hasWeatherWidget
            return

        ; Check if enough time has passed since last fetch
        intervalMs := this.weatherRefreshInterval * 60000
        if (A_TickCount - this.lastWeatherFetch) >= intervalMs
            this.FetchWeather()
    }

    NudgeWidget(dx, dy) {
        if this.selectedWidget = 0 || this.selectedWidget > this.widgets.Length
            return

        w := this.widgets[this.selectedWidget]
        w.x += dx
        w.y += dy

        ctrl := this.widgetControls[this.selectedWidget]
        ctrl.Move(w.x, w.y)
    }

    ResizeWidget(dw, dh) {
        if this.selectedWidget = 0 || this.selectedWidget > this.widgets.Length
            return

        w := this.widgets[this.selectedWidget]
        w.width := Max(150, w.width + dw)

        ; Weather needs more height for 4 lines, clock only needs 2 lines
        minHeight := w.type = "weather" ? 80 : 50
        w.height := Max(minHeight, w.height + dh)

        ; Calculate font size based on widget type and height
        ; Weather has 4 lines, clock has 2 lines
        divisor := w.type = "weather" ? 6 : 4
        w.fontSize := Max(10, Min(72, Round(w.height / divisor)))

        ; Recreate control with new size
        this.RecreateWidgetControl(this.selectedWidget)
    }

    CopyConfig() {
        config := "; Region configuration (each region can have its own source)`n"
        for i, r in this.regions {
            config .= "region" i " := {srcL: " r.srcL ", srcT: " r.srcT ", srcR: " r.srcR ", srcB: " r.srcB
            config .= ", destL: " r.destL ", destT: " r.destT ", destR: " r.destR ", destB: " r.destB
            config .= ", hSource: 0, sourceTitle: `"" r.sourceTitle "`", sourceExe: `"" r.sourceExe "`"}`n"
        }
        config .= "`nviewer := ThumbnailViewer("
        Loop this.regions.Length {
            if A_Index > 1
                config .= ", "
            config .= "region" A_Index
        }
        config .= ")"
        A_Clipboard := config
        this.ShowMessage("Configuration copied to clipboard")
    }

    ToggleSourceVisibility() {
        ; Collect unique source windows (only on current desktop, not cloaked)
        uniqueSources := Map()
        for r in this.regions {
            if r.hSource && !uniqueSources.Has(r.hSource) {
                ; Check if window is cloaked (on another virtual desktop)
                cloaked := 0
                DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", r.hSource, "UInt", 14, "UInt*", &cloaked, "UInt", 4)
                if !cloaked
                    uniqueSources[r.hSource] := true
            }
        }

        if uniqueSources.Count = 0
            return

        ; Check if any sources are hidden
        anyHidden := false
        for hwnd in uniqueSources {
            if this.hiddenSources.Has(hwnd) {
                anyHidden := true
                break
            }
        }

        if anyHidden {
            ; Restore all hidden sources - restore position
            for hwnd in uniqueSources {
                if this.hiddenSources.Has(hwnd) {
                    try {
                        info := this.hiddenSources[hwnd]
                        WinMove(info.x, info.y, info.w, info.h, hwnd)
                    }
                    this.hiddenSources.Delete(hwnd)
                }
            }
        } else {
            ; Hide sources by moving mostly off-screen (1 pixel visible to keep DWM live)
            for hwnd in uniqueSources {
                try {
                    WinGetPos(&x, &y, &w, &h, hwnd)
                    this.hiddenSources[hwnd] := {x: x, y: y, w: w, h: h}
                    WinMove(this.GetBarelyOffScreenX(w), y,,, hwnd)
                }
            }
        }
    }

    ; === Canvas Scaling Methods ===

    InitCanvas() {
        ; Set canvas to primary monitor resolution
        MonitorGet(MonitorGetPrimary(), &mLeft, &mTop, &mRight, &mBottom)
        this.canvasW := mRight - mLeft
        this.canvasH := mBottom - mTop
    }

    ; Get X position that keeps 1 pixel of a window on-screen (for DWM live capture)
    GetBarelyOffScreenX(windowWidth) {
        leftEdge := 0
        Loop MonitorGetCount() {
            MonitorGet(A_Index, &mLeft)
            if (A_Index = 1 || mLeft < leftEdge)
                leftEdge := mLeft
        }
        return leftEdge - windowWidth + 1
    }

    GetScaleFactor() {
        ; Returns scale factor: windowSize / canvasSize
        ; In fullscreen modes, scale is 1.0 (no scaling)
        if this.isFullscreen || this.isEditFullscreen
            return 1.0

        this.gui.GetClientPos(,, &cw, &ch)
        ; Use height-based scaling to maintain aspect ratio, leave space for controls
        availableH := ch - 40  ; Space for dropdown
        scaleX := cw / this.canvasW
        scaleY := availableH / this.canvasH
        return Min(scaleX, scaleY)
    }

    WindowToCanvas(x, y) {
        ; Convert window coordinates to canvas coordinates
        scale := this.GetScaleFactor()
        if scale = 0
            scale := 1
        return {x: Round(x / scale), y: Round(y / scale)}
    }

    IsSourceLive(hwnd) {
        ; Check if a source window is being rendered (live) or frozen
        if !hwnd
            return false

        ; Check if window is minimized
        if DllCall("IsIconic", "Ptr", hwnd)
            return false

        ; Check if window is cloaked (on another virtual desktop)
        cloaked := 0
        DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 14, "UInt*", &cloaked, "UInt", 4)
        if cloaked
            return false

        return true
    }

    UpdateAllThumbnails() {
        this.PositionControls()
        if this.isEditFullscreen
            this.PositionExitButton()
        Loop this.thumbnails.Length
            this.UpdateThumbnail(A_Index)
    }

    UpdateThumbnail(index) {
        if (index > this.thumbnails.Length || index > this.regions.Length)
            return

        thumb := this.thumbnails[index]
        r := this.regions[index]

        ; Skip if no thumbnail registered or no source
        if !thumb || !r.hSource
            return

        ; Get client area and calculate max bottom (leave space for dropdown when not fullscreen/edit fullscreen)
        this.gui.GetClientPos(,, &cw, &ch)
        maxBottom := (this.isFullscreen || this.isEditFullscreen) ? ch : ch - 40

        ; Scale destination coordinates from canvas to window
        scale := this.GetScaleFactor()
        scaledDestL := Round(r.destL * scale)
        scaledDestT := Round(r.destT * scale)
        scaledDestR := Round(r.destR * scale)
        scaledDestB := Min(Round(r.destB * scale), maxBottom)

        props := Buffer(48, 0)
        NumPut("UInt", 0x1F, props, 0)
        NumPut("Int", scaledDestL, props, 4)
        NumPut("Int", scaledDestT, props, 8)
        NumPut("Int", scaledDestR, props, 12)
        NumPut("Int", scaledDestB, props, 16)
        NumPut("Int", r.srcL, props, 20)
        NumPut("Int", r.srcT, props, 24)
        NumPut("Int", r.srcR, props, 28)
        NumPut("Int", r.srcB, props, 32)
        NumPut("UChar", 255, props, 36)
        NumPut("Int", 1, props, 40)
        NumPut("Int", 1, props, 44)

        DllCall("dwmapi\DwmUpdateThumbnailProperties", "Ptr", thumb, "Ptr", props)
    }

    ForceRedraw() {
        if !this.initialized
            return

        ; Keep window on top during fullscreen modes
        if (this.isEditFullscreen || this.isFullscreen) {
            ; HWND_TOPMOST = -1, SWP_NOMOVE = 0x2, SWP_NOSIZE = 0x1, SWP_NOACTIVATE = 0x10
            DllCall("SetWindowPos", "Ptr", this.gui.Hwnd, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)
        }

        ; Update "Not Live" overlay indicators
        this.UpdateNotLiveOverlays()
    }

    UpdateNotLiveOverlays() {
        ; Ensure we have enough overlay controls
        while this.notLiveOverlays.Length < this.regions.Length {
            ; Create overlay text control (initially hidden) with dark background
            this.gui.SetFont("s12 cFF4444 Bold", "Segoe UI")
            overlay := this.gui.AddText("w100 h24 Center Hidden Background222222 +0x200", "NOT LIVE")
            this.notLiveOverlays.Push(overlay)
            this.gui.SetFont()
        }

        ; Get scaling info
        scale := this.GetScaleFactor()

        ; Update each overlay
        for i, r in this.regions {
            if i > this.notLiveOverlays.Length
                continue

            overlay := this.notLiveOverlays[i]

            ; Check if source exists and is not live
            if r.hSource && !this.IsSourceLive(r.hSource) {
                ; Calculate position (center of region)
                scaledL := Round(r.destL * scale)
                scaledT := Round(r.destT * scale)
                scaledR := Round(r.destR * scale)
                scaledB := Round(r.destB * scale)

                regionW := scaledR - scaledL
                regionH := scaledB - scaledT
                overlayW := 100
                overlayH := 24
                overlayX := scaledL + (regionW - overlayW) // 2
                overlayY := scaledT + (regionH - overlayH) // 2

                overlay.Move(overlayX, overlayY, overlayW, overlayH)
                overlay.Visible := true

                ; Bring overlay to front
                DllCall("SetWindowPos", "Ptr", overlay.Hwnd, "Ptr", 0,
                    "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)
            } else {
                overlay.Visible := false
            }
        }
    }

    ShowMessage(text, duration := 2000) {
        ; Close existing message if any
        if this.messageGui {
            try this.messageGui.Destroy()
            this.messageGui := ""
        }

        ; Get main window position for centering
        try {
            WinGetPos(&winX, &winY, &winW, &winH, this.gui.Hwnd)
        } catch {
            winX := 0, winY := 0, winW := 800, winH := 600
        }

        ; Create message GUI
        msgGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")  ; E0x20 = click-through
        msgGui.BackColor := "222222"
        msgGui.SetFont("s14 cWhite", "Segoe UI")
        msgGui.AddText("Center w300 h40 +0x200", text)

        ; Calculate center position
        msgW := 320
        msgH := 60
        msgX := winX + (winW - msgW) // 2
        msgY := winY + (winH - msgH) // 2

        ; Make semi-transparent
        msgGui.Show("x" msgX " y" msgY " w" msgW " h" msgH " NoActivate")
        WinSetTransparent(200, msgGui.Hwnd)

        this.messageGui := msgGui

        ; Auto-close after duration
        SetTimer(() => this.CloseMessage(), -duration)
    }

    CloseMessage() {
        if this.messageGui {
            try this.messageGui.Destroy()
            this.messageGui := ""
        }
    }

    Cleanup() {
        SetTimer(() => this.CheckViewerMouse(), 0)
        SetTimer(() => this.RefreshThumbnails(), 0)
        SetTimer(() => this.UpdateWidgets(), 0)
        SetTimer(() => this.AnimateBackground(), 0)
        SetTimer(() => this.ForceRedraw(), 0)
        SetTimer(() => this.CheckWeatherRefresh(), 0)
        SetTimer(() => this.CheckMissingSources(), 0)

        ; Shutdown GDI+
        this.ShutdownGDIPlus()

        ; Restore all hidden source windows (restore position)
        for hwnd, info in this.hiddenSources {
            try WinMove(info.x, info.y, info.w, info.h, hwnd)
        }

        ; Restore windows that were moved from other desktops
        for hwnd, info in this.movedFromOtherDesktop {
            try WinMove(info.x, info.y, info.w, info.h, hwnd)
        }

        for thumb in this.thumbnails {
            if thumb
                DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", thumb)
        }
        ExitApp()
    }
}

; ===== MAIN =====
; Each region can have its own source window
region1 := {srcL: 0, srcT: 0, srcR: 400, srcB: 300, destL: 0, destT: 0, destR: 400, destB: 300, hSource: 0, sourceTitle: "", sourceExe: "", sourceClass: ""}

viewer := ThumbnailViewer(region1)  ; Start with no source - will prompt to select
