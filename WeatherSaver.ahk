#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent()

; ==============================================================================
; WeatherSaver.ahk
; Background weather data fetcher - saves WeatherAPI.com JSON to a local file
; so LiveView (or other tools) can read weather data from a file.
; ==============================================================================

global WS_ConfigFile := A_ScriptDir "\WeatherSaver.ini"
global WS_ApiKeyFile := A_ScriptDir "\api_keys.ini"
global WS_ApiKey := ""
global WS_Location := "New York"
global WS_Unit := "fahrenheit"
global WS_IntervalMinutes := 15
global WS_OutputFile := A_ScriptDir "\weather_data.json"
global WS_IsFetching := false
global WS_SearchResults := []

WS_Log("Starting WeatherSaver")

WS_LoadConfig()
WS_Log("Config loaded")

WS_SetupTray()
WS_Log("Tray setup done")

; Initial fetch via a one-off function so it doesn't conflict with the recurring timer
SetTimer(() => WS_DoFetch(), -100)
WS_Log("SetTimer called for initial WS_DoFetch")

if (WS_IntervalMinutes > 0) {
    SetTimer(WS_DoFetch, WS_IntervalMinutes * 60000)
    WS_Log("Set recurring timer for " WS_IntervalMinutes " minutes")
}

WS_Log("End of auto-execute thread")

WS_Log(msg) {
    FileAppend(FormatTime(, "HH:mm:ss") " - " msg "`n", A_ScriptDir "\weathersaver_debug.log")
}

; ---- Config Load/Save ----

WS_LoadConfig() {
    if FileExist(WS_ConfigFile) {
        try global WS_ApiKey := IniRead(WS_ConfigFile, "Settings", "APIKey", "")
        try global WS_Location := IniRead(WS_ConfigFile, "Settings", "Location", "New York")
        try global WS_Unit := IniRead(WS_ConfigFile, "Settings", "Unit", "fahrenheit")
        try {
            v := IniRead(WS_ConfigFile, "Settings", "IntervalMinutes", "15")
            global WS_IntervalMinutes := IsNumber(v) ? Integer(v) : 15
        }
        try global WS_OutputFile := IniRead(WS_ConfigFile, "Settings", "OutputFile", A_ScriptDir "\weather_data.json")
    }

    ; Fallback: read API key from api_keys.ini
    if (WS_ApiKey = "" && FileExist(WS_ApiKeyFile)) {
        try global WS_ApiKey := IniRead(WS_ApiKeyFile, "Weather", "APIKey", "")
        if (WS_ApiKey = "")
            try global WS_ApiKey := IniRead(WS_ApiKeyFile, "Weather", "WeatherAPIKey", "")
    }

    WS_SaveConfig()
}

WS_SaveConfig() {
    try {
        IniWrite(WS_ApiKey, WS_ConfigFile, "Settings", "APIKey")
        IniWrite(WS_Location, WS_ConfigFile, "Settings", "Location")
        IniWrite(WS_Unit, WS_ConfigFile, "Settings", "Unit")
        IniWrite(WS_IntervalMinutes, WS_ConfigFile, "Settings", "IntervalMinutes")
        IniWrite(WS_OutputFile, WS_ConfigFile, "Settings", "OutputFile")
    }
}

; ---- Tray Menu ----

WS_SetupTray() {
    A_IconTip := "WeatherSaver"
    m := A_TrayMenu
    m.Delete()
    m.Add("Fetch Now", WS_TrayFetch)
    m.Add("Open Output File", WS_TrayOpen)
    m.Add("Configure...", WS_TrayConfig)
    m.Add()
    m.Add("Exit", WS_TrayExit)
}

WS_TrayFetch(*) {
    WS_DoFetch()
}

WS_TrayOpen(*) {
    if FileExist(WS_OutputFile)
        Run('notepad.exe "' WS_OutputFile '"')
    else
        MsgBox("Output file does not exist yet:`n" WS_OutputFile, "WeatherSaver", "48")
}

WS_TrayConfig(*) {
    WS_ShowConfigGui()
}

WS_TrayExit(*) {
    ExitApp()
}

; ---- Fetch Weather ----

WS_DoFetch(*) {
    WS_Log("WS_DoFetch started")
    if WS_IsFetching {
        WS_Log("Already fetching, returning")
        return
    }
    global WS_IsFetching := true

    try {
        if (WS_ApiKey = "") {
            A_IconTip := "WeatherSaver: No API Key"
            WS_IsFetching := false
            WS_Log("No API Key")
            return
        }

        WS_Log("Encoding location...")
        encoded := WS_UriEncode(WS_Location)
        WS_Log("Location encoded: " encoded)
        
        url := "https://api.weatherapi.com/v1/current.json?key=" WS_ApiKey "&q=" encoded "&aqi=no"
        tempPath := WS_OutputFile ".tmp"
        
        WS_Log("Deleting old temp file...")
        try FileDelete(tempPath)

        WS_Log("Downloading from API...")
        ; Use AHK built-in Download command
        try {
            Download(url, tempPath)
            WS_Log("Download completed")
        } catch as err {
            WS_Log("Download threw error: " err.Message)
        }

        if FileExist(tempPath) && FileGetSize(tempPath) > 10 {
            try FileDelete(WS_OutputFile)
            try FileMove(tempPath, WS_OutputFile, 1)
            A_IconTip := "WeatherSaver: Updated " FormatTime(A_Now, "HH:mm:ss")
            WS_Log("Saved to " WS_OutputFile)
        } else {
            A_IconTip := "WeatherSaver: Download failed"
            try FileDelete(tempPath)
            WS_Log("Download failed or file empty")
        }
    } catch as err {
        A_IconTip := "WeatherSaver: " SubStr(err.Message, 1, 40)
        WS_Log("Outer error: " err.Message)
    }

    global WS_IsFetching := false
    WS_Log("WS_DoFetch finished")
}

; ---- Config GUI ----

WS_ShowConfigGui(*) {
    g := Gui("+AlwaysOnTop +ToolWindow", "WeatherSaver Settings")
    g.SetFont("s10")

    g.AddText("w320", "WeatherAPI.com API Key:")
    apiEdit := g.AddEdit("w320 Password", WS_ApiKey)

    g.AddText("w320", "City / Location:")
    locEdit := g.AddEdit("w240", WS_Location)
    searchBtn := g.AddButton("x+5 w75", "Search")

    resultsList := g.AddListBox("xm w320 r5 vResultsList", [])


    g.AddText("w320", "Output File Path:")
    outEdit := g.AddEdit("w240", WS_OutputFile)
    browseBtn := g.AddButton("x+5 w75", "Browse...")
    browseBtn.OnEvent("Click", (*) => WS_BrowseOutput(outEdit))

    g.AddText("xm w320", "Update Interval (minutes, 0 = manual only):")
    intEdit := g.AddEdit("w320 Number", WS_IntervalMinutes)

    saveBtn := g.AddButton("w155", "Save && Fetch")
    cancelBtn := g.AddButton("x+10 w155", "Cancel")

    saveBtn.OnEvent("Click", (*) => WS_ApplyConfig(g, apiEdit, locEdit, outEdit, intEdit))
    cancelBtn.OnEvent("Click", (*) => g.Destroy())

    searchBtn.OnEvent("Click", (*) => WS_SearchLocation(locEdit, resultsList, apiEdit.Value))
    resultsList.OnEvent("DoubleClick", (*) => WS_SelectLocation(locEdit, resultsList))

    g.Show()
}

WS_BrowseOutput(outEdit) {
    selected := FileSelect(16, WS_OutputFile != "" ? WS_OutputFile : A_ScriptDir, "Select Output File", "JSON (*.json)|Text (*.txt)|All (*.*)")
    if (selected != "")
        outEdit.Value := selected
}

WS_SearchLocation(locEdit, resultsList, currentApiKey) {
    query := locEdit.Value
    if query = ""
        return

    resultsList.Delete()
    resultsList.Add(["Searching..."])
    global WS_SearchResults := []

    if (currentApiKey = "") {
        resultsList.Delete()
        resultsList.Add(["API Key required for search"])
        return
    }

    url := "https://api.weatherapi.com/v1/search.json?key=" currentApiKey "&q=" WS_UriEncode(query)
    tempFile := A_Temp "\weathersaver_geocode.json"
    try FileDelete(tempFile)

    try {
        Download(url, tempFile)
    } catch {
        resultsList.Delete()
        resultsList.Add(["Search failed to connect"])
        return
    }

    try {
        response := FileRead(tempFile)
        resultsList.Delete()

        results := []
        pos := 1
        while RegExMatch(response, '\{[^{}]*"name":"([^"]+)"[^{}]*"region":"([^"]*)"[^{}]*"country":"([^"]+)"[^{}]*"lat":([\d.-]+)[^{}]*"lon":([\d.-]+)[^{}]*"url":"([^"]+)"[^{}]*\}', &m, pos) {
            name := m[1]
            region := m[2]
            country := m[3]
            lat := m[4]
            lon := m[5]
            urlVal := m[6]

            displayName := name
            if region != ""
                displayName .= ", " region
            displayName .= ", " country

            results.Push(displayName)
            WS_SearchResults.Push({name: urlVal != "" ? urlVal : lat "," lon, display: displayName})
            pos := m.Pos + m.Len
        }

        if results.Length > 0
            resultsList.Add(results)
        else
            resultsList.Add(["No results found"])
    } catch {
        resultsList.Delete()
        resultsList.Add(["Search parsing failed"])
    }
}

WS_SelectLocation(locEdit, resultsList) {
    idx := resultsList.Value
    if idx = 0 || idx > WS_SearchResults.Length
        return

    result := WS_SearchResults[idx]
    locEdit.Value := result.display
}

WS_ApplyConfig(g, apiEdit, locEdit, outEdit, intEdit) {
    global WS_ApiKey := apiEdit.Value
    global WS_Location := locEdit.Value
    global WS_OutputFile := outEdit.Value
    global WS_IntervalMinutes := Max(0, Integer(intEdit.Value))

    WS_SaveConfig()

    ; Sync API key back to api_keys.ini
    if FileExist(WS_ApiKeyFile)
        try IniWrite(WS_ApiKey, WS_ApiKeyFile, "Weather", "APIKey")

    g.Destroy()

    ; Reset timer
    SetTimer(WS_DoFetch, 0)
    if (WS_IntervalMinutes > 0)
        SetTimer(WS_DoFetch, WS_IntervalMinutes * 60000)

    WS_DoFetch()
}

; ---- Helpers ----

WS_UriEncode(str) {
    buf := Buffer(StrPut(str, "UTF-8"))
    StrPut(str, buf, "UTF-8")
    out := ""
    Loop buf.Size - 1 {
        b := NumGet(buf, A_Index - 1, "UChar")
        if (b >= 0x41 && b <= 0x5A) || (b >= 0x61 && b <= 0x7A) || (b >= 0x30 && b <= 0x39) || b = 0x2D || b = 0x2E || b = 0x5F || b = 0x7E
            out .= Chr(b)
        else
            out .= Format("%{:02X}", b)
    }
    return out
}
