Option Explicit

' =========================================================
'  ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ
' =========================================================
Dim g_jsonPath
Dim g_durations
Dim g_checkboxStates
Dim Savedownpl

' =========================================================
'  ОСНОВНЫЕ ФУНКЦИИ РЕДАКТОРА ПЛЕЙЛИСТОВ
' =========================================================

Sub EditPlaylist(fieldId)
    On Error Resume Next

    Dim el, inputEl, playlistUrl

    Set el = Document.getElementById(fieldId)
    If el Is Nothing Then Exit Sub

    Set inputEl = el.getElementsByTagName("input")(0)
    If inputEl Is Nothing Then Exit Sub

    playlistUrl = Trim(inputEl.value)
    If playlistUrl = "" Then Exit Sub
	
    ' --- Проверяем существует ли уже HTA для этого fieldId ---
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim htaPath: htaPath = "config\playlist\playlist_" & fieldId & ".hta"
    
    If fso.FileExists(htaPath) Then
        ' HTA уже существует - открываем его
        Dim shell: Set shell = CreateObject("WScript.Shell")
        shell.Run """" & htaPath & """", 1, False
        Exit Sub
    End If
	
    ' --- MIX / RADIO обнаружен ---
    If IsGeneratedList(playlistUrl) Then
        MsgBox "⚠️ Это автогенерируемый список (MIX / Radio)." & vbCrLf & _
               "Он не является обычным плейлистом, его содержимое непостоянно." & vbCrLf & _
               "Редактирование недоступно." & vbCrLf & vbCrLf & _
               "Для загрузки используйте кнопку «Скачать плейлист».", _
               vbExclamation, "MIX / Radio"
        Exit Sub
    End If

    ' --- Универсальная проверка (НЕ блокируем) ---
    If Not LooksLikePlaylist(playlistUrl) Then
        MsgBox "⚠️ Возможно это не плейлист, но попробуем обработать.", _
               vbInformation, "Предупреждение"
    End If

    ' --- нормальный плейлист ---
    Call StartPlaylistPicker(fieldId, playlistUrl)
End Sub

Function IsGeneratedList(url)
    If InStr(url, "list=rd") > 0 Or InStr(url, "start_radio=1") > 0 Then
        IsGeneratedList = True
    Else
        IsGeneratedList = False
    End If
End Function

Function LooksLikePlaylist(url)
    LooksLikePlaylist = _
        (InStr(url, "list=") > 0) Or _
        (InStr(url, "playlist") > 0) Or _
        (InStr(url, "index=") > 0) Or _
        (InStr(url, "collection") > 0) Or _
        (InStr(url, "/set/") > 0) Or _
        (InStr(url, "playlists") > 0) Or _
        (InStr(url, "album") > 0) Or _		
        (InStr(url, "collections") > 0) Or _
        (InStr(url, "set=") > 0) Or _
        (InStr(url, "/sets/") > 0) Or _
        (InStr(url, "/folder") > 0) Or _
        (InStr(url, "folder=") > 0) Or _
        (InStr(url, "/series") > 0) Or _
        (InStr(url, "series=") > 0)
End Function

Const BIN_FOLDER    = "bin"
Const CACHE_FOLDER  = "temp\cache"
Const PLAYLIST_DIR  = "config\playlist"
Const YTDLP_EXE     = "yt-dlp.exe"
Const TEMPLATE_HTA  = "src\utils\playlist.hta"

Sub StartPlaylistPicker(fieldId, playlistUrl)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim shell: Set shell = CreateObject("WScript.Shell")

    ' ---------- пути ----------
    Dim batPath, tmpTxtLog, cleanJson, htaPath
    batPath   = BIN_FOLDER & "\playlist_tmp_" & fieldId & ".bat"
    tmpTxtLog   = CACHE_FOLDER & "\playlist_tmp_" & fieldId & ".txt"
    cleanJson = PLAYLIST_DIR & "\playlist_" & fieldId & ".json"
    htaPath   = PLAYLIST_DIR & "\playlist_" & fieldId & ".hta"

    ' ---------- удалить старые ----------
    If fso.FileExists(batPath)   Then fso.DeleteFile batPath, True
    If fso.FileExists(tmpTxtLog)   Then fso.DeleteFile tmpTxtLog, True
    If fso.FileExists(cleanJson) Then fso.DeleteFile cleanJson, True
    If fso.FileExists(htaPath)   Then fso.DeleteFile htaPath, True

    ' ---------- создать BAT ----------
    Dim bat: Set bat = fso.CreateTextFile(batPath, True, False)

    bat.WriteLine "@echo off"
    bat.WriteLine "chcp 65001 >nul"
    bat.WriteLine "setlocal ENABLEDELAYEDEXPANSION"
    bat.WriteLine ""
    bat.WriteLine "REM === параметры ==="
    bat.WriteLine "set ""URL=" & playlistUrl & """"
    bat.WriteLine "set ""OUT=" & "..\" & CACHE_FOLDER & "\playlist_tmp_" & fieldId & ".txt"""
    bat.WriteLine ""
    bat.WriteLine "echo ===Parsing playlist data==="
    bat.WriteLine "echo URL: !URL!"
    bat.WriteLine "echo ==========================="
    bat.WriteLine ""
    bat.WriteLine "pushd %~dp0"
    bat.WriteLine ""
    bat.WriteLine "echo Executing yt-dlp..."
    bat.WriteLine "yt-dlp.exe --flat-playlist --print ""%%(playlist_title)s"" --print ""%%(playlist_index)s<TAB>%%(title)s<TAB>%%(url)s<TAB>%%(duration_string)s<TAB>"" ""!URL!"" > ""!OUT!"" 2>&1"

    bat.Close

    ' ---------- запуск батника ----------
    shell.Run "cmd /c """ & batPath & """", 1, True
    fso.DeleteFile batPath, True
	
    ' ---------- конвертировать TXT в JSON ----------
    On Error Resume Next
    Dim playlistData: Set playlistData = ParsePlaylistTxt(tmpTxtLog)
    Dim conversionError: conversionError = ""

    If Err.Number <> 0 Then
        conversionError = "Ошибка парсинга: " & Err.Description
    Else
        Call SavePlaylistJson(cleanJson, playlistData, playlistUrl)
        If Err.Number <> 0 Then
            conversionError = "Ошибка сохранения JSON: " & Err.Description
        End If
    End If

    ' В ЛЮБОМ СЛУЧАЕ удаляем временные файлы
    fso.DeleteFile batPath, True
    If fso.FileExists(tmpTxtLog) Then fso.DeleteFile tmpTxtLog, True

    ' Если была ошибка - создаем JSON с ошибкой
    If conversionError <> "" Then
        Call SaveErrorJson(cleanJson, conversionError, playlistUrl)
    End If

    On Error GoTo 0

    ' ---------- копировать HTA ----------
    fso.CopyFile TEMPLATE_HTA, htaPath, True

    ' ---------- открыть ----------
    shell.Run """" & htaPath & """", 1, False
End Sub

Sub SaveErrorJson(jsonPath, errorMsg, sourceUrl)
    Dim fso, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(jsonPath, True, False)
    
    file.WriteLine "{"
    file.WriteLine "  ""playlist_title"": ""Ошибка создания плейлиста"","
    file.WriteLine "  ""source_url"": """ & EscapeJson(sourceUrl) & ""","
    file.WriteLine "  ""error"": """ & EscapeJson(errorMsg) & ""","
    file.WriteLine "  ""items"": []"
    file.WriteLine "}"
    file.Close
End Sub

Function ParsePlaylistTxt(txtPath)
    Dim fso, file, lines, i, playlistTitle, items(), itemCount
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(txtPath) Then
        Set ParsePlaylistTxt = CreateObject("Scripting.Dictionary")
        ParsePlaylistTxt("playlist_title") = "Файл не найден"
        ParsePlaylistTxt("items") = Array()
        ParsePlaylistTxt("item_count") = 0
        Exit Function
    End If
    
    Set file = fso.OpenTextFile(txtPath, 1, False)
    
    Dim content: content = file.ReadAll()
    file.Close()
    
    lines = Split(content, vbLf)
    
    ' Первая строка (нечетная) - заголовок плейлиста
    If UBound(lines) >= 0 Then 
        playlistTitle = Trim(lines(0))
    Else
        playlistTitle = "Без названия"
    End If
    
    ' Парсим ЧЕТНЫЕ строки (индекс 1, 3, 5...) - это видео
    itemCount = 0
    ReDim items(100)
    
    For i = 1 To UBound(lines) Step 2
        If i <= UBound(lines) And Trim(lines(i)) <> "" Then
            Set items(itemCount) = ParsePlaylistLine(lines(i))
            itemCount = itemCount + 1
        End If
    Next
    
    If itemCount > 0 Then
        ReDim Preserve items(itemCount - 1)
    Else
        ReDim items(0)
    End If
    
    Dim result: Set result = CreateObject("Scripting.Dictionary")
    result("playlist_title") = playlistTitle
    result("items") = items
    result("item_count") = itemCount
    
    Set ParsePlaylistTxt = result
End Function

Function ParsePlaylistLine(line)
    Dim parts, item
    Set item = CreateObject("Scripting.Dictionary")
    
    parts = Split(line, "<TAB>")
    
    ' Заполняем поля
    If UBound(parts) >= 0 Then item("index") = Trim(parts(0))
    If UBound(parts) >= 1 Then item("title") = Trim(parts(1))
    If UBound(parts) >= 2 Then item("url") = Trim(parts(2))
    If UBound(parts) >= 3 Then item("duration") = Trim(parts(3))
    item("selected") = True ' По умолчанию ВСЕ выбраны
    
    Set ParsePlaylistLine = item
End Function

Sub SavePlaylistJson(jsonPath, playlistData, sourceUrl)
    Dim fso, file, i, item
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(jsonPath, True, False)
    
    file.WriteLine "{"
    file.WriteLine "  ""playlist_title"": """ & EscapeJson(playlistData("playlist_title")) & ""","
    file.WriteLine "  ""source_url"": """ & EscapeJson(sourceUrl) & ""","
    file.WriteLine "  ""items"": ["
    
    Dim items: items = playlistData("items")
    For i = 0 To UBound(items)
        Set item = items(i)
        file.Write "    {""index"": """ & EscapeJson(item("index")) & """, ""title"": """ & EscapeJson(item("title")) & """, ""duration"": """ & EscapeJson(item("duration")) & """, ""url"": """ & EscapeJson(item("url")) & """, ""selected"": " & LCase(item("selected")) & "}"
        If i < UBound(items) Then
            file.WriteLine ","
        Else
            file.WriteLine ""
        End If
    Next
    
    file.WriteLine "  ]"
    file.WriteLine "}"
    file.Close
End Sub

Function EscapeJson(text)
    If IsNull(text) Then
        EscapeJson = ""
    Else
        EscapeJson = Replace(Replace(Replace(text, "\", "\\"), """", "\"""), vbCrLf, "\n")
    End If
End Function

' =========================================================
'  ИНИЦИАЛИЗАЦИЯ ОКНА ПЛЕЙЛИСТА
' =========================================================

Sub InitPlaylistWindow()
    On Error Resume Next
    
    If InStr(LCase(Document.location.pathname), "playlist_") = 0 Then Exit Sub

    ' Инициализируем глобальные переменные
    Set g_durations = CreateObject("Scripting.Dictionary")
    Set g_checkboxStates = CreateObject("Scripting.Dictionary")
    
    g_jsonPath = DetectJsonPath()

    If g_jsonPath = "" Then
        MsgBox "Не найден JSON плейлиста", vbCritical
        Exit Sub
    End If

    LoadPlaylist
	  
	DisplayPlaylistSettings()
	InitializePlaylistAuth() 


End Sub

Function DetectJsonPath()
    Dim fso, htaFullPath, folder, fname, id

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' полный путь к текущему HTA-файлу
    htaFullPath = Replace(Document.location.pathname, "/", "\")

    If InStr(htaFullPath, ":\") = 0 Then
        ' убираем возможный префикс file:///
        htaFullPath = Mid(htaFullPath, InStr(htaFullPath, "\"))
    End If

    folder = fso.GetParentFolderName(htaFullPath)
    fname = fso.GetFileName(htaFullPath)

    id = Replace(fname, "playlist_", "")
    id = Replace(id, ".hta", "")

    DetectJsonPath = folder & "\playlist_" & id & ".json"
End Function

' =========================================================
'  ЗАГРУЗКА И ОТОБРАЖЕНИЕ ПЛЕЙЛИСТА
' =========================================================

Sub LoadPlaylist()
    On Error Resume Next
    
    Dim json, playlistTitle, sourceUrl, pos, block, idx, title, duration, url, selected
    
    json = ReadFile(g_jsonPath)
    If json = "" Then 
        MsgBox "Не удалось загрузить JSON файл: " & g_jsonPath
        Exit Sub
    End If
    
    playlistTitle = ExtractValue(json, "playlist_title")
    sourceUrl = ExtractValue(json, "source_url")
    
    Document.getElementById("playlistTitle").innerText = playlistTitle
    Document.getElementById("sourceUrl").innerHTML = "<a href=""" & sourceUrl & """ target=""_blank"" style=""color: #6cb6ff; text-decoration: underline;"">" & sourceUrl & "</a>"
    
    ' Очищаем контейнер
    Dim container
    Set container = Document.getElementById("playlistContainer")
    container.innerHTML = ""
    
    ' Очищаем глобальные массивы
    g_durations.RemoveAll
    g_checkboxStates.RemoveAll
    
    ' Парсим и отображаем элементы
    pos = InStr(json, """items""")
    If pos > 0 Then pos = InStr(pos, json, "[")
    
    If pos > 0 Then
        Do
            block = NextJsonObject(json, pos)
            If block = "" Then Exit Do

            idx = ExtractValue(block, "index")
            title = ExtractValue(block, "title")
            duration = ExtractValue(block, "duration")
            url = ExtractValue(block, "url")
            selected = ExtractValue(block, "selected")

            ' Сохраняем в глобальные словари
            g_durations(idx) = duration
            g_checkboxStates(idx) = (LCase(selected) = "true")

            AddRow container, idx, title, duration, url, selected
        Loop
    End If
    
    ' Обновляем общий чекбокс после загрузки
    UpdateSelectAllCheckbox
    
    ' Обновляем общее время
    UpdateTotalTime
End Sub

Sub AddRow(container, idx, title, duration, url, selected)
    Dim chk, html

    If LCase(selected) = "true" Then
        chk = "checked"
    Else
        chk = ""
    End If

    html = ""
    html = html & "<table class='playlistTable'>"
    html = html & "<tr>"
    html = html & "<td class='checkboxCell'>" & _
                  "<input type='checkbox' class='pl-check' data-index='" & idx & "' " & chk & _
                  " onchange='ItemCheckboxChanged'></td>" ' УБИРАЕМ (this)
    html = html & "<td class='indexCell'>" & idx & "</td>"
    html = html & "<td class='titleCell' title='" & Replace(title, "'", "&#39;") & "'>" & title & "</td>"
    html = html & "<td class='timeCell'>" & duration & "</td>"
    html = html & "</tr></table>"

    container.insertAdjacentHTML "beforeEnd", html
End Sub

' =========================================================
'  УПРАВЛЕНИЕ СОСТОЯНИЯМИ ЧЕКБОКСОВ
' =========================================================

Sub ToggleAllItems()
    On Error Resume Next
    
    Dim master, container, inputs, i, idx
    Set master = Document.getElementById("selectAllBox")
    Set container = Document.getElementById("playlistContainer")
    Set inputs = container.getElementsByTagName("input")
    
    For i = 0 To inputs.length - 1
        If inputs(i).className = "pl-check" Then
            idx = inputs(i).getAttribute("data-index")
            inputs(i).Checked = master.Checked
            ' ОБНОВЛЯЕМ СОСТОЯНИЕ В ПАМЯТИ
            g_checkboxStates(idx) = master.Checked
        End If
    Next
    
    ' Обновляем общее время
    UpdateTotalTime
End Sub

Sub ItemCheckboxChanged()
    On Error Resume Next
    
    ' Получаем элемент из события
    Dim cb
    Set cb = window.event.srcElement
    
    If cb Is Nothing Then
        Exit Sub
    End If
    
    Dim idx
    idx = cb.getAttribute("data-index")
    If idx = "" Then Exit Sub
    
    ' Обновляем состояние в памяти
    g_checkboxStates(idx) = cb.Checked
    
    UpdateSelectAllCheckbox
    
End Sub

Sub UpdateSelectAllCheckbox()
    On Error Resume Next
    
    Dim container, inputs, i, allChecked
    Set container = Document.getElementById("playlistContainer")
    Set inputs = container.getElementsByTagName("input")
    
    If inputs.length = 0 Then Exit Sub
    
    allChecked = True
    
    For i = 0 To inputs.length - 1
        If inputs(i).className = "pl-check" And Not inputs(i).Checked Then
            allChecked = False
            Exit For
        End If
    Next
    
    Document.getElementById("selectAllBox").Checked = allChecked
End Sub

Sub UpdateJsonSelected(index, state)
    On Error Resume Next
    
    Dim json, oldObj, newObj, startPos, endPos, objStart, objEnd
    
    json = ReadFile(g_jsonPath)
    If json = "" Then Exit Sub
    
    ' Ищем объект с нужным индексом
    startPos = InStr(json, """index"": """ & index & """")
    If startPos = 0 Then Exit Sub
    
    ' Находим начало и конец объекта
    objStart = startPos
    Do While objStart > 1
        If Mid(json, objStart, 1) = "{" Then Exit Do
        objStart = objStart - 1
    Loop
    
    objEnd = objStart
    Dim bracketCount: bracketCount = 0
    Do While objEnd <= Len(json)
        If Mid(json, objEnd, 1) = "{" Then bracketCount = bracketCount + 1
        If Mid(json, objEnd, 1) = "}" Then 
            bracketCount = bracketCount - 1
            If bracketCount = 0 Then Exit Do
        End If
        objEnd = objEnd + 1
    Loop
    
    If objEnd > Len(json) Then Exit Sub
    
    oldObj = Mid(json, objStart, objEnd - objStart + 1)
    
    ' Обновляем selected
    If InStr(oldObj, """selected"":") > 0 Then
        newObj = Replace(oldObj, """selected"": true", """selected"": " & LCase(state))
        newObj = Replace(newObj, """selected"": false", """selected"": " & LCase(state))
    Else
        ' Добавляем selected если его нет
        newObj = Left(oldObj, Len(oldObj) - 1) & ", ""selected"": " & LCase(state) & "}"
    End If
    
    ' Заменяем в JSON
    json = Replace(json, oldObj, newObj)
    WriteFile g_jsonPath, json
End Sub

' =========================================================
'  КНОПКИ УПРАВЛЕНИЯ
' =========================================================
Sub Savedownplaylist()
    On Error Resume Next
    Savedownpl = "true"
    SaveCurrentState
End Sub

Sub SaveCurrentState()
    On Error Resume Next
    
    If g_jsonPath = "" Then Exit Sub
    
    ' Читаем текущий JSON чтобы взять структуру
    Dim json, fso, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(g_jsonPath, 1)
    json = file.ReadAll()
    file.Close()
    
    ' Получаем состояния всех чекбоксов
    Dim container, inputs, i, idx, isChecked
    Set container = Document.getElementById("playlistContainer")
    Set inputs = container.getElementsByTagName("input")
    
    ' Проходим по всем чекбоксам и обновляем JSON
    For i = 0 To inputs.length - 1
        If inputs(i).className = "pl-check" Then
            idx = inputs(i).getAttribute("data-index")
            isChecked = inputs(i).Checked
            
            ' Ищем и заменяем selected для этого индекса
            json = UpdateSelectedInJson(json, idx, isChecked)
        End If
    Next
    
    ' Сохраняем обновленный JSON
    Set file = fso.CreateTextFile(g_jsonPath, True)
    file.Write json
    file.Close()
UpdateTotalTime
If Savedownpl = "true" Then
        Savedownpl = ""
        downplaylist()
    End If

End Sub

Sub CloseWindow()
    On Error Resume Next
    window.close
End Sub

Function UpdateSelectedInJson(json, index, isChecked)
    Dim pos, searchStr, selectedPos, valueStart, valueEnd, oldValue, newValue
    
    ' Ищем объект с нужным индексом
    searchStr = """index"": """ & index & """"
    pos = InStr(json, searchStr)
    If pos = 0 Then Exit Function
    
    ' Ищем selected после этого индекса (в пределах того же объекта)
    selectedPos = InStr(pos, json, """selected"":")
    If selectedPos = 0 Then Exit Function
    
    ' Находим начало значения selected
    valueStart = InStr(selectedPos, json, ":") + 1
    ' Пропускаем пробелы
    Do While Mid(json, valueStart, 1) = " " And valueStart < Len(json)
        valueStart = valueStart + 1
    Loop
    
    ' Находим конец значения selected (до запятой или закрывающей скобки)
    valueEnd = valueStart
    Do While valueEnd <= Len(json)
        Dim ch
        ch = Mid(json, valueEnd, 1)
        If ch = "," Or ch = "}" Then Exit Do
        valueEnd = valueEnd + 1
    Loop
    
    ' Извлекаем старое значение
    oldValue = Mid(json, valueStart, valueEnd - valueStart)
    oldValue = Trim(oldValue)
    
    ' Определяем новое значение
    If isChecked Then
        newValue = "true"
    Else
        newValue = "false"
    End If
    
    ' Заменяем в JSON
    UpdateSelectedInJson = Left(json, valueStart - 1) & newValue & Mid(json, valueEnd)
End Function

Sub RestoreFromJson()
    On Error Resume Next
    ' Просто перезагружаем плейлист (берет состояния из JSON)
    LoadPlaylist
End Sub

' =========================================================
'  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' =========================================================

Function ReadFile(path)
    On Error Resume Next
    Dim fso, tf
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tf = fso.OpenTextFile(path, 1)
    ReadFile = tf.ReadAll
    tf.Close
End Function

Sub WriteFile(path, content)
    On Error Resume Next
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(path, True, False)
    f.Write content
    f.Close
End Sub

Function ExtractValue(txt, key)
    Dim p, i, ch, result, inString

    p = InStr(txt, """" & key & """")
    If p = 0 Then 
        ExtractValue = ""
        Exit Function
    End If

    p = InStr(p, txt, ":")
    If p = 0 Then 
        ExtractValue = ""
        Exit Function
    End If
    
    p = p + 1
    
    ' Пропускаем пробелы
    Do While p <= Len(txt) And (Mid(txt, p, 1) = " " Or Mid(txt, p, 1) = vbTab)
        p = p + 1
    Loop
    
    If p > Len(txt) Then 
        ExtractValue = ""
        Exit Function
    End If
    
    ' Обрабатываем разные типы значений
    If Mid(txt, p, 1) = """" Then
        ' Строковое значение в кавычках (index, title, duration, url)
        p = p + 1
        result = ""
        For i = p To Len(txt)
            ch = Mid(txt, i, 1)
            If ch = """" And Mid(txt, i - 1, 1) <> "\" Then Exit For
            result = result & ch
        Next
    Else
        ' Булево значение без кавычек (selected: true/false)
        result = ""
        For i = p To Len(txt)
            ch = Mid(txt, i, 1)
            If ch = "," Or ch = "}" Or ch = " " Or ch = vbCr Or ch = vbLf Then Exit For
            result = result & ch
        Next
        result = Trim(result)
    End If

    ExtractValue = result
End Function

Function NextJsonObject(ByRef txt, ByRef pos)
    Dim s, e, d, i, ch

    s = InStr(pos, txt, "{")
    If s = 0 Then Exit Function

    d = 0
    For i = s To Len(txt)
        ch = Mid(txt, i, 1)
        If ch = "{" Then d = d + 1
        If ch = "}" Then d = d - 1
        If d = 0 Then
            e = i
            Exit For
        End If
    Next

    NextJsonObject = Mid(txt, s, e - s + 1)
    pos = e + 1
End Function

' =========================================================
'  ФУНКЦИИ ДЛЯ РАБОТЫ СО ВРЕМЕНЕМ
' =========================================================

Sub UpdateTotalTime()
    On Error Resume Next
    Dim totalTimeElement
    Set totalTimeElement = Document.getElementById("totalTime")
    If Not totalTimeElement Is Nothing Then
        totalTimeElement.innerText = "Время: " & CalculateTotalTime()
    End If
End Sub

Function CalculateTotalTime()
    On Error Resume Next
    
    Dim totalSeconds, key, duration, isChecked
    
    totalSeconds = 0
    
    ' Проходим по всем элементам в памяти
    For Each key In g_durations.Keys
        duration = g_durations(key)
        isChecked = g_checkboxStates(key)
        
        If isChecked And duration <> "" Then
            totalSeconds = totalSeconds + TimeStringToSeconds(duration)
        End If
    Next
    
    CalculateTotalTime = FormatTotalTime(totalSeconds)
End Function

Function TimeStringToSeconds(timeStr)
    On Error Resume Next
    
    Dim parts, hours, minutes, seconds
    
    timeStr = Trim(timeStr)
    If timeStr = "" Then
        TimeStringToSeconds = 0
        Exit Function
    End If
    
    parts = Split(timeStr, ":")
    
    If UBound(parts) = 2 Then
        ' Формат H:MM:SS
        hours = CInt(parts(0))
        minutes = CInt(parts(1))
        seconds = CInt(parts(2))
    ElseIf UBound(parts) = 1 Then
        ' Формат MM:SS
        hours = 0
        minutes = CInt(parts(0))
        seconds = CInt(parts(1))
    Else
        ' Неизвестный формат
        TimeStringToSeconds = 0
        Exit Function
    End If
    
    TimeStringToSeconds = (hours * 3600) + (minutes * 60) + seconds
End Function

Function FormatTotalTime(totalSeconds)
    On Error Resume Next
    
    Dim hours, minutes
    
    hours = totalSeconds \ 3600
    minutes = (totalSeconds Mod 3600) \ 60
    
    ' Форматируем как "H ч. MM мин." (секунды не показываем)
    If hours > 0 Then
        FormatTotalTime = hours & " ч. " & Right("0" & minutes, 2) & " мин."
    Else
        FormatTotalTime = minutes & " мин."
    End If
End Function

' ==================== НАСТРОЙКИ ДЛЯ HTA ПЛЕЙЛИСТОВ ====================

' ★★★ ЗАГРУЗКА НАСТРОЕК ПЛЕЙЛИСТА (ВОЗВРАЩАЕТ СЛОВАРЬ) ★★★
Function LoadPlaylistSettingsForPlaylist()
    On Error Resume Next
    Dim fso, settingsPath, settings, savePath, quality, format, subsValue, embeddedFlag, detectedBrowser
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Файл настроек в config\playlist\ относительно корня приложения
    settingsPath = "config\playlist\playlist_settings.txt"
    
    ' Создаем словарь для результатов
    Dim resultDict
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    ' Устанавливаем значения по умолчанию
    resultDict("savePath") = ""
    resultDict("defaultQuality") = "360"
    resultDict("defaultFormat") = "mp4"
    resultDict("subtitles") = "none"
    resultDict("embeddedSubs") = "false"
    resultDict("detectedBrowser") = ""
    resultDict("proxy") = ""
    
    ' Читаем настройки из файла
    If fso.FileExists(settingsPath) Then
        Dim settingsFile, settingsArray
        Set settingsFile = fso.OpenTextFile(settingsPath, 1)
        settings = settingsFile.ReadAll
        settingsFile.Close
        
        settingsArray = Split(settings, "|")
        
        ' Заполняем словарь значениями из файла
        If UBound(settingsArray) >= 0 Then 
            resultDict("savePath") = settingsArray(0)
        End If
        If UBound(settingsArray) >= 1 Then 
            resultDict("defaultQuality") = settingsArray(1)
        End If
        If UBound(settingsArray) >= 2 Then 
            resultDict("defaultFormat") = settingsArray(2)
        End If
        If UBound(settingsArray) >= 3 Then 
            resultDict("proxy") = settingsArray(3)
        End If
        If UBound(settingsArray) >= 4 Then 
            resultDict("subtitles") = settingsArray(4)
        End If
        If UBound(settingsArray) >= 5 Then 
            resultDict("embeddedSubs") = settingsArray(5)
        End If
        ' ★★★ ВАЖНО: detectedBrowser в позиции 6 ★★★
        If UBound(settingsArray) >= 6 Then 
            resultDict("detectedBrowser") = Trim(settingsArray(6))
        End If
    End If
    
    ' Возвращаем словарь
    Set LoadPlaylistSettingsForPlaylist = resultDict
End Function

' ★★★ ОТОБРАЖЕНИЕ НАСТРОЕК ПЛЕЙЛИСТА В ИНТЕРФЕЙСЕ ★★★
Sub DisplayPlaylistSettings()
    On Error Resume Next
    
    Dim settings
    Set settings = LoadPlaylistSettingsForPlaylist()
    If settings Is Nothing Then Exit Sub
    
    Dim savePath, quality, format, subsValue, embeddedFlag, detectedBrowser
    
    savePath = settings("savePath")
    quality = settings("defaultQuality")
    format = settings("defaultFormat")
    subsValue = settings("subtitles")
    embeddedFlag = settings("embeddedSubs")
    detectedBrowser = settings("detectedBrowser")
    
    ' Формируем отображение
    Dim subtitlesText
    If format = "mp3" Or subsValue = "none" Then
        subtitlesText = "Без субтитров"
    Else
        If LCase(embeddedFlag) = "true" Then
            subtitlesText = "Субтитры: " & subsValue & " (встроенные)"
        Else
            subtitlesText = "Субтитры: " & subsValue & " (внешние)"
        End If
    End If
    
    Dim qualityFormat
    If format = "mp3" Then
        qualityFormat = "🎵 " & format
    Else
        qualityFormat = "📺 " & quality & "p 🎬 " & format & " 📝 " & subtitlesText
    End If

    Dim html
    html = "<div style='display: flex; justify-content: space-between; align-items: center; line-height: 1.5;'>"
    html = html & "<div>" & "Текущие настройки:" & "&nbsp;&nbsp;"
    html = html & qualityFormat & " 📁 " & savePath & "&nbsp;&nbsp;"
    
 ' БЛОК С ЧЕКБОКСОМ АВТОРИЗАЦИИ 
If detectedBrowser <> "" Then
    ' Браузер найден - добавляем чекбокс
    html = html & " <label title='Использовать авторизацию через " & detectedBrowser & "' style='cursor:pointer;'>"
    html = html & "<input type='checkbox' id='usePlaylistAuth' onclick='VBScript:UpdatePlaylistAuthStatus()' style='vertical-align:middle;'>"
    html = html & "<span id='playlistAuthStatus'>" & detectedBrowser & "</span>"
Else
    ' Браузер не найден - только текст
    html = html & " Не авторизован 🔒"
End If

html = html & "</div>"
    
    Document.getElementById("playlistSettings").innerHTML = html
End Sub

' ★★★ ИНИЦИАЛИЗАЦИЯ АВТОРИЗАЦИИ ДЛЯ РЕДАКТОРА ПЛЕЙЛИСТОВ ★★★
Sub InitializePlaylistAuth()
    On Error Resume Next
    
    ' Инициализируем только если это окно плейлиста
    If InStr(LCase(Document.location.pathname), "playlist_") = 0 Then Exit Sub
    
    Dim settings, authCheckbox, statusEl
    
    ' Загружаем настройки плейлиста
    Set settings = LoadPlaylistSettingsForPlaylist()
    If settings Is Nothing Then Exit Sub
    
    Set authCheckbox = Document.getElementById("usePlaylistAuth")
    Set statusEl = Document.getElementById("playlistAuthStatus")
    
    If Not authCheckbox Is Nothing And Not statusEl Is Nothing Then
        Dim browserName
        browserName = settings("detectedBrowser")
        
        If browserName <> "" And browserName <> "Не авторизован" Then
            ' Браузер найден - чекбокс включен по умолчанию
            authCheckbox.Checked = False
            statusEl.innerText = browserName & " 🔐 выкл."
            statusEl.style.color = "#ff6b6b"  ' красный
        Else
            ' Браузер не найден - чекбокс выключен
            authCheckbox.Checked = False
            authCheckbox.disabled = True  ' делаем неактивным
            statusEl.innerText = "Авторизация недоступна"
            statusEl.style.color = "#888"  ' серый
        End If
    End If
End Sub

'' ★★★ ОБНОВЛЕНИЕ СТАТУСА ПРИ ИЗМЕНЕНИИ ЧЕКБОКСА ★★★
Sub UpdatePlaylistAuthStatus()
    On Error Resume Next
    
    Dim authCheckbox, statusEl, settings
    
    Set authCheckbox = Document.getElementById("usePlaylistAuth")
    Set statusEl = Document.getElementById("playlistAuthStatus")
    
    If authCheckbox Is Nothing Or statusEl Is Nothing Then Exit Sub
    
    ' Если чекбокс неактивен (браузер не найден) - ничего не делаем
    If authCheckbox.disabled Then Exit Sub
    
    ' Загружаем настройки для получения имени браузера
    Set settings = LoadPlaylistSettingsForPlaylist()
    If settings Is Nothing Then Exit Sub
    
    Dim browserName
    browserName = settings("detectedBrowser")
    
    If browserName <> "" And browserName <> "Не авторизован" Then
        If authCheckbox.Checked Then
            statusEl.innerText = browserName & " 🔓 вкл.  "
            statusEl.style.color = "#4CAF50"  ' зеленый
        Else
            statusEl.innerText = browserName & " 🔐 выкл."
            statusEl.style.color = "#ff6b6b"  ' красный
        End If
    Else
        statusEl.innerText = "Авторизация недоступна"
        statusEl.style.color = "#888"
    End If
End Sub