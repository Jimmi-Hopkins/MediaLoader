' Модуль управления интерфейсом

' Инициализация окна
Sub Window_onLoad()
    On Error Resume Next
    Window.ResizeTo 1280, 1024
    
    Dim screenWidth, screenHeight
    screenWidth = Screen.AvailWidth
    screenHeight = Screen.AvailHeight    
    Dim windowLeft, windowTop
    windowLeft = (screenWidth - 1280) / 2
    windowTop = (screenHeight - 720) / 2
    
    If windowLeft < 0 Then windowLeft = 0
    If windowTop < 0 Then windowTop = 0
    
    Window.MoveTo windowLeft, windowTop
		
End Sub
Sub UpdateProxyButtonColor()
    On Error Resume Next

    Dim fso, settingsPath, txt, arr, proxy
    Set fso = CreateObject("Scripting.FileSystemObject")

    settingsPath = fso.BuildPath( _
        fso.GetParentFolderName(window.location.pathname), _
        "config\downloader_settings.txt" _
    )

    If Not fso.FileExists(settingsPath) Then Exit Sub

    Dim f
    Set f = fso.OpenTextFile(settingsPath, 1)
    txt = f.ReadAll
    f.Close

    arr = Split(txt, "|")

    ' proxy = 4-й параметр
    If UBound(arr) >= 3 Then
        proxy = Trim(arr(3))
    Else
        proxy = ""
    End If

    If proxy = "" Or InStr(proxy, "http://ip:port") > 0 Then
        Document.getElementById("proxyButton").style.color = "#ff4747"  ' КРАСНЫЙ
    Else
        Document.getElementById("proxyButton").style.color = "#3cff3c"  ' ЗЕЛЁНЫЙ
    End If
End Sub

Sub ShowProxySettings()
    Document.getElementById("proxyPopup").style.display = "block"
End Sub

Sub HideProxySettings()
    Document.getElementById("proxyPopup").style.display = "none"
    Call SaveSettings()          ' ✅ сохраняем настройки
    Call UpdateProxyButtonColor  ' ✅ обновляем цвет кнопки
End Sub
Sub InitProxyPlaceholder()
    Dim proxyField
    Set proxyField = Document.getElementById("proxy")

    If Trim(proxyField.value) = "" Then
        proxyField.value = "http://ip:port или http://логин:пароль@ip:port"
        proxyField.style.color = "#888888" ' серый placeholder
    End If
End Sub

' ------------------------------
' Деактивация при выборе mp3
' ------------------------------
Sub FormatSelectionChanged()
    On Error Resume Next
    Dim formatSelect, qualitySelect, subtitlesSelect, embeddedSubsCheckbox
    
    Set formatSelect = Document.getElementById("defaultFormat")
    Set qualitySelect = Document.getElementById("defaultQuality")
    Set subtitlesSelect = Document.getElementById("subtitles")
    Set embeddedSubsCheckbox = Document.getElementById("embeddedSubs")
    
    If formatSelect.value = "mp3" Then
        ' Делаем элементы неактивными для MP3
        qualitySelect.disabled = True
        subtitlesSelect.disabled = True
        embeddedSubsCheckbox.disabled = True
    Else
        ' Включаем элементы обратно для других форматов
        qualitySelect.disabled = False
        subtitlesSelect.disabled = False
        embeddedSubsCheckbox.disabled = False
      
    End If
End Sub
' ------------------------------
' Деактивация чекбокса при выборе без субтитров
' ------------------------------
Sub SubtitlesChanged()
    On Error Resume Next
    Dim subtitlesSelect, embeddedSubsCheckbox
    
    Set subtitlesSelect = Document.getElementById("subtitles")
    Set embeddedSubsCheckbox = Document.getElementById("embeddedSubs")
    
    ' Если выбрано "нет субтитров" - деактивируем чекбокс
    If subtitlesSelect.value = "none" Then
        embeddedSubsCheckbox.disabled = True
       
    Else
        embeddedSubsCheckbox.disabled = False
      
    End If

End Sub

' Обновление статуса авторизации при изменении чекбокса
Sub UpdateAuthStatus()
    On Error Resume Next
    Dim statusEl, authCheckbox
    Set statusEl = Document.getElementById("authBrowserStatus")
    Set authCheckbox = Document.getElementById("useBrowserAuth")
    
    If detectedBrowser <> "" Then
        If Not authCheckbox Is Nothing And authCheckbox.Checked Then
            statusEl.innerText =  detectedBrowser & " вкл "
            statusEl.style.color = "lime"
        Else
            statusEl.innerText = detectedBrowser & " выкл"
            statusEl.style.color = "red"
        End If
    End If
End Sub
' Генерация полей для ссылок
'Sub GenerateUrlFields()
 '   On Error Resume Next
  '  Dim container, i, html
   ' Set container = Document.getElementById("urlFieldsContainer")
    
'html = ""
'For i = 1 To 5
 '   html = html & "<div class=""url-row"">" & _
 '       "<input type=""text"" id=""url" & i & """ placeholder=""Ссылка на видео"">" & _
 '       "<button onclick=""DownloadVideo 'url" & i & "'"">Скачать по умолчанию</button>" & _
 '       "<div class=""quality-buttons"">" & _
 '       "<button class=""max"" onclick=""DownloadVideoQuality 'url" & i & "','max'"">ТОП</button>" & _
 '       "<button class=""quality-btn"" onclick=""DownloadVideoQuality 'url" & i & "','1080'"">1080</button>" & _
 '       "<button class=""quality-btn"" onclick=""DownloadVideoQuality 'url" & i & "','720'"">720</button>" & _
 '       "<button class=""quality-btn"" onclick=""DownloadVideoQuality 'url" & i & "','480'"">480</button>" & _
 '       "<button class=""quality-btn"" onclick=""DownloadVideoQuality 'url" & i & "','360'"">360</button>" & _
 '       "<button class=""quality-btn"" onclick=""DownloadVideoQuality 'url" & i & "','144'"">144</button>" & _
 '       "<button class=""audio-btn"" onclick=""DownloadAudio 'url" & i & "'"">MP3</button>" & _
 '       "</div>" & _
 '       "<span class=""status"" id=""status" & i & """></span>" & _
 '       "</div>"
'Next
   
'    container.innerHTML = html
'End Sub

' Показать информационное окно
Sub ShowInfo()
    On Error Resume Next
    Document.getElementById("infoPanel").style.display = "block"
    Window.setTimeout "Document.getElementById('infoPanel').className = 'show'", "VBScript"
End Sub

Sub HideInfo()
    On Error Resume Next
    Document.getElementById("infoPanel").className = ""
    Document.getElementById("infoPanel").style.display = "none"
End Sub

' Выбор папки для сохранения
Dim g_SavePath

Sub SelectFolder()
    On Error Resume Next
    Dim shell, folder
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, "Выберите папку для сохранения", 0, "")
    If Not folder Is Nothing Then
        g_SavePath = folder.Self.Path
        Document.getElementById("savePath").value = g_SavePath
        WriteDebug "SelectFolder: выбран путь " & g_SavePath
		End If
End Sub
 
Sub resetsettings()
    On Error Resume Next
    If MsgBox("Сбросить настройки по умолчанию?", vbYesNo + vbQuestion, "Подтверждение") = vbYes Then
        SettingsPlaylist = "true"
        Call SaveSettings()
        SettingsPlaylist = ""
        ShowTempMessage "✅ Настройки сброшены"
        ' Обновляем окно
        LoadPlaylistSettings
        LoadPlaylistList
    End If
End Sub

Sub copyPlaylistSettings()
    On Error Resume Next
        SettingsPlaylist = "true"
        Call SaveSettings()
        SettingsPlaylist = ""
  End Sub
 
' ==================== ИСТОРИЯ ПЛЕЙЛИСТОВ ====================

Sub playlist_history()
    On Error Resume Next
    ShowPlaylistHistory
End Sub

Sub ShowPlaylistHistory()
    On Error Resume Next
    Document.getElementById("playlistPopup").style.display = "block"
    
    ' Загружаем настройки и список плейлистов
    LoadPlaylistSettings
    LoadPlaylistList
End Sub

Sub HidePlaylistHistory()
    On Error Resume Next
    Document.getElementById("playlistPopup").style.display = "none"
End Sub

Sub LoadPlaylistSettings()
    On Error Resume Next
    Dim fso, playlistFolder, settingsPath, settings, savePath, quality, format, subsValue, embeddedFlag
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    playlistFolder = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), "config\playlist\")
    settingsPath = fso.BuildPath(playlistFolder, "playlist_settings.txt")
    
	 ' Если файла настроек нет - создаем его из настроек по умолчанию
    If Not fso.FileExists(settingsPath) Then
        copyPlaylistSettings()
    End If
	
      ' Читаем настройки из файла

        Dim settingsFile, settingsArray
        Set settingsFile = fso.OpenTextFile(settingsPath, 1)
        settings = settingsFile.ReadAll
        settingsFile.Close
        
        settingsArray = Split(settings, "|")
        
        ' Берем значения ИЗ ФАЙЛА
        If UBound(settingsArray) >= 0 Then savePath = settingsArray(0)
        If UBound(settingsArray) >= 1 Then quality = settingsArray(1)
        If UBound(settingsArray) >= 2 Then format = settingsArray(2)
        If UBound(settingsArray) >= 4 Then subsValue = settingsArray(4)
        If UBound(settingsArray) >= 5 Then embeddedFlag = (settingsArray(5) = "true")
        
        ' Формируем отображение
        Dim subtitlesText
        If format = "mp3" Or subsValue = "none" Then
            subtitlesText = "Без субтитров"
        Else
            If embeddedFlag Then
                subtitlesText = "Субтитры: " & subsValue & " (встроенные)"
            Else
                subtitlesText = "Субтитры: " & subsValue & " (внешние)"
            End If
        End If
        
		Dim qualityFormat
		If format = "mp3" Then
			qualityFormat = "🎵 " & format
		ElseIf quality = "max" Then
			qualityFormat = "📺 ТОП 🎬 " & format & " 📝 " & subtitlesText
		Else
			qualityFormat = "📺 " & quality & "p 🎬 " & format & " 📝 " & 	subtitlesText
		End If
  
        Dim html
        html = "<div style='display: flex; justify-content: space-between; align-items: center; line-height: 1.5;'>"
        html = html & "<div>" & "Текущие настройки:" & "&nbsp;&nbsp;"
        html = html & qualityFormat & " 📁 " & savePath & "&nbsp;&nbsp;"
        html = html & "<button onclick=""VBScript:resetsettings"" title=""Загрузить настройки по умолчанию"" style='height: 24px; padding: 2px 8px; font-size: 12px;width: 90px;'>🔄 Сбросить</button>"
        html = html & "</div>"
        
        Document.getElementById("playlistSettings").innerHTML = html

End Sub


Sub LoadPlaylistList()
    On Error Resume Next
    Dim fso, playlistFolder, files, file, fileCollection, i
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    playlistFolder = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), "config\playlist\")
    
    If Not fso.FolderExists(playlistFolder) Then
        Document.getElementById("playlistList").innerHTML = "<div style='color:#888; text-align:center; padding:20px;'>Папка плейлистов не найдена</div>"
        Exit Sub
    End If
    
    Set files = fso.GetFolder(playlistFolder).Files
    Set fileCollection = CreateObject("Scripting.Dictionary")
    
    ' Собираем HTA файлы плейлистов
    For Each file In files
        If LCase(fso.GetExtensionName(file.Name)) = "hta" Then
            fileCollection.Add file.Name, file.Path
        End If
    Next
    
    ' Проверяем есть ли плейлисты
    If fileCollection.Count = 0 Then
        Document.getElementById("playlistList").innerHTML = "<div style='color:#888; text-align:center; padding:20px;'>Список плейлистов пуст</div>"
        Exit Sub
    End If
    
    ' Показываем список плейлистов
    Dim html, key, playlistId, jsonPath, jsonFile, jsonContent, title
    html = ""
    
    For Each key In fileCollection.Keys
        playlistId = Replace(key, ".hta", "")
        jsonPath = fso.BuildPath(playlistFolder, playlistId & ".json")
        title = "Нет заголовка"
        
        ' Пытаемся прочитать заголовок из JSON
If fso.FileExists(jsonPath) Then
    On Error Resume Next
    Set jsonFile = fso.OpenTextFile(jsonPath, 1)
    jsonContent = jsonFile.ReadAll
    jsonFile.Close
    
    ' Ищем playlist_title в JSON (пробуем разные варианты)
    Dim titlePattern, titleStart, titleEnd
    titlePattern = """playlist_title"": """  ' с пробелом и кавычкой
    titleStart = InStr(1, jsonContent, titlePattern, 1)
    
    If titleStart = 0 Then
        ' Пробуем без пробела
        titlePattern = """playlist_title"":"""
        titleStart = InStr(1, jsonContent, titlePattern, 1)
    End If
    
    If titleStart > 0 Then
        titleStart = titleStart + Len(titlePattern)
        titleEnd = InStr(titleStart, jsonContent, """", 1)
        If titleEnd > titleStart Then
            title = Mid(jsonContent, titleStart, titleEnd - titleStart)
        End If
    Else
        ' Проверяем есть ли ошибка
        If InStr(1, jsonContent, """error"":", 1) > 0 Then
            title = "Ошибка создания плейлиста"
        End If
    End If
End If
        
   html = html & "<div style='display:flex; justify-content:space-between; align-items:center; padding:10px; border-bottom:1px solid #333;'>"
        html = html & "<button onclick=""VBScript:OpenPlaylist '" & playlistId & "'"" style='flex-grow:1; text-align:left; margin-right:10px; padding:8px 12px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;' title='" & title & "'>" & title & "</button>"
   html = html & "<button onclick=""VBScript:DeletePlaylist '" & playlistId & "','" & Replace(title, "'", "''") & "'"" style='flex-shrink:0;'>🗑️</button>"
        html = html & "</div>"
    Next
    
    Document.getElementById("playlistList").innerHTML = html
End Sub

Sub OpenPlaylist(playlistId)
    On Error Resume Next
    
    Dim fso, playlistPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    playlistPath = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), "config\playlist\" & playlistId & ".hta")
    
    If fso.FileExists(playlistPath) Then
        CreateObject("WScript.Shell").Run Chr(34) & playlistPath & Chr(34)
        HidePlaylistHistory()
    Else
        ShowTempMessage "❌ Файл плейлиста не найден: " & playlistId
    End If
End Sub

Sub DeletePlaylist(playlistId, playlistTitle)
    On Error Resume Next
    
    If MsgBox("Вы уверены, что хотите удалить плейлист '" & playlistTitle & "'?", vbYesNo + vbQuestion, "Подтверждение удаления") = vbYes Then
        Dim fso, playlistFolder, htaPath, jsonPath
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        playlistFolder = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), "config\playlist\")
        htaPath = fso.BuildPath(playlistFolder, playlistId & ".hta")
        jsonPath = fso.BuildPath(playlistFolder, playlistId & ".json")
        
        If fso.FileExists(htaPath) Then fso.DeleteFile htaPath
        If fso.FileExists(jsonPath) Then fso.DeleteFile jsonPath
        
        ShowTempMessage "✅ Плейлист удален"
        
        ' Обновляем список
        LoadPlaylistList
    End If
End Sub