Public activeDownloadsCount, monitorInterval
' Инициализация приложения
Sub InitializeApp()
	On Error Resume Next
	SettingsPlaylist = ""
	activeDownloadsCount = 0
	monitorInterval = Empty
	Window_onLoad()
	Call InitializeBatFiles()
    LoadSettings()
	FormatSelectionChanged()
	SubtitlesChanged()
	Call CheckAutoDownload()
    Call InitProxyPlaceholder()
    Call UpdateProxyButtonColor()
    Call InitUrlFields()
    Call ToggleAutoDownload()
    StartStatusMonitor()
	'Call LoadFieldsFromMetadataLog()
	'Call LoadMetadataHistory()
	
End Sub

' Загрузка настроек
Sub LoadSettings()
    On Error Resume Next
    Dim fso, settingsFile, settingsPath, settings
    Set fso = CreateObject("Scripting.FileSystemObject")
    settingsPath = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), "config\downloader_settings.txt")

    If fso.FileExists(settingsPath) Then
        Set settingsFile = fso.OpenTextFile(settingsPath, 1)
        settings = Split(settingsFile.ReadAll, "|")
        settingsFile.Close

        ' 0: savePath
        If UBound(settings) >= 0 Then Document.getElementById("savePath").value = settings(0)

        ' 1: defaultQuality
        If UBound(settings) >= 1 Then Document.getElementById("defaultQuality").value = settings(1)

        ' 2: defaultFormat
        If UBound(settings) >= 2 Then Document.getElementById("defaultFormat").value = settings(2)

        ' 3: proxy
        If UBound(settings) >= 3 Then
            If Trim(settings(3)) <> "" Then
                Document.getElementById("proxy").value = settings(3)
            Else
                CheckProxyPlaceholder()
            End If
        Else
            CheckProxyPlaceholder()
        End If
		
		' 4: subtitles
		If UBound(settings) >= 4 Then 
			If Trim(settings(4)) <> "" Then
				Document.getElementById("subtitles").value = settings(4)
			End If
		End If

		' 5: embeddedSubs  
		If UBound(settings) >= 5 Then
			Dim embeddedSubsCheckbox
			Set embeddedSubsCheckbox = Document.getElementById("embeddedSubs")
			If Not embeddedSubsCheckbox Is Nothing Then
				embeddedSubsCheckbox.Checked = (settings(5) = "true")
			End If
		End If


' 6: detectedBrowser
If UBound(settings) >= 6 Then
    Dim detected, st, authCheckbox
    detected = Trim(settings(6))
    detectedBrowser = detected
    Set st = Document.getElementById("authBrowserStatus")
    Set authCheckbox = Document.getElementById("useBrowserAuth")
    
    If Not st Is Nothing Then
        If detected = "" Then
            st.innerText = "Не авторизован"
            st.style.color = "red"
        Else
            If Not authCheckbox Is Nothing Then
                If authCheckbox.Checked Then
                    st.innerText = detected & " вкл " 
                    st.style.color = "lime"
                Else
                    st.innerText = detected & " выкл" 
                    st.style.color = "red"
                End If
            Else
                st.innerText = detected
                st.style.color = "orange"
            End If
        End If
    End If
End If

        ' 7: autoCapture
        If UBound(settings) >= 7 Then
            Dim captureCheckbox
            Set captureCheckbox = Document.getElementById("autoCapture")
            If Not captureCheckbox Is Nothing Then
                captureCheckbox.Checked = (settings(7) = "true")
            End If
        End If

        ' 8: autoDownload
        If UBound(settings) >= 8 Then
            Dim downloadCheckbox
            Set downloadCheckbox = Document.getElementById("autoDownload")
            If Not downloadCheckbox Is Nothing Then
                downloadCheckbox.Checked = (settings(8) = "true")
            End If
        End If

        window.setTimeout "CheckProxyPlaceholder()", 100
    End If

    LoadProxyHistory()
End Sub

' Сохранение настроек
Sub SaveSettings()
    On Error Resume Next
    Dim fso, settingsFile, settingsPath
    Dim savePath, defaultQuality, defaultFormat, proxy
    Dim subtitles, embeddedSubs, detectedBrowser, autoCapture, autoDownload

 Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ★★★ ДОБАВЛЯЕМ УСЛОВИЕ ДЛЯ ВЫБОРА ПУТИ и КУКОВ★★★
    If SettingsPlaylist = "true" Then
        settingsPath = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), _
                                     "config\playlist\playlist_settings.txt")
    Else
        settingsPath = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), _
                                     "config\downloader_settings.txt")
    End If

    If SettingsPlaylist = "true" Then
        Dim mainSettingsPath, mainSettings
        mainSettingsPath = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), _
                                        "config\downloader_settings.txt")
        If fso.FileExists(mainSettingsPath) Then
            Set settingsFile = fso.OpenTextFile(mainSettingsPath, 1)
            mainSettings = Split(settingsFile.ReadAll, "|")
            settingsFile.Close
            If UBound(mainSettings) >= 6 Then 
                detectedBrowser = mainSettings(6) 
            Else
                detectedBrowser = ""
            End If
        Else
            detectedBrowser = ""
        End If
    End If
	
    ' === ЧИТАЕМ СТАРЫЕ НАСТРОЙКИ ===
    Dim oldSettings, txt
    If fso.FileExists(settingsPath) Then
        Set settingsFile = fso.OpenTextFile(settingsPath, 1)
        txt = settingsFile.ReadAll
        settingsFile.Close
        oldSettings = Split(txt, "|")
    Else
        ReDim oldSettings(8)
    End If

    ' === ВОССТАНАВЛИВАЕМ СТАРЫЕ ЗНАЧЕНИЯ ===
    If SettingsPlaylist <> "true" Then
        If UBound(oldSettings) >= 6 Then detectedBrowser = oldSettings(6) Else detectedBrowser = ""
    End If
    If UBound(oldSettings) >= 7 Then autoCapture = oldSettings(7) Else autoCapture = ""
    If UBound(oldSettings) >= 8 Then autoDownload = oldSettings(8) Else autoDownload = ""

    ' === ОБНОВЛЯЕМ ТЕКУЩИЕ ПАРАМЕТРЫ ===
    savePath = Document.getElementById("savePath").value
    defaultQuality = Document.getElementById("defaultQuality").value
    defaultFormat = Document.getElementById("defaultFormat").value
    proxy = Document.getElementById("proxy").value
    
    ' ★★★ ДОБАВЛЯЕМ СУБТИТРЫ ★★★
    subtitles = Document.getElementById("subtitles").value
    
    Dim embeddedSubsEl
    Set embeddedSubsEl = Document.getElementById("embeddedSubs")
    If Not embeddedSubsEl Is Nothing Then
        If embeddedSubsEl.Checked Then 
            embeddedSubs = "true" 
        Else 
            embeddedSubs = ""
        End If
    End If
    ' ★★★ КОНЕЦ СУБТИТРОВ ★★★

    ' Обновляем autoCapture
    Dim captureCheckbox
    Set captureCheckbox = Document.getElementById("autoCapture")
    If Not captureCheckbox Is Nothing Then
        If captureCheckbox.Checked Then autoCapture = "true" Else autoCapture = ""
    End If

    ' Обновляем autoDownload
    Dim downloadCheckbox
    Set downloadCheckbox = Document.getElementById("autoDownload")
    If Not downloadCheckbox Is Nothing Then
        If downloadCheckbox.Checked Then autoDownload = "true" Else autoDownload = ""
    End If

    ' === СОХРАНЯЕМ ВСЕ НАСТРОЙКИ ===
    If LCase(Trim(proxy)) = "http://ip:port или http://логин:пароль@ip:port" Then
        proxy = ""
    End If

    ReDim Preserve oldSettings(8)
    oldSettings(0) = savePath
    oldSettings(1) = defaultQuality
    oldSettings(2) = defaultFormat
    oldSettings(3) = proxy
    oldSettings(4) = subtitles      ' ← сохраняем субтитры
    oldSettings(5) = embeddedSubs   ' ← сохраняем чекбокс
    oldSettings(6) = detectedBrowser
    oldSettings(7) = autoCapture
    oldSettings(8) = autoDownload

    Set settingsFile = fso.CreateTextFile(settingsPath, True)
    settingsFile.Write Join(oldSettings, "|")
    settingsFile.Close

    SaveProxyHistory()
    Call UpdateProxyButtonColor()
End Sub

' Включение/выключение автозагрузки
Sub ToggleAutoDownload()
    On Error Resume Next
    ' Просто сохраняем настройки
    SaveSettings()
End Sub
' Выбор папки
Sub SelectFolder()
    On Error Resume Next
    Dim shell, folder
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, "Выберите папку для сохранения", 0, "")
    
    If Not folder Is Nothing Then
        Document.getElementById("savePath").value = folder.Self.Path
'        ShowTempMessage "❌ ВНИМАНИЕ! Файлы с одинаковыми именами будут перезаписаны без предупреждения!"
    End If
End Sub

'  Ярлык на рабочий стол
Sub CreateShortcut()
    Set WshShell = CreateObject("WScript.Shell")
    htaPath = window.location.pathname
    desktopPath = WshShell.SpecialFolders("Desktop")
    
    Set shortcut = WshShell.CreateShortcut(desktopPath & "\MultiLoader.lnk")
    shortcut.TargetPath = "mshta.exe"
    shortcut.Arguments = Chr(34) & htaPath & Chr(34)
    shortcut.WorkingDirectory = Left(htaPath, InStrRev(htaPath, "\"))
    shortcut.Save
    
    ShowTempMessage "✅ Ярлык создан на рабочем столе"
End Sub

' ★★★ ОЧИСТКА TEMP ПАПОК ★★★
Sub CleanTempFolders()
    On Error Resume Next
    Dim fso, folder, file, files
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Очищаем temp\cache\
    If fso.FolderExists("temp\cache\") Then
        Set folder = fso.GetFolder("temp\cache\")
        Set files = folder.Files
        For Each file In files
            fso.DeleteFile file.Path
        Next
    End If
    
    ' Очищаем temp\logs\
    If fso.FolderExists("temp\logs\") Then
        Set folder = fso.GetFolder("temp\logs\")
        Set files = folder.Files
        For Each file In files
            fso.DeleteFile file.Path
        Next
    End If
    
    ' Очищаем temp\bat\
    If fso.FolderExists("temp\bat\") Then
        Set folder = fso.GetFolder("temp\bat\")
        Set files = folder.Files
        For Each file In files
            fso.DeleteFile file.Path
        Next
    End If
End Sub

' Выход из приложения
Sub ExitApp()
    ShowTempMessage "Настройки сохранены!"
    SaveSettings()
	CleanTempFolders()
    window.setTimeout GetRef("CloseWindow"), 500
End Sub

Sub CloseWindow()
    Window.Close
End Sub

' Обработчик закрытия окна
Sub Window_onBeforeUnload()
	CleanTempFolders() 
    SaveSettings()
End Sub

' ==================== МОНИТОРИНГ СТАТУСОВ ЗАГРУЗОК ====================

Sub StartStatusMonitor()
    If IsEmpty(monitorInterval) Then
        monitorInterval = window.setInterval(GetRef("CheckAllDownloadsStatus"), 500)
    End If
End Sub

Sub IncrementDownloadsCount()
    activeDownloadsCount = activeDownloadsCount + 1
End Sub

Sub DecrementDownloadsCount()
    activeDownloadsCount = activeDownloadsCount - 1
    If activeDownloadsCount < 0 Then activeDownloadsCount = 0
End Sub

Sub CheckAllDownloadsStatus()
    On Error Resume Next
    If activeDownloadsCount = 0 Then Exit Sub

    Dim fso, folder, file, fieldId, statusValue
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists("temp\logs\") Then Exit Sub
    
    Set folder = fso.GetFolder("temp\logs\")
    
    For Each file In folder.Files
        If LCase(Right(file.Name, 7)) = ".status" Then
            fieldId = Replace(file.Name, ".status", "")
            
            ' Читаем статус (1 или 0)
            Dim statusFile
            Set statusFile = fso.OpenTextFile(file.Path, 1)
            statusValue = Trim(statusFile.ReadLine)
            statusFile.Close
            
            If statusValue = "1" Or statusValue = "0" Then
                ' Ищем файл в temp\cache\ для получения title
                Dim cachedFile, title
                cachedFile = FindFirstCachedFile(fieldId)
                
                If cachedFile <> "" Then
                    title = GetTitleFromFileName(cachedFile, fieldId)
                    ' Обновляем строку в metadata: статус И title
                    UpdateMetadataFieldId fieldId, title, statusValue
                Else
                    ' Если файл не найден, все равно обновляем статус
                    UpdateMetadataFieldId fieldId, "", statusValue
                End If
                
                ' Обновляем интерфейс
                Dim url
                url = GetUrlFromMetadata(fieldId)
                
                If statusValue = "1" Then
                    UpdateStatus fieldId, url, "completed"
                    ' Перемещаем файл если нашли
                    If cachedFile <> "" Then
                        MoveCachedFile fieldId, cachedFile, title
                    End If
                Else
                    UpdateStatus fieldId, url, "error"
                End If
                
                ' Удаляем файл статуса
                fso.DeleteFile file.Path
                DecrementDownloadsCount()
            End If
        End If
    Next
End Sub

Sub UpdateMetadataFieldId(fieldId, title, statusValue)
    On Error Resume Next
    Dim fso, logPath, logFile, tempFile, line, arr, newStatus
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    If Not fso.FileExists(logPath) Then Exit Sub
    
    ' Определяем статус
    If statusValue = "1" Then
        newStatus = "completed"
    Else
        newStatus = "error"
    End If
    
    Dim tempPath
    tempPath = logPath & ".tmp"
    
    Set logFile = fso.OpenTextFile(logPath, 1)
    Set tempFile = fso.CreateTextFile(tempPath, True)
    
    Do Until logFile.AtEndOfStream
        line = Trim(logFile.ReadLine)
        
        If line <> "" Then
            arr = Split(line, "|")
            
            ' Если нашли нужный fieldId (обязательно должен быть)
            If UBound(arr) >= 0 And arr(0) = fieldId Then
                ' Обновляем статус (колонка 3)
                arr(3) = newStatus
                
                ' Обновляем title (колонка 4) если передан
                If title <> "" Then
                    ' Гарантируем что есть 4-я колонка
                    If UBound(arr) < 4 Then
                        ReDim Preserve arr(4)
                    End If
                    arr(4) = title
                End If
                
                line = Join(arr, "|")
            End If
            
            tempFile.WriteLine line
        End If
    Loop
    
    logFile.Close
    tempFile.Close
    
    ' Заменяем файл
    fso.DeleteFile logPath
    fso.MoveFile tempPath, logPath
End Sub

Function FindFirstCachedFile(fieldId)
    On Error Resume Next
    Dim fso, cacheFolder, fileObj
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    FindFirstCachedFile = ""
    
    If Not fso.FolderExists("temp\cache\") Then Exit Function
    
    Set cacheFolder = fso.GetFolder("temp\cache\")
    
    ' Ищем первый файл с префиксом fieldId_
    For Each fileObj In cacheFolder.Files
        If Left(fileObj.Name, Len(fieldId) + 1) = fieldId & "_" Then
            FindFirstCachedFile = fileObj.Name
            Exit Function
        End If
    Next
End Function

Function GetTitleFromFileName(fullName, fieldId)
    On Error Resume Next
    ' Убираем fieldId_ из начала
    GetTitleFromFileName = Mid(fullName, Len(fieldId) + 2)
End Function

Sub MoveCachedFile(fieldId, cachedFileName, cleanTitle)
    On Error Resume Next
    Dim fso, savePath, sourcePath, destPath
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    savePath = Document.getElementById("savePath").value
    If savePath = "" Then Exit Sub
    If Right(savePath, 1) <> "\" Then savePath = savePath & "\"
    
    sourcePath = "temp\cache\" & cachedFileName
    
    ' Если cleanTitle уже очищен - используем его
    ' Иначе очищаем прямо здесь
    If cleanTitle = "" Then
        cleanTitle = GetSafeTitle(cachedFileName, fieldId)
    End If
    
    destPath = savePath & cleanTitle
    
    If fso.FileExists(sourcePath) Then
        ' Удаляем старый файл если существует
        If fso.FileExists(destPath) Then
            fso.DeleteFile destPath
        End If
        
        ' Пробуем переместить
        fso.MoveFile sourcePath, destPath
        
        ' Если ошибка (из-за имени файла), пробуем с другим именем
        If Err.Number <> 0 Then
            Err.Clear
            ' Пробуем просто с fieldId
            destPath = savePath & fieldId & ".mp4"
            fso.MoveFile sourcePath, destPath
        End If
    End If
End Sub

Function GetTitleFromFileName(fullName, fieldId)
    On Error Resume Next
    
    ' Убираем fieldId_ из начала
    Dim rawTitle
    rawTitle = Mid(fullName, Len(fieldId) + 2)
    
    ' Очищаем от .part
    If LCase(Right(rawTitle, 5)) = ".part" Then
        rawTitle = Left(rawTitle, Len(rawTitle) - 5)
    End If
    
    ' Убираем расширение файла
    Dim fso, baseName, extension
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Сохраняем оригинальное расширение
    extension = fso.GetExtensionName(rawTitle)
    baseName = fso.GetBaseName(rawTitle)
    
    ' Очищаем имя файла от запрещенных символов
    Dim cleanName
    cleanName = CleanFileName(baseName)
    
    ' Возвращаем с расширением
    If extension <> "" Then
        GetTitleFromFileName = cleanName & "." & extension
    Else
        GetTitleFromFileName = cleanName
    End If
End Function

Function CleanFileName(fileName)
    On Error Resume Next
    
    Dim result, i, char
    
    result = ""
    
    ' Заменяем запрещенные символы Windows
    For i = 1 To Len(fileName)
        char = Mid(fileName, i, 1)
        
        Select Case char
            ' Запрещенные символы в Windows
            Case "\", "/", ":", "*", "?", """", "<", ">", "|"
                result = result & "_"
            ' Нестандартные символы
            Case "？", "！", "，", "。", "；", "：", "「", "」", "【", "】"
                result = result & "_"
            ' Обычные символы - оставляем как есть
            Case Else
                result = result & char
        End Select
    Next
    
    ' Убираем начальные/конечные пробелы и точки
    result = Trim(result)
    result = RTrim(result, ".")
    
    ' Если после очистки имя пустое - даем дефолтное
    If result = "" Then result = "video"
    
    ' Ограничиваем длину (Windows max 255, но лучше короче)
    If Len(result) > 150 Then
        result = Left(result, 150)
    End If
    
    CleanFileName = result
End Function

' Вспомогательная функция для удаления символов с конца строки
Function RTrim(str, chars)
    On Error Resume Next
    Dim result
    result = str
    
    Do While Right(result, 1) = chars
        result = Left(result, Len(result) - 1)
    Loop
    
    RTrim = result
End Function