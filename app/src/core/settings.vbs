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

	
    Dim fso, folder, files, file, fieldId, status, url
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists("temp\logs\") Then
        Set folder = fso.GetFolder("temp\logs\")
        Set files = folder.Files
        
        For Each file In files
            If LCase(Right(file.Name, 7)) = ".status" Then
                fieldId = Replace(file.Name, ".status", "")
                
                ' Читаем статус
                Dim statusFile
                Set statusFile = fso.OpenTextFile(file.Path, 1)
                status = Trim(statusFile.ReadLine)
                statusFile.Close
                
                If status = "1" Or status = "0" Then
                    ' Нашли завершенную загрузку
                    url = GetUrlFromMetadata(fieldId)
                    
                    If status = "1" Then
                        UpdateMetadataLogStatus fieldId, url, "completed"
                        UpdateStatus fieldId, url, "completed"
                        MoveCompletedFile fieldId
                    Else
                        UpdateMetadataLogStatus fieldId, url, "error" 
                        UpdateStatus fieldId, url, "error"
                    End If
                    
                    ' Удаляем статус-файл и уменьшаем счетчик
                    fso.DeleteFile file.Path
                    DecrementDownloadsCount()
                End If
            End If
        Next
    End If
End Sub

Function GetUrlFromMetadata(fieldId)
    On Error Resume Next
    Dim fso, logFile, logPath, line, arr
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    GetUrlFromMetadata = ""
    
    If fso.FileExists(logPath) Then
        Set logFile = fso.OpenTextFile(logPath, 1)
        Do Until logFile.AtEndOfStream
            line = Trim(logFile.ReadLine)
            If line <> "" Then
                arr = Split(line, "|")
                If UBound(arr) >= 2 Then
                    If arr(0) = fieldId Then
                        GetUrlFromMetadata = arr(2)
                        Exit Do
                    End If
                End If
            End If
        Loop
        logFile.Close
    End If
End Function

Sub MoveCompletedFile(fieldId)
    On Error Resume Next
    Dim fso, cacheFolder, files, file, savePath, newName
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    savePath = Document.getElementById("savePath").value
    If savePath = "" Then Exit Sub
    If Right(savePath, 1) <> "\" Then savePath = savePath & "\"
    
    If fso.FolderExists("temp\cache\") Then
        Set cacheFolder = fso.GetFolder("temp\cache\")
        Set files = cacheFolder.Files
        
        For Each file In files
            If Left(file.Name, Len(fieldId) + 1) = fieldId & "_" Then
                newName = Mid(file.Name, Len(fieldId) + 2)
                
                If fso.FileExists(savePath & newName) Then
                    fso.DeleteFile savePath & newName
                End If
                fso.MoveFile file.Path, savePath & newName
                Exit For
            End If
        Next
    End If
End Sub

