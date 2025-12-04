' ==============================
' downloadauto.vbs - автозагрузка одиночных видео
' ==============================
Option Explicit

' Загрузка одного видео
Sub DownloadSingleVideo(url, fieldId)
    On Error Resume Next
	
    ' === ОБНОВЛЯЕМ СТАТУС НА DOWNLOADING ===
    UpdateStatus fieldId, url, STATUS_DOWNLOADING
	
    Dim shell, fso, outputPath, currentDir, cmd, proxy
    Dim defaultQuality, defaultFormat, actualFormat, binPath
    Dim tempCachePath, tempLogsPath, logFilePath
	
    ' Получаем настройки
    defaultQuality = Document.getElementById("defaultQuality").value
    defaultFormat = Document.getElementById("defaultFormat").value
    proxy = GetProxyAddress()
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    currentDir = fso.GetParentFolderName(window.location.pathname)
    binPath = fso.BuildPath(currentDir, "bin\yt-dlp.exe")
    
    ' Временные пути
    tempCachePath = currentDir & "\temp\cache\"
    tempLogsPath = currentDir & "\temp\logs\"

    ' Создаем временные папки если не существуют
    If Not fso.FolderExists(tempCachePath) Then fso.CreateFolder(tempCachePath)
    If Not fso.FolderExists(tempLogsPath) Then fso.CreateFolder(tempLogsPath)

    ' Лог файл для этого fieldId
    logFilePath = tempLogsPath & fieldId & ".log"
    
    ' Проверка yt-dlp
    If Not fso.FileExists(binPath) Then 
        ' Если ошибка - обновляем статус
        UpdateStatus fieldId, url, STATUS_ERROR
        Exit Sub
    End If
    
    ' ★★★ СОЗДАЕМ УНИКАЛЬНЫЙ БАТНИК-ОБЕРТКУ ★★★
    Dim batPath
    batPath = CreateWrapperBat(fieldId, currentDir)
    
    ' Формируем команду через батник-обертку
    cmd = "cd /d " & Chr(34) & currentDir & Chr(34) & " && " & batPath
    
    If proxy <> "" Then
        cmd = cmd & " --proxy " & Chr(34) & proxy & Chr(34)
    End If
    
    ' Авторизация
    Dim authParams
    authParams = GetBrowserAuthParams()
    If Trim(authParams) <> "" Then
        cmd = cmd & " " & authParams
    End If
    
    ' Субтитры
    Dim subtitleKeys
    subtitleKeys = GenerateSubtitleKeys()
    If Trim(subtitleKeys) <> "" Then
        cmd = cmd & " " & subtitleKeys
    End If
    
    ' ★★★ ПРОВЕРКА НА MP3 ★★★
    If LCase(defaultFormat) = "mp3" Then
        ' АУДИО РЕЖИМ - упрощенные параметры
        cmd = cmd & " -x --audio-format mp3 --audio-quality 0"
        cmd = cmd & " -o " & Chr(34) & tempCachePath & fieldId & "_%(title)s.%(ext)s" & Chr(34)
    Else
        ' ВИДЕО РЕЖИМ - существующая логика
        actualFormat = defaultFormat
        If defaultFormat = "webm" Or defaultFormat = "mkv" Then
            actualFormat = "best"
        End If
    
        If actualFormat = "best" Then
            If defaultQuality = "max" Then
                cmd = cmd & " -o " & Chr(34) & tempCachePath & fieldId & "_%(title)s_top.%(ext)s" & Chr(34)
            Else
                cmd = cmd & " -o " & Chr(34) & tempCachePath & fieldId & "_%(title)s_" & defaultQuality & "p.%(ext)s" & Chr(34) & _
                      " -f " & Chr(34) & "best[height<=" & defaultQuality & "]" & Chr(34)
            End If
        Else
            If defaultQuality = "max" Then
                cmd = cmd & " -o " & Chr(34) & tempCachePath & fieldId & "_%(title)s_top.%(ext)s" & Chr(34) & _
                      " -f " & Chr(34) & "best[ext=" & actualFormat & "]/best" & Chr(34)
            Else
                cmd = cmd & " -o " & Chr(34) & tempCachePath & fieldId & "_%(title)s_" & defaultQuality & "p.%(ext)s" & Chr(34) & _
                      " -f " & Chr(34) & "best[height<=" & defaultQuality & "][ext=" & actualFormat & "]/best[height<=" & defaultQuality & "]" & Chr(34)
            End If
        End If
    End If
    
    ' Добавляем URL (статус-файлы создаются в батнике)
    cmd = cmd & " " & Chr(34) & url & Chr(34)
    
    ' === Обновляем статус в metadata_history.log ===
    If url <> "" Then
        On Error Resume Next
        UpdateMetadataLogStatus CStr(fieldId), url, "downloading"
        On Error GoTo 0
    End If
	
    ' ★★★ УВЕЛИЧИВАЕМ СЧЕТЧИК ★★★
    activeDownloadsCount = activeDownloadsCount + 1
    	
' Запускаем батник
shell.Run "cmd /c " & cmd, 1, False

End Sub

' ★★★ СОЗДАНИЕ БАТНИКА-ОБЕРТКИ ★★★
Function CreateWrapperBat(fieldId, currentDir)
    On Error Resume Next
    Dim fso, batContent, batPath, batFolder, batFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Создаем папку для батников если нет
    batFolder = currentDir & "\temp\bat\"
    If Not fso.FolderExists(batFolder) Then
        fso.CreateFolder(batFolder)
    End If
    
    batPath = "temp\bat\" & fieldId & ".bat"
    Dim fullBatPath
    fullBatPath = currentDir & "\" & batPath
    
  batContent = "@echo off" & vbCrLf & _
			 "chcp 1251 >nul" & vbCrLf & _ 
             "cd /d " & Chr(34) & currentDir & Chr(34) & vbCrLf & _
             "bin\yt-dlp.exe %*" & vbCrLf & _
             "if %errorlevel% equ 0 (" & vbCrLf & _
             "echo 1 > temp\logs\" & fieldId & ".status" & vbCrLf & _
             ") else (" & vbCrLf & _
             "echo 0 > temp\logs\" & fieldId & ".status" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo Были ошибки при загрузке!" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo Решение:" & vbCrLf & _
             "echo • Используйте прокси/VPN" & vbCrLf & _
             "echo • Для прямых эфиров - дождитесь обработки YouTube" & vbCrLf & _
             "echo • Проверьте доступность видео" & vbCrLf & _
             "echo • Проверьте правильность ссылок" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo Нажмите любую клавишу для закрытия..." & vbCrLf & _
             "pause >nul" & vbCrLf & _
             ")" & vbCrLf & _
             "del " & Chr(34) & currentDir & "\temp\bat\" & fieldId & ".bat" & Chr(34) & vbCrLf
        
    Set batFile = fso.CreateTextFile(fullBatPath, True)
    batFile.Write batContent
    batFile.Close
    
    
    CreateWrapperBat = batPath
End Function

' ==================== ПОЛУЧЕНИЕ ПАРАМЕТРОВ АВТОРИЗАЦИИ ====================

Function GetBrowserAuthParams()
    On Error Resume Next
    Dim authCheckbox
    
    ' Если detectedBrowser пустой - грузим настройки
    If detectedBrowser = "" Then
        LoadSettings()
    End If
    
    Set authCheckbox = Document.getElementById("useBrowserAuth")
    
    If Not authCheckbox Is Nothing And authCheckbox.Checked Then
        If detectedBrowser <> "" Then
            GetBrowserAuthParams = "--cookies-from-browser " & Chr(34) & detectedBrowser & Chr(34)
        Else
            GetBrowserAuthParams = ""
        End If
    Else
        GetBrowserAuthParams = ""
    End If
End Function

' ==================== ГЕНЕРАЦИЯ ПАРАМЕТРОВ СУБТИТРОВ ====================

Function GenerateSubtitleKeys()
    On Error Resume Next
    Dim subtitlesSelect, embeddedSubsCheckbox, formatSelect
    Dim subValue, embedValue, keys
    
    Set subtitlesSelect = Document.getElementById("subtitles")
    Set embeddedSubsCheckbox = Document.getElementById("embeddedSubs")
    Set formatSelect = Document.getElementById("defaultFormat")
    
    ' ★★★ ЕСЛИ ФОРМАТ MP3 - НЕТ СУБТИТРОВ ★★★
    If LCase(formatSelect.value) = "mp3" Then
        GenerateSubtitleKeys = ""
        Exit Function
    End If
    
    subValue = subtitlesSelect.value
    embedValue = embeddedSubsCheckbox.Checked
    
    keys = ""
    
    Select Case subValue
        Case "none"
            keys = ""
            
        Case "ru", "en"
            keys = "--write-subs --sub-langs " & subValue & " --ignore-errors"
            If embedValue Then
                keys = keys & " --embed-subs"
            End If
            
        Case "auto"
            keys = "--write-auto-subs" & " --ignore-errors"
            If embedValue Then
                keys = keys & " --embed-subs"
            End If
    End Select
       
    GenerateSubtitleKeys = keys
End Function

' ==================== ПОЛУЧЕНИЕ АДРЕСА ПРОКСИ ====================

Function GetProxyAddress()
    On Error Resume Next
    Dim proxyField, proxy
    Set proxyField = Document.getElementById("proxy")
    
    If Not proxyField Is Nothing Then
        proxy = Trim(proxyField.value)
        ' Убираем placeholder
        If proxy = "http://ip:port или http://логин:пароль@ip:port" Then
            proxy = ""
        End If
    End If
    
    GetProxyAddress = proxy
End Function
' ==================== МАССОВАЯ ЗАГРУЗКА ВСЕХ WAITING ССЫЛОК ====================

Sub DownloadAll()
    On Error Resume Next
    Dim fso, logFile, logPath, line, arr, fieldId, url, status
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    If Not fso.FileExists(logPath) Then
        ShowTempMessage "❌ Нет ссылок для загрузки!"
        Exit Sub
    End If
    
    Dim waitingUrls
    Set waitingUrls = CreateObject("Scripting.Dictionary")
    
    ' Ищем все ссылки со статусом waiting
    Set logFile = fso.OpenTextFile(logPath, 1)
    Do Until logFile.AtEndOfStream
        line = Trim(logFile.ReadLine)
        If line <> "" Then
            arr = Split(line, "|")
            If UBound(arr) >= 3 Then
                fieldId = arr(0)
                url = arr(2)
                status = arr(3)
                
                If status = "waiting" Then
                    waitingUrls.Add fieldId, url
                End If
            End If
        End If
    Loop
    logFile.Close
    
    If waitingUrls.Count = 0 Then
        ShowTempMessage "❌ Нет ссылок со статусом 'ожидание'!"
        Exit Sub
    End If
    
    ' Запускаем загрузку для каждой ссылки
    Dim keys, i
    keys = waitingUrls.Keys
    For i = 0 To waitingUrls.Count - 1
        fieldId = keys(i)
        url = waitingUrls(fieldId)
        DownloadSingleVideo url, fieldId
    Next
    
    ShowTempMessage "✅ Запущено " & waitingUrls.Count & " загрузок!"
End Sub