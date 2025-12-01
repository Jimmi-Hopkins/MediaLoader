' downloadplaylist.vbs
' Загрузка выбранных видео из JSON-плейлиста (config\playlist\playlist_<ID>.json)
' Берёт настройки из config\playlist\playlist_settings.txt (позиции 0..6)
Option Explicit
Dim detectedBrowser

' Entry point (вызывается из playlist HTA)
Sub downplaylist()
    On Error Resume Next
    Dim playlistId, jsonPath, playlistData, selectedUrls
    Dim fso, basePath, settings

    playlistId = ExtractPlaylistId()
    If playlistId = "" Then Exit Sub
    
    jsonPath = DetectJsonPathForDownload()

    Set playlistData = LoadPlaylistJson(jsonPath)
    If playlistData Is Nothing Then Exit Sub
    
    Set selectedUrls = GetSelectedUrls(playlistData)
    If selectedUrls Is Nothing Then Exit Sub
    If selectedUrls.Count = 0 Then Exit Sub

    Set settings = LoadPlaylistSettings()

    If settings Is Nothing Then Exit Sub

    ' ★★★ ВСЕ ПЕРЕМЕННЫЕ ОПРЕДЕЛЕНЫ ★★★
    
    DownloadPlaylistVideos selectedUrls, playlistData, settings, playlistId
End Sub

Function DetectJsonPathForDownload()
    Dim fso, playlistId, currentFolder, absolutePath
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    playlistId = ExtractPlaylistId()
    If playlistId = "" Then
        DetectJsonPathForDownload = ""
        Exit Function
    End If
    
    ' Получаем абсолютный путь к папке с HTA
    currentFolder = fso.GetParentFolderName(window.location.pathname)
    absolutePath = fso.BuildPath(currentFolder, "playlist_" & playlistId & ".json")
    
    DetectJsonPathForDownload = absolutePath
End Function
' -----------------------
' Чтение локальных настроек playlist_settings.txt (в config\playlist)
' Формат: same as downloader_settings.txt (позиции 0..6; 0=savePath,1=defaultQuality,2=defaultFormat,3=proxy,4=subtitles,5=embeddedSubs,6=detectedBrowser)
' -----------------------
Function LoadPlaylistSettings()
    On Error Resume Next
    Dim fso, settingsPath, tf, txt, arr, result
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Жесткий путь к основному файлу настроек
    settingsPath = fso.BuildPath(fso.GetParentFolderName(fso.GetParentFolderName(window.location.pathname)), "playlist\playlist_settings.txt")
        
    If Not fso.FileExists(settingsPath) Then
        MsgBox "ФАЙЛ НАСТРОЕК НЕ НАЙДЕН!"
        Set LoadPlaylistSettings = Nothing
        Exit Function
    End If

    Set tf = fso.OpenTextFile(settingsPath, 1)
    txt = tf.ReadAll
    tf.Close
    
    arr = Split(txt, "|")
    Set result = CreateObject("Scripting.Dictionary")

    If UBound(arr) >= 0 Then result("savePath") = arr(0) Else result("savePath") = ""
    If UBound(arr) >= 1 Then result("defaultQuality") = arr(1) Else result("defaultQuality") = "360"
    If UBound(arr) >= 2 Then result("defaultFormat") = arr(2) Else result("defaultFormat") = "mp4"
    If UBound(arr) >= 3 Then result("proxy") = arr(3) Else result("proxy") = ""
    If UBound(arr) >= 4 Then result("subtitles") = arr(4) Else result("subtitles") = "none"
    If UBound(arr) >= 5 Then result("embeddedSubs") = LCase(arr(5)) Else result("embeddedSubs") = "false"
    If UBound(arr) >= 6 Then
        result("detectedBrowser") = Trim(arr(6))
        detectedBrowser = Trim(arr(6))
    Else
        result("detectedBrowser") = ""
    End If

    Set LoadPlaylistSettings = result
End Function

' -----------------------
' Загрузка JSON плейлиста (возвращает словарь: raw_json, playlist_title, source_url)
' -----------------------
Function LoadPlaylistJson(path)
    On Error Resume Next
    Dim fso, txt, result
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(path) Then
        Set LoadPlaylistJson = Nothing
        Exit Function
    End If

    txt = ReadFileForDownload(path)
    If txt = "" Then
        Set LoadPlaylistJson = Nothing
        Exit Function
    End If

    Set result = CreateObject("Scripting.Dictionary")
    result("raw_json") = txt
    result("playlist_title") = ExtractValue(txt, "playlist_title")
    result("source_url") = ExtractValue(txt, "source_url")

    Set LoadPlaylistJson = result
	
End Function

' -----------------------
' Извлечение выбранных URL'ов (selected = true)
' -----------------------
Function GetSelectedUrls(playlistData)

    On Error Resume Next
    Dim selectedUrls, json, itemsStart, itemsEnd, itemsContent, itemBlocks, i, itemBlock
    Set selectedUrls = CreateObject("Scripting.Dictionary")
    If playlistData Is Nothing Then
        Set GetSelectedUrls = selectedUrls
        Exit Function
    End If

    json = playlistData("raw_json")
    itemsStart = InStr(json, """items""")
    If itemsStart = 0 Then
        Set GetSelectedUrls = selectedUrls
        Exit Function
    End If

    itemsStart = InStr(itemsStart, json, "[")
    If itemsStart = 0 Then
        Set GetSelectedUrls = selectedUrls
        Exit Function
    End If

    itemsEnd = InStr(itemsStart, json, "]")
    If itemsEnd = 0 Then
        Set GetSelectedUrls = selectedUrls
        Exit Function
    End If

    itemsContent = Mid(json, itemsStart + 1, itemsEnd - itemsStart - 1)
    ' Разделяем на блоки по "},", но учитываем возможность запятых в полях — простая эвристика, совпадает с форматом, который вы используете
    itemBlocks = Split(itemsContent, "},")

    For i = 0 To UBound(itemBlocks)
        itemBlock = Trim(itemBlocks(i))
        If itemBlock <> "" Then
            If Left(itemBlock,1) = "{" Then itemBlock = Mid(itemBlock,2)
            If Right(itemBlock,1) = "}" Then itemBlock = Left(itemBlock, Len(itemBlock)-1)
            Dim selected, url
            selected = ExtractValue("{" & itemBlock & "}", "selected")
            url = ExtractValue("{" & itemBlock & "}", "url")
            If LCase(Trim(selected)) = "true" And Trim(url) <> "" Then
                selectedUrls.Add selectedUrls.Count, url
            End If
        End If
    Next

    Set GetSelectedUrls = selectedUrls
End Function

' -----------------------
' Основная загрузка (формирует путь, команду и запускает yt-dlp)
' selectedUrls — Dictionary с url-ами
' settings — словарь настроек из playlist_settings.txt
' playlistId — строка id
' -----------------------
Sub DownloadPlaylistVideos(selectedUrls, playlistData, settings, playlistId)
    On Error Resume Next
	
    Dim fso, shell, binPath, currentDir, outputRoot, playlistTitle, finalOutputPath
    Dim defaultQuality, defaultFormat, actualFormat, proxy, validUrls, i, url

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    ' путь к yt-dlp относительно папки playlist (portable)
    currentDir = fso.GetParentFolderName(window.location.pathname) ' это папка config\playlist
    binPath = "..\..\bin\yt-dlp.exe"

    If Not fso.FileExists(fso.BuildPath(currentDir, binPath)) Then
        ' если yt-dlp не найден — тихо выходим
        Exit Sub
    End If

    ' Настройки
    defaultQuality = settings("defaultQuality")
    defaultFormat = settings("defaultFormat")
    proxy = settings("proxy")

    ' Путь сохранения — берем из settings(0)
    outputRoot = settings("savePath")
    If Trim(outputRoot) = "" Then Exit Sub
    If Right(outputRoot,1) <> "\" Then outputRoot = outputRoot & "\"

    ' Папка для этого плейлиста: savepath\playlist_<ID>_<CleanTitle>\
    playlistTitle = CleanFileName(playlistData("playlist_title"))
    finalOutputPath = outputRoot & "playlist_" & playlistId & "_" & playlistTitle & "\"
    If Not fso.FolderExists(finalOutputPath) Then
        On Error Resume Next
        fso.CreateFolder(finalOutputPath)
    End If

    ' Собираем URL'ы
    validUrls = ""
    For i = 0 To selectedUrls.Count - 1
        url = Trim(selectedUrls(i))
        If url <> "" Then
            validUrls = validUrls & " " & Chr(34) & url & Chr(34)
        End If
    Next

    If Trim(validUrls) = "" Then Exit Sub

    actualFormat = defaultFormat
    If LCase(defaultFormat) = "webm" Or LCase(defaultFormat) = "mkv" Then actualFormat = "best"

    ' Формируем команду: cd в папку config\playlist (currentDir) затем относительный путь к yt-dlp
    Dim cmd
    cmd = "cd /d " & Chr(34) & currentDir & Chr(34) & " && " & binPath

    If Trim(proxy) <> "" Then
        cmd = cmd & " --proxy " & Chr(34) & proxy & Chr(34)
    End If

    Dim authParams
    authParams = GetBrowserAuthParamsFromSettings(settings)
    If Trim(authParams) <> "" Then cmd = cmd & " " & authParams

    Dim subtitleKeys
    subtitleKeys = GenerateSubtitleKeysFromSettings(settings)
    If Trim(subtitleKeys) <> "" Then cmd = cmd & " " & subtitleKeys

    ' Формируем параметры качества/формата/выхода
    If actualFormat = "best" Then
        If defaultQuality = "max" Then
            cmd = cmd & " -o " & Chr(34) & finalOutputPath & "%(title)s_top.%(ext)s" & Chr(34)
        Else
            cmd = cmd & " -o " & Chr(34) & finalOutputPath & "%(title)s_" & defaultQuality & "p.%(ext)s" & Chr(34) & _
                  " -f " & Chr(34) & "best[height<=" & defaultQuality & "]" & Chr(34)
        End If
    Else
        If defaultQuality = "max" Then
            cmd = cmd & " -o " & Chr(34) & finalOutputPath & "%(title)s_top.%(ext)s" & Chr(34) & _
                  " -f " & Chr(34) & "best[ext=" & actualFormat & "]/best" & Chr(34)
        Else
            cmd = cmd & " -o " & Chr(34) & finalOutputPath & "%(title)s_" & defaultQuality & "p.%(ext)s" & Chr(34) & _
                  " -f " & Chr(34) & "best[height<=" & defaultQuality & "][ext=" & actualFormat & "]/best[height<=" & defaultQuality & "]" & Chr(34)
        End If
    End If

    cmd = cmd & validUrls

    ' Запускаем в видимом окне консоли (не блокируем HTA)
    cmd = "cmd /c " & Chr(34) & "echo Загрузка " & selectedUrls.Count & " видео из плейлиста..." & " && " & cmd & _
          " && (echo. && echo Все видео загружены!) && exit || (echo. && echo Были ошибки при загрузке!) & pause" & Chr(34)

    shell.Run cmd, 1, False
End Sub

' -----------------------
' Функция для генерации параметров субтитров из settings (позиции 4 и 5)
' -----------------------
Function GenerateSubtitleKeysFromSettings(settings)
    On Error Resume Next
    Dim subValue, embedValue, keys
    subValue = settings("subtitles")
    embedValue = settings("embeddedSubs") ' "true"/"false" или ""

    keys = ""
    Select Case LCase(Trim(subValue))
        Case "none"
            keys = ""
        Case "ru", "en"
            keys = "--write-subs --sub-langs " & subValue & " --ignore-errors"
            If LCase(Trim(embedValue)) = "true" Then keys = keys & " --embed-subs"
        Case "auto"
            keys = "--write-auto-subs --ignore-errors"
            If LCase(Trim(embedValue)) = "true" Then keys = keys & " --embed-subs"
    End Select

    GenerateSubtitleKeysFromSettings = keys
End Function

' -----------------------
' Получение параметра авторизации (--cookies-from-browser ...) из settings (detectedBrowser)
' -----------------------
Function GetBrowserAuthParamsFromSettings(settings)
    On Error Resume Next
    Dim det
    det = ""
    If IsObject(settings) Then
        If settings.Exists("detectedBrowser") Then det = Trim(settings("detectedBrowser"))
    End If
    If det <> "" Then
        GetBrowserAuthParamsFromSettings = "--cookies-from-browser " & Chr(34) & det & Chr(34)
    Else
        GetBrowserAuthParamsFromSettings = ""
    End If
End Function

' ==================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================

Function ReadFileForDownload(path)
    On Error Resume Next
    
    Dim fso, tf, content
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(path) Then
          ReadFileForDownload = ""
        Exit Function
    End If

    Set tf = fso.OpenTextFile(path, 1) ' 1 = ForReading
    If tf Is Nothing Then
        ReadFileForDownload = ""
        Exit Function
    End If
    
    content = tf.ReadAll
    tf.Close
    
    ReadFileForDownload = content
End Function

Function ExtractValue(txt, key)
    Dim p, i, ch, result
    p = InStr(txt, """" & key & """")
    If p = 0 Then ExtractValue = "": Exit Function
    p = InStr(p, txt, ":")
    If p = 0 Then ExtractValue = "": Exit Function
    p = p + 1
    Do While p <= Len(txt) And (Mid(txt,p,1) = " " Or Mid(txt,p,1) = vbTab)
        p = p + 1
    Loop
    If p > Len(txt) Then ExtractValue = "": Exit Function

    If Mid(txt, p, 1) = """" Then
        p = p + 1
        result = ""
        For i = p To Len(txt)
            ch = Mid(txt, i, 1)
            If ch = """" And Mid(txt, i - 1, 1) <> "\" Then Exit For
            result = result & ch
        Next
    Else
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

Function CleanFileName(fileName)
    On Error Resume Next
    Dim invalidChars, i, result, ch
    invalidChars = "\/:*?""<>|"
    result = ""
    For i = 1 To Len(fileName)
        ch = Mid(fileName, i, 1)
        If InStr(invalidChars, ch) = 0 Then result = result & ch Else result = result & "_"
    Next
    If Len(result) > 100 Then result = Left(result, 100)
    CleanFileName = Trim(result)
End Function

Function ExtractPlaylistId()
    On Error Resume Next
    Dim htaPath, fso, fileName, id
    Set fso = CreateObject("Scripting.FileSystemObject")
    htaPath = Replace(window.location.pathname, "/", "\")
    fileName = fso.GetFileName(htaPath)
    If InStr(fileName, "playlist_") = 1 And InStr(fileName, ".hta") > 0 Then
        id = Replace(fileName, "playlist_", "")
        id = Replace(id, ".hta", "")
        ExtractPlaylistId = id
    Else
        ExtractPlaylistId = ""
    End If
End Function
