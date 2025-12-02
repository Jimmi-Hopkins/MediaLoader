Option Explicit

' ==============================
'  links.vbs — автогенерация и проверка ссылок
' ==============================

Dim existingUrls()
Dim lastClipboardValue
Dim fso
Dim clipboardInterval
Public SettingsPlaylist


Set fso = CreateObject("Scripting.FileSystemObject")
ReDim existingUrls(-1)

' ------------------------------
' Инициализация
' ------------------------------
Sub InitUrlFields()
    On Error Resume Next
    Dim container
    Set container = Document.getElementById("urlContainer")

    ' СБРОСИТЬ отслеживание буфера при старте
	
    lastClipboardValue = ""
    
    ' ★★★ ЗАГРУЗКА СТАТУСОВ ИЗ ЛОГА ПРИ СТАРТЕ ★★★
    LoadFieldsFromMetadataLog()
    LoadStatusesFromMetadataLog()
    
	If container.children.length = 0 Then
        container.innerHTML = "<div style='color: #666; padding: 10px; text-align: center;'>Скопируйте ссылку в буфер обмена чтобы появился дополнительный интерфейс</div>"
    End If
	
    ' Загружаем настройки - мониторинг запустится внутри 
    LoadAutoCaptureSetting()
End Sub

' ------------------------------
' Загрузка настройки автоперехвата
' ------------------------------
Sub LoadAutoCaptureSetting()
    On Error Resume Next
    Dim settings, autoCaptureCheckbox
    Set autoCaptureCheckbox = Document.getElementById("autoCapture")
    
    settings = LoadSettings()
    If IsArray(settings) And UBound(settings) >= 7 Then
        autoCaptureCheckbox.Checked = (settings(7) = "true")
    Else
        autoCaptureCheckbox.Checked = False
    End If
    
    ' Запускаем мониторинг если включено
    If autoCaptureCheckbox.Checked Then
	ClearClipboard
        StartClipboardMonitoring()
    End If
End Sub

' ------------------------------
' Запуск мониторинга буфера
' ------------------------------
Sub StartClipboardMonitoring()
    On Error Resume Next
    ' Останавливаем предыдущий интервал
    If Not IsEmpty(clipboardInterval) Then
        window.clearInterval clipboardInterval
        clipboardInterval = Empty ' ← ВАЖНО: сбрасываем переменную
    End If
    
    lastClipboardValue = GetClipboardText()
    clipboardInterval = window.setInterval(GetRef("CheckClipboardChange"), 300)
End Sub

' ------------------------------
' Включение/выключение автоперехвата
' ------------------------------
Sub ToggleAutoCapture()
    On Error Resume Next
    Dim autoCaptureCheckbox
    Set autoCaptureCheckbox = Document.getElementById("autoCapture")
    
    ' Сразу сохраняем настройку
    SaveSettings()
    
    If autoCaptureCheckbox.Checked Then
        ClearClipboard
        lastClipboardValue = ""
        StartClipboardMonitoring()
    Else
        ' Останавливаем мониторинг
        If Not IsEmpty(clipboardInterval) Then
            window.clearInterval clipboardInterval
        End If
    End If
End Sub

' ------------------------------
' Сохранение настройки автоперехвата
' ------------------------------
Sub SaveAutoCaptureSetting(isEnabled)
    On Error Resume Next
    ' Просто вызываем общее сохранение настроек
    SaveSettings()
End Sub

' ------------------------------
' Автозагрузка полей из лога при старте (ОБНОВЛЕННАЯ)
' ------------------------------
Sub LoadFieldsFromMetadataLog()
    On Error Resume Next

    Dim fso, logFile, logPath, line, arr
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    If Not fso.FileExists(logPath) Then
        Exit Sub
    End If
       
    ' Очищаем контейнер если нужно
    Dim container
    Set container = Document.getElementById("urlContainer")
    If container.innerHTML Like "*Скопируйте ссылку в буфер обмена*" Then
        container.innerHTML = ""
    End If
    
    ' Читаем лог и восстанавливаем поля
    Set logFile = fso.OpenTextFile(logPath, 1)
    Do Until logFile.AtEndOfStream
        line = Trim(logFile.ReadLine)
        If line <> "" Then
            arr = Split(line, "|")
            If UBound(arr) >= 3 Then
                ' Восстанавливаем все поля кроме удаленных
                If arr(3) <> "removed" And arr(2) <> "" Then
                    If Not UrlExists(arr(2)) Then
                        RestoreUrlFieldFromLog arr(0), arr(2), arr(3)
                    End If
                End If
            End If
        End If
    Loop
    logFile.Close
End Sub

' ------------------------------
' Создание поля ссылки (ОБНОВЛЕННАЯ - поддерживает все статусы)
' ------------------------------
Sub RestoreUrlFieldFromLog(fieldId, url, status)
    On Error Resume Next
    Dim container, newDiv, html, domain
    
    Set container = Document.getElementById("urlContainer")

    Set newDiv = Document.createElement("div")
    newDiv.className = "url-block"
    newDiv.id = fieldId

    ' ★★★ РАЗДЕЛЯЕМ ОБРАБОТКУ ПО СТАТУСАМ ★★★
    If status = STATUS_ACTION Then
        ' Для ACTION статуса создаем поле с кнопками подтверждения
        domain = GetDomainFromUrl(url)
        html = "<input type='text' class='url-input action-required' value='" & url & "' " & _
               "style='color: red;' readonly>" & _
               " <span id='" & fieldId & "_status' title='Ссылка требует подтверждения'></span>" & _
               " <button data-fieldid='" & fieldId & "' data-save='false' onclick='VBScript:HandleConfirmClick()' title='Подтвердить добавление ссылки'>✔</button>" & _
               " <button data-fieldid='" & fieldId & "' data-save='true' onclick='VBScript:HandleConfirmClick()' title='Подтвердить и добавить сайт в список поддерживаемых'>💾</button>" & _
               " <button onclick='VBScript:RemoveUrlField(""" & fieldId & """)' title='Удалить ссылку'>🗑️</button>"
    ElseIf status = STATUS_PLAYLIST Then
        ' Для плейлистов
        html = "<input type='text' class='url-input' value='" & url & "' readonly>" & _
               " <span id='" & fieldId & "_status' title='Плейлист'>📓</span>" & _
               " <button data-fieldid='" & fieldId & "' onclick='VBScript:DownloadPlaylist(""" & fieldId & """)' title='Скачать плейлист'>📥</button>" & _
               " <button data-fieldid='" & fieldId & "' onclick='VBScript:saveEditPlaylist(""" & fieldId & """)' title='Редактировать плейлист'>✏️</button>" & _
               " <button onclick='VBScript:RemoveUrlField(""" & fieldId & """)' title='Удалить плейлист'>🗑️</button>"
    Else
        ' Для остальных статусов используем стандартную обработку + кнопка 📥
        html = ProcessSupportedUrl(fieldId, url)
    End If
    
    newDiv.innerHTML = html
    
    ' Добавляем в массив
    If url <> "" And Not UrlExists(url) Then
        ReDim Preserve existingUrls(UBound(existingUrls) + 1)
        existingUrls(UBound(existingUrls)) = url
    End If
    
    container.appendChild newDiv
    
    ' ★★★ ВОССТАНАВЛИВАЕМ СТАТУС ПОСЛЕ СОЗДАНИЯ HTML ★★★
    RestoreStatusInUI fieldId, status
End Sub

' ------------------------------
' Добавление нового поля (ОБНОВЛЕННАЯ)
' ------------------------------
Sub AddUrlField(url)
    On Error Resume Next
    Dim container, newDiv, fieldId, urlStatus, html, startStatus
    
    ' Проверяем ссылку
    urlStatus = IsSupportedUrl(url)
    If urlStatus = "invalid" Then Exit Sub
    
    Set container = Document.getElementById("urlContainer")
    
    fieldId = CLng(Timer * 10000)
    Set newDiv = Document.createElement("div")
    newDiv.className = "url-block"
    newDiv.id = fieldId

    ' РЕШАЕМ СТАРТОВЫЙ СТАТУС
    If urlStatus = "unsupported" Then
        startStatus = STATUS_ACTION
   
        html = "<input type='text' class='url-input action-required' value='" & url & "' " & _
               "style='color: red;' readonly>" & _
               " <span id='" & fieldId & "_status' title='Ссылка требует подтверждения'>❗</span>" & _
               " <button data-fieldid='" & fieldId & "' data-save='false' onclick='VBScript:HandleConfirmClick()' title='Добавить ссылку'>✔</button>" & _
               " <button data-fieldid='" & fieldId & "' data-save='true' onclick='VBScript:HandleConfirmClick()' title='Добавить сайт в список поддерживаемых'>💾</button>" & _
               " <button onclick='VBScript:RemoveUrlField(""" & fieldId & """)' title='Удалить ссылку'>🗑️</button>"
        
    ElseIf urlStatus = "supported" Then
        
        ' Определяем: playlist или одиночная
        If IsPlaylistUrl(url) Then
            startStatus = STATUS_PLAYLIST
            html = "<input type='text' class='url-input' value='" & url & "' readonly>" & _
                   " <span id='" & fieldId & "_status' title='Плейлист'>📓</span>" & _
                   " <button data-fieldid='" & fieldId & "' onclick='VBScript:DownloadPlaylist(""" & fieldId & """)' title='Скачать плейлист'>📥</button>" & _
                   " <button data-fieldid='" & fieldId & "' onclick='VBScript:saveEditPlaylist(""" & fieldId & """)' title='Редактировать плейлист'>✏️</button>" & _
                   " <button onclick='VBScript:RemoveUrlField(""" & fieldId & """)' title='Удалить плейлист'>🗑️</button>"
        Else
            startStatus = STATUS_WAITING
            ' ★★★ ДОБАВЛЯЕМ КНОПКУ СРАЗУ ПРИ СОЗДАНИИ ПОЛЯ ★★★
            html = "<input type='text' class='url-input' value='" & url & "' " & _
                   "onchange='VBScript:CheckUrlStatus(""" & fieldId & """)'>" & _
                   " <span id='" & fieldId & "_status'></span>" & _
                   " <button onclick='VBScript:RemoveUrlField(""" & fieldId & """)' title='Удалить ссылку'>🗑️</button>" & _
                   " <button onclick='VBScript:RedownloadVideo(""" & fieldId & """)' title='Инидивидуальное скачивание'>📥</button>"
        End If
    End If
    
    ' Запись статуса один раз
    WriteToMetadataLog fieldId, url, startStatus

    ' Вставка HTML
    newDiv.innerHTML = html
    container.appendChild newDiv

    ' Добавляем в массив
    If url <> "" And Not UrlExists(url) Then
        ReDim Preserve existingUrls(UBound(existingUrls) + 1)
        existingUrls(UBound(existingUrls)) = url
        
        ' Обновляем статус рядом с полем
        If urlStatus <> "unsupported" Then
            CheckUrlStatus fieldId
        End If

        ' АВТОЗАГРУЗКА: только если обычное supported видео
        If startStatus = STATUS_WAITING Then
        
            Dim autoDownloadCheckbox, autoCaptureCheckbox
            Set autoDownloadCheckbox = Document.getElementById("autoDownload")
            Set autoCaptureCheckbox = Document.getElementById("autoCapture")

            If Not autoDownloadCheckbox Is Nothing And autoDownloadCheckbox.Checked And _
               Not autoCaptureCheckbox Is Nothing And autoCaptureCheckbox.Checked Then

                DownloadSingleVideo url, fieldId

            End If

        End If
    End If
End Sub

' ------------------------------
' Обработчики кликов (без параметров)
' ------------------------------
Sub HandleConfirmClick()
    On Error Resume Next
    Dim button, fieldId, saveDomain
    Set button = Window.Event.SrcElement
    fieldId = button.getAttribute("data-fieldid")
    saveDomain = (button.getAttribute("data-save") = "true")
    ConfirmUrlField fieldId, saveDomain
End Sub

Sub HandleRemoveClick()
    On Error Resume Next
    Dim button, fieldId, domain
    Set button = Window.Event.SrcElement
    fieldId = button.getAttribute("data-fieldid")
    domain = button.getAttribute("data-domain")
    RemoveUrlField fieldId
End Sub
' ------------------------------
' Подтверждение поля с действием
' ------------------------------
Sub ConfirmUrlField(fieldId, saveDomain)
    On Error Resume Next
    Dim el, inputEl, url, domain
    
    Set el = Document.getElementById(fieldId)
    If el Is Nothing Then Exit Sub
    
    Set inputEl = el.getElementsByTagName("input")(0)
    If inputEl Is Nothing Then Exit Sub
    
    url = Trim(inputEl.value)
    domain = GetDomainFromUrl(url)
    
    If saveDomain And domain <> "" Then
        ' подтверждения
        Dim userResponse
        userResponse = MsgBox("Доп. Проверка:" & vbCrLf & vbCrLf & _
                            domain & " будет добавлен в Ваш список разрешенных." & vbCrLf & _
                            "Правка Вашего списка: app\supportedsites.md" & vbCrLf & _
                            "под строкой ===user list=====", _
                            vbYesNo + vbInformation, "Подтверждение добавления домена")
        
        If userResponse = vbYes Then      
            AppendUserSite "supportedsites.md", domain    
        Else          
            Exit Sub
        End If
    End If
    
    ' Обновляем статус в логе
    Dim newStatus
    If IsPlaylistUrl(url) Then
        newStatus = STATUS_PLAYLIST
    Else
        newStatus = STATUS_WAITING
    End If
    
    ' ★★★ ОБНОВЛЯЕМ HTML С КНОПКОЙ ПОВТОРНОЙ ЗАГРУЗКИ ★★★
    el.innerHTML = "<input type='text' class='url-input' value='" & url & "' " & _
                   "onchange='VBScript:CheckUrlStatus(""" & fieldId & """)'>" & _
                   " <span id='" & fieldId & "_status'></span>" & _
                   " <button onclick='VBScript:RemoveUrlField(""" & fieldId & """)' title='Удалить ссылку'>🗑️</button>" & _
                   " <button onclick='VBScript:RedownloadVideo(""" & fieldId & """)' title='Инидивидуальное скачивание'>📥</button>"
    
    UpdateMetadataLogStatus fieldId, url, newStatus
    UpdateStatus fieldId, url, newStatus
End Sub


Function ProcessSupportedUrl(fieldId, url)
    On Error Resume Next
    Dim html, currentStatus, title, displayText
    
    ' Получаем текущий статус из metadata_history.log
    currentStatus = GetCurrentStatus(fieldId)
    
    ' ★★★ ПОЛУЧАЕМ TITLE ИЗ METADATA ★★★
    title = GetTitleFromMetadata(fieldId)
    
    ' ★★★ ВЫБИРАЕМ ЧТО ПОКАЗЫВАТЬ: TITLE ИЛИ URL ★★★
    If title <> "" Then
        displayText = title
    Else
        displayText = url
    End If
        
    ' Проверяем плейлист
    If IsPlaylistUrl(url) Then
        html = "<input type='text' class='url-input' value='" & displayText & "' readonly>" & _
               " <span id='" & fieldId & "_status' title='Плейлист'>📓</span>" & _
               " <button data-fieldid='" & fieldId & "' onclick='VBScript:DownloadPlaylist(""" & fieldId & """)' title='Скачать плейлист'>📥</button>" & _
               " <button data-fieldid='" & fieldId & "' onclick='VBScript:saveEditPlaylist(""" & fieldId & """)' title='Редактировать плейлист'>✏️</button>" & _
               " <button onclick='VBScript:RemoveUrlField(""" & fieldId & """)' title='Удалить плейлист'>🗑️</button>"
    Else
        html = "<input type='text' class='url-input' value='" & displayText & "' " & _
               "onchange='VBScript:CheckUrlStatus(""" & fieldId & """)'>" & _
               " <span id='" & fieldId & "_status'></span>" & _
               " <button onclick='VBScript:RemoveUrlField(""" & fieldId & """)' title='Удалить ссылку'>🗑️</button>"
        
        ' ★★★ ДОБАВЛЯЕМ КНОПКУ ПОВТОРНОЙ ЗАГРУЗКИ ДЛЯ ВСЕХ СТАТУСОВ ★★★
        If currentStatus = STATUS_WAITING Or currentStatus = STATUS_DOWNLOADING Or _
           currentStatus = STATUS_COMPLETED Or currentStatus = STATUS_ERROR Then
            html = html & " <button onclick='VBScript:RedownloadVideo(""" & fieldId & """)' title='Инидивидуальное скачивание'>📥</button>"
        End If
    End If
    
    ProcessSupportedUrl = html
End Function

' ★★★ ПОВТОРНОЕ СКАЧИВАНИЕ ★★★
Sub RedownloadVideo(fieldId)
    On Error Resume Next
    Dim url
    url = GetUrlFromMetadata(fieldId)
    
    If url <> "" Then
        ' Меняем статус на waiting и запускаем загрузку
        UpdateMetadataLogStatus fieldId, url, "waiting"
        UpdateStatus fieldId, url, "waiting"
        DownloadSingleVideo url, fieldId
    End If
End Sub

' ★★★ ПОЛУЧЕНИЕ ТЕКУЩЕГО СТАТУСА ИЗ METADATA ★★★
Function GetCurrentStatus(fieldId)
    On Error Resume Next
    Dim fso, logFile, logPath, line, arr
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    GetCurrentStatus = ""
    
    If fso.FileExists(logPath) Then
        Set logFile = fso.OpenTextFile(logPath, 1)
        Do Until logFile.AtEndOfStream
            line = Trim(logFile.ReadLine)
            If line <> "" Then
                arr = Split(line, "|")
                If UBound(arr) >= 3 Then
                    If arr(0) = fieldId Then
                        GetCurrentStatus = arr(3)  ' статус в 4-й колонке
                        Exit Do
                    End If
                End If
            End If
        Loop
        logFile.Close
    End If
End Function

' ★★★ ПОЛУЧЕНИЕ TITLE ИЗ METADATA ★★★
Function GetTitleFromMetadata(fieldId)
    On Error Resume Next
    Dim fso, logFile, logPath, line, arr
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    GetTitleFromMetadata = ""
    
    If fso.FileExists(logPath) Then
        Set logFile = fso.OpenTextFile(logPath, 1)
        Do Until logFile.AtEndOfStream
            line = Trim(logFile.ReadLine)
            If line <> "" Then
                arr = Split(line, "|")
                If UBound(arr) >= 4 Then
                    If arr(0) = fieldId Then
                        GetTitleFromMetadata = arr(4)  ' title в 5-й колонке
                        Exit Do
                    End If
                End If
            End If
        Loop
        logFile.Close
    End If
End Function

' ★★★ ПРОВЕРКА ПЛЕЙЛИСТА ★★★
Function IsPlaylistUrl(url)
    On Error Resume Next
    
    Dim u
    u = LCase(url)

    ' Универсальные ключевые слова
    Dim keys, k
    keys = Array( _
        "list=", "playlist", "playlists", _
        "album", "collection", "collections", _
        "index=", "set=", "/set/", "/sets/", _
        "/folder", "folder=", "/series", "series=" _
    )

    For Each k In keys
        If InStr(u, k) > 0 Then
            IsPlaylistUrl = True
            Exit Function
        End If
    Next

    IsPlaylistUrl = False
End Function

' ★★★ СКАЧАТЬ ПЛЕЙЛИСТ ★★★
Sub DownloadPlaylist(fieldId)
    On Error Resume Next
    Dim el, inputEl, url
    Set el = Document.getElementById(fieldId)
    If el Is Nothing Then Exit Sub
    
    Set inputEl = el.getElementsByTagName("input")(0)
    url = Trim(inputEl.value)
    
    ' ★★★ ПОДТВЕРЖДЕНИЕ И СМЕНА СТАТУСА ★★★
    Dim userChoice
    userChoice = MsgBox("📓 Вы уверены, что хотите скачать весь плейлист?" & vbCrLf & vbCrLf & _
                        "Ссылка: " & url, vbYesNo + vbQuestion, "Подтверждение плейлиста")
    
    If userChoice = vbYes Then
        ' ★★★ МЕНЯЕМ СТАТУС НА WAITING ДЛЯ ЗАГРУЗКИ ★★★
        UpdateStatus fieldId, url, STATUS_WAITING
        DownloadSingleVideo url, fieldId
    End If
End Sub

' ★★★ РЕДАКТИРОВАТЬ ПЛЕЙЛИСТ ★★★
Sub saveEditPlaylist(fieldId)
    On Error Resume Next

    Dim el, inputEl, playlistUrl

    ' достаём URL плейлисты из поля
    Set el = Document.getElementById(fieldId)
    If el Is Nothing Then Exit Sub

    Set inputEl = el.getElementsByTagName("input")(0)
    If inputEl Is Nothing Then Exit Sub

    playlistUrl = Trim(inputEl.value)
    If playlistUrl = "" Then Exit Sub

	SettingsPlaylist = "true"
	Call SaveSettings()
	SettingsPlaylist = ""
    Call EditPlaylist(fieldId)

End Sub

' ------------------------------
' Удаление поля для невалидного домена (на перспективу)
' ------------------------------
'Sub RemoveActionField(fieldId, domain)
'    On Error Resume Next
'    ' Удаляем поле из DOM
'    RemoveUrlField fieldId
'    
'    ' Удаляем домен из пользовательского списка если есть
'    If domain <> "" Then
'        RemoveDomainFromUserList domain
'    End If
'End Sub

' ------------------------------
' Удаление домена из пользовательского списка
' ------------------------------
'Sub RemoveDomainFromUserList(domain)
'    On Error Resume Next
'    Dim fso, siteListPath, tempPath, logFile, tempFile, line, inUserSection
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    siteListPath = "supportedsites.md"
'    tempPath = "supportedsites.tmp"
'    
'    If Not fso.FileExists(siteListPath) Then Exit Sub
'    
'    Set logFile = fso.OpenTextFile(siteListPath, 1)
'    Set tempFile = fso.CreateTextFile(tempPath, True)
'    
'    inUserSection = False
'    Do Until logFile.AtEndOfStream
'        line = Trim(logFile.ReadLine)
'        
'        If line = "===user list=====" Then
'            inUserSection = True
'            tempFile.WriteLine line
'        ElseIf inUserSection Then
'            ' Пропускаем строку с этим доменом
'            If InStr(LCase(line), LCase(domain)) = 0 Then
'                tempFile.WriteLine line
'            End If
'        Else
'            tempFile.WriteLine line
'        End If
'    Loop
'    
'    logFile.Close
'    tempFile.Close
'    
'    ' Заменяем оригинальный файл
'    fso.DeleteFile siteListPath
'    fso.MoveFile tempPath, siteListPath
'End Sub

' ------------------------------
' Обновление статуса в логе метаданных
' ------------------------------
Sub UpdateMetadataLogStatus(fieldId, url, newStatus)
    On Error Resume Next
    Dim fso, logPath, tempPath, logFile, tempFile, line, arr
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    tempPath = "metadata_history.tmp"
    
    If Not fso.FileExists(logPath) Then Exit Sub
    
    Set logFile = fso.OpenTextFile(logPath, 1)
    Set tempFile = fso.CreateTextFile(tempPath, True)
    
    Do Until logFile.AtEndOfStream
        line = Trim(logFile.ReadLine)
        If line <> "" Then
            arr = Split(line, "|")
            If UBound(arr) >= 2 Then
                ' Обновляем строку с нужным fieldId
                If arr(0) = fieldId And arr(2) = url Then
                    arr(3) = newStatus
                    line = Join(arr, "|")
                End If
            End If
            tempFile.WriteLine line
        End If
    Loop
    
    logFile.Close
    tempFile.Close
    
    ' Заменяем оригинальный файл
    fso.DeleteFile logPath
    fso.MoveFile tempPath, logPath
End Sub

' ------------------------------
' Функция записи в лог метаданных (ЧИСТОВАЯ)
' ------------------------------
Sub WriteToMetadataLog(fieldId, url, status)
    On Error Resume Next
    Dim fso, logFile, logPath, timestamp
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    timestamp = Now()
    
    ' Безопасное получение значений
    Dim savePath, defaultFormat, defaultQuality, proxy, subtitles, embeddedSubs, detectedBrowser
    
    savePath = Document.getElementById("savePath").value
    If savePath = "" Then savePath = "."
    
    defaultFormat = Document.getElementById("defaultFormat").value  
    If defaultFormat = "" Then defaultFormat = "mp4"
    
    defaultQuality = Document.getElementById("defaultQuality").value
    If defaultQuality = "" Then defaultQuality = "max"
    
    proxy = GetProxyAddress()
    If proxy = "" Then proxy = "none"
    
    subtitles = Document.getElementById("subtitles").value
    If subtitles = "" Then subtitles = "none"
    
    embeddedSubs = Document.getElementById("embeddedSubs").Checked
    If embeddedSubs Then embeddedSubs = "True" Else embeddedSubs = "False"
    
    ' Формируем строку лога
    Dim logEntry
    logEntry = fieldId & "|" & timestamp & "|" & url & "|" & status & "|||" & _
               savePath & "|" & defaultFormat & "|" & defaultQuality & "|" & _
               proxy & "|" & subtitles & "|" & embeddedSubs & "|" & Split(authBrowserStatus.innerText, " ")(0)
    
    Set logFile = fso.OpenTextFile(logPath, 8, True)
    logFile.WriteLine logEntry
    logFile.Close
End Sub

' ------------------------------
' Удаление поля
' ------------------------------
Sub RemoveUrlField(fieldId)
    On Error Resume Next
    Dim el, inputEl, url, i, j
    Set el = Document.getElementById(fieldId)
    
    If Not el Is Nothing Then
        ' Находим URL для удаления
        Set inputEl = el.getElementsByTagName("input")(0)
        If Not inputEl Is Nothing Then
            url = Trim(inputEl.value)
            
            ' Удаляем из массива
            If url <> "" Then
                For i = 0 To UBound(existingUrls)
                    If LCase(existingUrls(i)) = LCase(url) Then
                        ' Сдвигаем массив
                        For j = i To UBound(existingUrls) - 1
                            existingUrls(j) = existingUrls(j + 1)
                        Next
                        If UBound(existingUrls) > 0 Then
                            ReDim Preserve existingUrls(UBound(existingUrls) - 1)
                        Else
                            ReDim existingUrls(-1)
                        End If
                        Exit For
                    End If
                Next
            End If
            
            ' Удаляем запись из metadata_history.log
            RemoveFromMetadataLog fieldId, url
            
            ' СБРАСЫВАЕМ отслеживание буфера для этой ссылки
            If lastClipboardValue = url Then
                ClearClipboard
                lastClipboardValue = ""
            End If
        End If
         
        ' Удаляем из DOM
        el.parentNode.removeChild el
        
        ' Если полей не осталось - показываем плейсхолдер
        Set container = Document.getElementById("urlContainer")
        If container.children.length = 0 Then
            container.innerHTML = "<div style='color: #666; padding: 10px; text-align: center;'>Скопируйте ссылку в буфер обмена чтобы появился дополнительный интерфейс</div>"
        End If
    End If
End Sub

' ------------------------------
' Удаление записи из metadata_history.log
' ------------------------------
Sub RemoveFromMetadataLog(fieldId, url)
    On Error Resume Next
    Dim fso, logPath, tempPath, logFile, tempFile, line, arr
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    tempPath = "metadata_history.tmp"
    
    If Not fso.FileExists(logPath) Then Exit Sub
    
    Set logFile = fso.OpenTextFile(logPath, 1) ' 1 = ForReading
    Set tempFile = fso.CreateTextFile(tempPath, True) ' True = Overwrite
    
    Do Until logFile.AtEndOfStream
        line = Trim(logFile.ReadLine)
        If line <> "" Then
            arr = Split(line, "|")
            ' Пропускаем запись с этим fieldId и URL
            If UBound(arr) >= 2 Then
                If arr(0) <> fieldId And arr(2) <> url Then
                    tempFile.WriteLine line
                End If
            End If
        End If
    Loop
    
    logFile.Close
    tempFile.Close
    
    ' Заменяем оригинальный файл временным
    fso.DeleteFile logPath
    fso.MoveFile tempPath, logPath
End Sub

' ------------------------------
' Очистка всех полей и метаданных
' ------------------------------
Sub ClearAllFields()
    On Error Resume Next
    
    ' ★★★ ПОДТВЕРЖДЕНИЕ ★★★
    Dim userResponse
    userResponse = MsgBox("Вы уверены, что хотите удалить ВСЕ ссылки и очистить историю загрузок?" & vbCrLf & vbCrLf & _
                         "Это действие нельзя отменить!", _
                         vbYesNo + vbExclamation, "Подтверждение очистки")
    
    If userResponse <> vbYes Then Exit Sub
    
    Dim i, container
    Set container = Document.getElementById("urlContainer")
    
    ' Очищаем контейнер
    container.innerHTML = "<div style='color: #666; padding: 10px; text-align: center;'>Скопируйте ссылку в буфер обмена чтобы появился дополнительный интерфейс</div>"
    
    ' Очищаем массив
    ReDim existingUrls(-1)
    
    ' Очищаем буфер обмена
    ClearClipboard
    
    ' СБРОСИТЬ отслеживание буфера
    lastClipboardValue = ""
    
    ' Очищаем metadata_history.log
    ClearMetadataLog
    
    ShowTempMessage "✅ Все ссылки и история очищены!"
End Sub

' ------------------------------
' Очистка лога метаданных
' ------------------------------
Sub ClearMetadataLog()
    On Error Resume Next
    Dim fso, logPath
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    If fso.FileExists(logPath) Then
        fso.DeleteFile logPath
    End If
    
End Sub

' ------------------------------
' Очистка буфера обмена
' ------------------------------
Sub ClearClipboard()
    On Error Resume Next
    Dim htmlFile
    Set htmlFile = CreateObject("htmlfile")
    htmlFile.ParentWindow.ClipboardData.SetData "text", ""
End Sub

' ------------------------------
' Проверка статуса ссылки
' ------------------------------
Sub CheckUrlStatus(fieldId)
    On Error Resume Next
    Dim el, inputEl, statusEl, url
    
    Set el = Document.getElementById(fieldId)
    If el Is Nothing Then Exit Sub
    
    Set inputEl = el.getElementsByTagName("input")(0)
    If inputEl Is Nothing Then Exit Sub
    
    ' Ищем span элементами
    Dim allSpans, i
    Set allSpans = el.getElementsByTagName("span")
    For i = 0 To allSpans.length - 1
        If allSpans(i).id = fieldId & "_status" Then
            Set statusEl = allSpans(i)
            Exit For
        End If
    Next
    
    If statusEl Is Nothing Then Exit Sub

    url = Trim(inputEl.value)
	
	  ' ★★★ Плейлисты НЕ переопределяем ★★★
	   If IsPlaylistUrl(url) Then
        UpdateStatus fieldId, url, STATUS_PLAYLIST
        Exit Sub
    End If
	
    If url = "" Then
        UpdateStatus fieldId, url, STATUS_ERROR
        Exit Sub
    End If

    ' Проверяем валидность ссылки
    Dim urlStatus
    urlStatus = IsSupportedUrl(url)
    
    If urlStatus = "invalid" Then
        UpdateStatus fieldId, url, STATUS_ERROR
    ElseIf urlStatus = "unsupported" Then
        UpdateStatus fieldId, url, STATUS_ACTION
 ElseIf urlStatus = "playlist" Then
     UpdateStatus fieldId, url, STATUS_PLAYLIST
    Else
        ' Для supported - проверяем существование файла
        Dim savePath, baseName, filePath
        savePath = Document.getElementById("savePath").value
        If savePath = "" Then savePath = "."
        
        baseName = Replace(url, "https://", "")
        baseName = Replace(baseName, "http://", "")
        baseName = Replace(baseName, "/", "_")
        baseName = Replace(baseName, "?", "_")
        baseName = Replace(baseName, "&", "_")
        baseName = Left(baseName, 100)
        
        filePath = fso.BuildPath(savePath, baseName & ".mp4")
        
        If fso.FileExists(filePath) Then
            UpdateStatus fieldId, url, STATUS_COMPLETED
        Else
            UpdateStatus fieldId, url, STATUS_WAITING
        End If
    End If
End Sub

' ------------------------------
' Проверка уникальности
' ------------------------------
Function UrlExists(url)
    On Error Resume Next
    Dim i
    UrlExists = False
    For i = 0 To UBound(existingUrls)
        If LCase(existingUrls(i)) = LCase(url) Then
            UrlExists = True
            Exit For
        End If
    Next
End Function

' ------------------------------
' Мониторинг буфера обмена (ОБНОВЛЕННЫЙ)
' ------------------------------
Sub CheckClipboardChange()
    On Error Resume Next
    Dim autoCaptureCheckbox
    Set autoCaptureCheckbox = Document.getElementById("autoCapture")
    
    ' ПРОВЕРЯЕМ СОСТОЯНИЕ ГАЛОЧКИ НАПРЯМУЮ
    If Not autoCaptureCheckbox.Checked Then Exit Sub
    
    Dim currentClipboard
    currentClipboard = GetClipboardText()

    If currentClipboard <> "" And currentClipboard <> lastClipboardValue Then
        ' Проверяем что это HTTP/HTTPS ссылка
        If Left(LCase(currentClipboard), 7) = "http://" Or Left(LCase(currentClipboard), 8) = "https://" Then
            If Not UrlExists(currentClipboard) Then
                ' ВСЕГДА создаем новое поле для новой ссылки
                AddUrlField currentClipboard
            End If
        End If
        lastClipboardValue = currentClipboard
    End If
End Sub

' ------------------------------
' Получение текста из буфера
' ------------------------------
Function GetClipboardText()
    On Error Resume Next
    Dim htmlFile, clip
    Set htmlFile = CreateObject("htmlfile")
    GetClipboardText = htmlFile.ParentWindow.ClipboardData.GetData("text")
End Function

' ------------------------------
' Проверка ссылки на валидность
' ------------------------------
Function IsSupportedUrl(url)
    On Error Resume Next
	
    Dim fso, file, line, domain, supported, userInput, siteListPath
    siteListPath = "supportedsites.md"
    supported = False
    
    ' Проверяем что это действительно URL (начинается с http)
    If Not (Left(LCase(url), 7) = "http://" Or Left(LCase(url), 8) = "https://") Then
        IsSupportedUrl = "invalid"
        Exit Function
    End If

    ' Извлекаем домен из ссылки
    domain = GetDomainFromUrl(url)
    If domain = "" Then
        IsSupportedUrl = "invalid"
        Exit Function
    End If

    ' Читаем список поддерживаемых доменов
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(siteListPath) Then
        Set file = fso.OpenTextFile(siteListPath, 1)
        Do Until file.AtEndOfStream
            line = LCase(file.ReadLine)
            If InStr(line, LCase(domain)) > 0 Then
                supported = True
                Exit Do
            End If
        Loop
        file.Close
    End If

    If supported Then
        IsSupportedUrl = "supported"
    Else
        IsSupportedUrl = "unsupported"
    End If
End Function

' ------------------------------
' Извлечение домена из URL
' ------------------------------
Function GetDomainFromUrl(url)
    On Error Resume Next
    Dim matches, regex
    Set regex = New RegExp
    regex.Pattern = "https?://([^/]+)/?"
    regex.IgnoreCase = True
    If regex.Test(url) Then
        Set matches = regex.Execute(url)
        GetDomainFromUrl = matches(0).SubMatches(0)
    Else
        GetDomainFromUrl = ""
    End If
End Function

' ------------------------------
' Добавление сайта в пользовательский список
' ------------------------------
Sub AppendUserSite(siteListPath, domain)
    On Error Resume Next
    Dim fso, file, text, entry
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Формат записи в стиле yt-dlp
    entry = " - **" & domain & "**" & vbCrLf

    If fso.FileExists(siteListPath) Then
        Set file = fso.OpenTextFile(siteListPath, 1)
        text = file.ReadAll
        file.Close
    Else
        text = ""
    End If

    ' Проверяем наличие блока ===user list=====
    If InStr(text, "===user list=====") = 0 Then
        text = text & vbCrLf & "===user list=====" & vbCrLf
    End If

    ' Добавляем новый домен
    text = text & entry

    ' Сохраняем файл
    Set file = fso.OpenTextFile(siteListPath, 2, True)
    file.Write text
    file.Close
End Sub
' ======= ЕДИНАЯ СИСТЕМА СТАТУСОВ =======
Const STATUS_WAITING     = "waiting"
Const STATUS_DOWNLOADING = "downloading" 
Const STATUS_COMPLETED   = "completed"
Const STATUS_ERROR       = "error"
Const STATUS_ACTION      = "action"
Const STATUS_PLAYLIST    = "playlist"

Const ICON_WAITING       = "🟡"
Const ICON_DOWNLOADING   = "⏳"
Const ICON_COMPLETED     = "✅"
Const ICON_ERROR         = "❌"
Const ICON_ACTION        = "❗"
Const ICON_PLAYLIST      = "📓"

' ------------------------------
' ОБНОВЛЕНИЕ СТАТУСА (универсальная функция)
' ------------------------------
Sub UpdateStatus(fieldId, url, newStatus)
    On Error Resume Next
 
    ' === Обновляем лог ===
    UpdateMetadataLogStatus CStr(fieldId), url, newStatus
    
    ' === ★★★ ПЕРЕСОЗДАЕМ HTML ЕСЛИ СТАТУС ИЗМЕНИЛСЯ НА ERROR/COMPLETED ★★★ ===
    Dim el
    Set el = Document.getElementById(fieldId)
    If Not el Is Nothing Then
        If newStatus = "completed" Or newStatus = "error" Then
            el.innerHTML = ProcessSupportedUrl(fieldId, url)
        End If
    End If
    
    ' === Определяем эмодзи и title для интерфейса ===
    Dim icon, statusTitle
    Select Case LCase(newStatus)
        Case STATUS_WAITING:     
            icon = ICON_WAITING
            statusTitle = "Ожидает загрузки"
        Case STATUS_DOWNLOADING: 
            icon = ICON_DOWNLOADING
            statusTitle = "Загружается..."
        Case STATUS_COMPLETED:   
            icon = ICON_COMPLETED
            statusTitle = "Загрузка завершена/файл существует"
        Case STATUS_ERROR:       
            icon = ICON_ERROR
            statusTitle = "ОШИБКА загрузки" & vbCrLf & _
                         "Решение:" & vbCrLf & _
                         "• Используйте прокси/VPN" & vbCrLf & _
                         "• Для прямых эфиров - дождитесь обработки YouTube" & vbCrLf & _
                         "• Проверьте доступность видео" & vbCrLf & _
                         "• Проверьте правильность ссылок"
        Case STATUS_ACTION:      
            icon = ICON_ACTION
            statusTitle = "Требуется подтверждение"
        Case STATUS_PLAYLIST:    
            icon = ICON_PLAYLIST
            statusTitle = "Плейлист"
        Case Else:               
            icon = "❔"
            statusTitle = "Неизвестный статус"
    End Select
    
    ' === Обновляем отображение на форме ===
    Dim statusEl
    Set statusEl = Document.getElementById(fieldId & "_status")
    If Not statusEl Is Nothing Then
        statusEl.innerText = icon
        statusEl.title = statusTitle
    End If
End Sub

' ------------------------------
' ЗАГРУЗКА СТАТУСОВ ПРИ СТАРТЕ
' ------------------------------
Sub LoadStatusesFromMetadataLog()
    On Error Resume Next
    Dim fso, logFile, logPath, line, arr, fieldId, url, status
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    If Not fso.FileExists(logPath) Then Exit Sub
    
    Set logFile = fso.OpenTextFile(logPath, 1)
    Do Until logFile.AtEndOfStream
        line = Trim(logFile.ReadLine)
        If line <> "" Then
            arr = Split(line, "|")
            If UBound(arr) >= 3 Then
                fieldId = arr(0)
                url = arr(2)  
                status = arr(3)
                
                ' Восстанавливаем статус в интерфейсе
                RestoreStatusInUI fieldId, status
				' Если это плейлист — восстановить HTML кнопок
If status = STATUS_PLAYLIST Then
    Call RestorePlaylistUI(fieldId)
End If

            End If
        End If
    Loop
    logFile.Close
End Sub
Sub RestorePlaylistUI(fieldId)
    On Error Resume Next

    Dim el, url

    Set el = Document.getElementById(fieldId)
    If el Is Nothing Then Exit Sub

    ' Получаем текущий URL
    url = el.getElementsByTagName("input")(0).value

    ' Пересобираем HTML плейлиста
    el.innerHTML = _
        "<input type='text' class='url-input' value='" & url & "' readonly>" & _
        " <span id='" & fieldId & "_status' title='Плейлист'>📓</span>" & _
        " <button data-fieldid='" & fieldId & "' onclick='VBScript:DownloadPlaylist(""" & fieldId & """)' title='Скачать весь плейлист'>📥</button>" & _
        " <button data-fieldid='" & fieldId & "' onclick='VBScript:saveEditPlaylist(""" & fieldId & """)' title='Редактировать плейлист'>✏️</button>" & _
        " <button onclick='VBScript:RemoveUrlField(""" & fieldId & """)' title='Удалить плейлист'>🗑️</button>"
End Sub

' ------------------------------
' ВОССТАНОВЛЕНИЕ СТАТУСА В ИНТЕРФЕЙСЕ С TITLE
' ------------------------------
Sub RestoreStatusInUI(fieldId, status)
    On Error Resume Next
    Dim statusEl, icon, statusTitle
    
    ' Определяем эмодзи и title по статусу
    Select Case LCase(status)
        Case STATUS_WAITING:     
            icon = ICON_WAITING
            statusTitle = "Ожидает загрузки"
        Case STATUS_DOWNLOADING: 
            icon = ICON_DOWNLOADING
            statusTitle = "Загружается..."
        Case STATUS_COMPLETED:   
            icon = ICON_COMPLETED
            statusTitle = "Загрузка завершена"
        Case STATUS_ERROR:       
            icon = ICON_ERROR
            statusTitle = "ОШИБКА загрузки" & vbCrLf & _
                         "Решение:" & vbCrLf & _
                         "• Используйте прокси/VPN" & vbCrLf & _
                         "• Для прямых эфиров - дождитесь обработки YouTube" & vbCrLf & _
                         "• Проверьте доступность видео" & vbCrLf & _
                         "• Проверьте правильность ссылок"
        Case STATUS_ACTION:      
            icon = ICON_ACTION
            statusTitle = "Требуется подтверждение домена"
        Case STATUS_PLAYLIST:    
            icon = ICON_PLAYLIST
            statusTitle = "Плейлист"
        Case Else:               
            icon = "❔"
            statusTitle = "Неизвестный статус"
    End Select
    
    ' Обновляем элемент интерфейса
    Set statusEl = Document.getElementById(fieldId & "_status")
    If Not statusEl Is Nothing Then
        statusEl.innerText = icon
        statusEl.title = statusTitle  ' ★★★ ДОБАВЛЯЕМ TITLE ★★★
    End If
End Sub

' ------------------------------
' ПОИСК FIELDID ПО URL (для массовой загрузки)
' ------------------------------
Function FindFieldIdByUrl(url)
    On Error Resume Next
    Dim fso, logFile, logPath, line, arr
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = "metadata_history.log"
    
    FindFieldIdByUrl = ""
    
    If fso.FileExists(logPath) Then
        Set logFile = fso.OpenTextFile(logPath, 1)
        Do Until logFile.AtEndOfStream
            line = Trim(logFile.ReadLine)
            If line <> "" Then
                arr = Split(line, "|")
                If UBound(arr) >= 3 Then
                    If Trim(arr(2)) = url Then
                        FindFieldIdByUrl = arr(0)
                        Exit Do
                    End If
                End If
            End If
        Loop
        logFile.Close
    End If
End Function

