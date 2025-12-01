' Модуль работы с прокси-серверами
'

' Получение адреса прокси с обработкой placeholder
Function GetProxyAddress()
    On Error Resume Next
    Dim proxy
    proxy = Trim(Document.getElementById("proxy").Value)
    
    ' Проверяем placeholder
    If proxy = "http://ip:port или http://логин:пароль@ip:port" Then
        GetProxyAddress = ""
        Exit Function
    End If
    
    ' Автоматическое добавление http:// если нет протокола
    If proxy <> "" Then
        If Left(LCase(proxy), 4) <> "http" And Left(LCase(proxy), 5) <> "socks" Then
            proxy = "http://" & proxy
        End If
    End If
    
    GetProxyAddress = proxy
End Function

' Тестирование прокси
Sub TestProxy()
    On Error Resume Next
    Dim proxy, shell, cmd, fso, currentDir, tempFile, resultFile, result
    Dim testUrl, i, foundUrl, testButton, tempPath, f, line, parts, btn

    ' Находим кнопку тестирования
    Set testButton = Nothing
    For Each btn In Document.getElementsByTagName("button")
        If InStr(btn.innerHTML, "Тест прокси") > 0 Then
            Set testButton = btn
            Exit For
        End If
    Next

    ' Меняем текст кнопки и блокируем её
    If Not testButton Is Nothing Then
        testButton.innerHTML = "⏳ Тестирую..."
        testButton.disabled = True
    End If

    proxy = GetProxyAddress()

    If proxy = "" Then
        ShowTempMessage "❌ Введите адрес прокси-сервера!"
        If Not testButton Is Nothing Then
            testButton.innerHTML = "Тест прокси"
            testButton.disabled = False
        End If
        Exit Sub
    End If

    ' Новый вариант: читаем первую ссылку из metadata_history.log
    foundUrl = ""
    Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("metadata_history.log") Then
    Set f = fso.OpenTextFile("metadata_history.log", 1, False)
    Do Until f.AtEndOfStream
        line = Trim(f.ReadLine)
        If line <> "" Then
            parts = Split(line, "|")
            If UBound(parts) >= 2 Then
                foundUrl = Trim(parts(2))
               MsgBox "Не закрывайте приложение до окончания теста!" & vbCrLf & _
       "🔍 Будет проверена ссылка:" & vbCrLf & _
       foundUrl & vbCrLf & vbCrLf & _
       "Нажмите OK для продолжения.", vbInformation, "Тест прокси"
                      
                Exit Do
            End If
        End If
    Loop
    f.Close
End If

' Если нет ссылок — используем тестовую
If foundUrl = "" Then
    foundUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ"
    MsgBox  "Не закрывайте приложение до окончания теста!" & vbCrLf & _
           "🔍 Будет проверена тестовая ссылка" & vbCrLf & vbCrLf & _
           "Нажмите ОК для продолжения.", vbInformation, "Тест прокси"
End If

    ' Создаём путь для лога
    currentDir = fso.GetParentFolderName(window.location.pathname)
    tempPath = fso.BuildPath(currentDir, "temp\logs\proxy_test_result.txt")

    ' Создаём папку для логов, если не существует
    If Not fso.FolderExists(fso.GetParentFolderName(tempPath)) Then
        fso.CreateFolder(fso.GetParentFolderName(tempPath))
    End If

    ' Запускаем команду для тестирования прокси
    Set shell = CreateObject("WScript.Shell")
    cmd = "cd /d " & Chr(34) & currentDir & Chr(34) & _
          " && bin\yt-dlp --proxy " & Chr(34) & proxy & Chr(34) & _
          " --get-title " & Chr(34) & foundUrl & Chr(34) & _
          " > " & Chr(34) & tempPath & Chr(34) & " 2>&1"

    shell.Run "cmd /c " & cmd, 0, True

' Проверяем результат
If fso.FileExists(tempPath) Then
    Set resultFile = fso.OpenTextFile(tempPath, 1)
    result = resultFile.ReadAll
    resultFile.Close
    fso.DeleteFile tempPath

    ' --- 🔍 Логика анализа результата ---
    If InStr(result, "Sign in") > 0 Or InStr(result, "not a bot") > 0 Then
        MsgBox "⚠️ YouTube запросил подтверждение, что вы не бот." & vbCrLf & _
               "Прокси, вероятно, рабочий, но YouTube ограничил доступ." & vbCrLf & _
               vbCrLf & "Рекомендации:" & vbCrLf & _
               "• Попробуйте другую ссылку (например, короткое видео)" & vbCrLf & _
			   "• Очистите все поля для использования тестовой ссылки" & vbCrLf & _
               "• Смените прокси / IP-адрес", vbExclamation, "Антибот YouTube"

    ElseIf InStr(result, "ERROR") = 0 And InStr(result, "unable") = 0 And InStr(result, "Cannot") = 0 Then
        If Len(Trim(result)) > 10 Then
            MsgBox "✅ Прокси РАБОТАЕТ отлично!" & vbCrLf & _
                   "Заголовок: " & Left(Trim(result), 100), vbInformation
            AddToProxyHistory proxy
        Else
            MsgBox "⚠️ Неожиданный результат:" & vbCrLf & result, vbInformation
        End If

    ElseIf InStr(result, "Unable to connect") > 0 Then
        MsgBox "❌ Прокси НЕ РАБОТАЕТ! Не удалось подключиться", vbExclamation

    ElseIf InStr(result, "407") > 0 Then
        MsgBox "❌ Ошибка авторизации прокси! Проверьте логин/пароль", vbExclamation

    ElseIf InStr(result, "403") > 0 Then
        MsgBox "❌ Прокси запретил доступ (403)", vbExclamation

    Else
        MsgBox "❌ Прокси НЕ РАБОТАЕТ! Проверьте ссылку." & vbCrLf & _
               Left(result, 200), vbExclamation
    End If

Else
    MsgBox "❌ Не удалось протестировать прокси. Файл результата не создан.", vbExclamation
End If


    ' Восстанавливаем кнопку
    If Not testButton Is Nothing Then
        testButton.innerHTML = "Тест прокси"
        testButton.disabled = False
    End If
End Sub


' Сохранение истории прокси
Sub SaveProxyHistory()
    On Error Resume Next
    Dim fso, historyFile, historyPath, history
    Set fso = CreateObject("Scripting.FileSystemObject")
    historyPath = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), "config\proxy_history.txt")
    
    history = GetProxyHistory()
    
    Set historyFile = fso.CreateTextFile(historyPath, True)
    historyFile.Write history
    historyFile.Close
End Sub

' Загрузка истории прокси
Sub LoadProxyHistory()
    On Error Resume Next
    Dim fso, historyFile, historyPath, history
    Set fso = CreateObject("Scripting.FileSystemObject")
    historyPath = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), "config\proxy_history.txt")
    
    If fso.FileExists(historyPath) Then
        Set historyFile = fso.OpenTextFile(historyPath, 1)
        history = historyFile.ReadAll
        historyFile.Close
        
        ' Убедимся что placeholder есть перед обновлением
        Dim historySelect
        Set historySelect = Document.getElementById("proxyHistory")
        If historySelect.Options.Length = 0 Or historySelect.Options(0).Value <> "--placeholder--" Then
            historySelect.innerHTML = "<option value=""--placeholder--"">-- История прокси --</option>"
        End If
        
        UpdateProxyDatalist(history)
    End If
End Sub

' Переключение отображения истории прокси
Sub ToggleProxyHistory()
    On Error Resume Next
    Dim proxyInput, historySelect
    Set proxyInput = Document.getElementById("proxy")
    Set historySelect = Document.getElementById("proxyHistory")
    
    If historySelect.style.display = "none" Then
        proxyInput.style.display = "none"
        historySelect.style.display = "inline-block"
        historySelect.focus()
        LoadProxyHistory()
    Else
        proxyInput.style.display = "inline-block"
        historySelect.style.display = "none"
        ' Убедимся что убрали placeholder стиль при возврате
        If proxyInput.Value <> "" And proxyInput.Value <> "http://ip:port или http://логин:пароль@ip:port" Then
            proxyInput.className = ""
        End If
    End If
End Sub

' Выбор прокси из истории
Sub SelectProxyFromHistory()
    On Error Resume Next
    Dim historySelect, proxyInput
    Set historySelect = Document.getElementById("proxyHistory")
    Set proxyInput = Document.getElementById("proxy")
    
    If historySelect.Value <> "" And historySelect.Value <> "--placeholder--" Then
        ' Убираем placeholder-стиль и вставляем значение
        proxyInput.Value = historySelect.Value
        proxyInput.className = ""  ' Убираем класс placeholder
        proxyInput.style.display = "inline-block"
        historySelect.style.display = "none"
        
        ' Сохраняем настройки сразу
        SaveSettings
    End If
End Sub

' Получение истории прокси
Function GetProxyHistory()
    Dim historySelect, i, history
    Set historySelect = Document.getElementById("proxyHistory")
    history = ""
    
    ' Получаем опции из select
    For i = 0 To historySelect.Options.Length - 1
        If historySelect.Options(i).Value <> "" And historySelect.Options(i).Value <> "--placeholder--" Then
            If history <> "" Then history = history & ","
            history = history & historySelect.Options(i).Value
        End If
    Next
    
    GetProxyHistory = history
End Function

' Обновление списка истории
Sub UpdateProxyDatalist(history)
    Dim proxyList, i, historySelect, optionElement
    Set historySelect = Document.getElementById("proxyHistory")
    
    ' Удаляем ВСЕ старые прокси (кроме placeholder)
    For i = historySelect.Options.Length - 1 To 1 Step -1
        historySelect.remove(i)
    Next
    
    ' Добавляем ТОЛЬКО новые прокси из переданной истории
    proxyList = Split(history, ",")
    For i = 0 To UBound(proxyList)
        If Trim(proxyList(i)) <> "" And Trim(proxyList(i)) <> "--placeholder--" Then
            Set optionElement = Document.createElement("option")
            optionElement.Value = Trim(proxyList(i))
            optionElement.innerHTML = Trim(proxyList(i))
            historySelect.appendChild(optionElement)
        End If
    Next
End Sub

' Добавление прокси в историю
Sub AddToProxyHistory(proxyAddress)
    On Error Resume Next
    Dim history, proxyList, i, exists
    history = GetProxyHistory()
    proxyList = Split(history, ",")
    exists = False
    
    ' Проверяем нет ли уже такого прокси в истории
    For i = 0 To UBound(proxyList)
        If LCase(Trim(proxyList(i))) = LCase(Trim(proxyAddress)) Then
            exists = True
            Exit For
        End If
    Next
    
    ' Добавляем если нет
    If Not exists Then
        If history = "" Then
            history = proxyAddress
        Else
            history = proxyAddress & "," & history
        End If
        
        ' Сохраняем только последние 10 прокси
        proxyList = Split(history, ",")
        If UBound(proxyList) > 9 Then
            history = ""
            For i = 0 To 9
                If i > 0 Then history = history & ","
                history = history & proxyList(i)
            Next
        End If
        
        ' Обновляем datalist
        UpdateProxyDatalist(history)
        ' Сохраняем историю
        SaveProxyHistory()
    End If
End Sub

' Очистка истории прокси
Sub ClearProxyHistory()
    If MsgBox("Очистить историю прокси?", vbYesNo + vbQuestion, "Подтверждение") = vbYes Then
        On Error Resume Next
        Dim fso, historyPath
        Set fso = CreateObject("Scripting.FileSystemObject")
        historyPath = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), "config\proxy_history.txt")
        
        If fso.FileExists(historyPath) Then
            fso.DeleteFile historyPath
        End If
        
        Document.getElementById("proxyHistory").innerHTML = "<option value=""--placeholder--"">-- История прокси --</option>"
        
    End If
End Sub
