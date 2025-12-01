' Вспомогательные функции
'

' Отображение временного сообщения
Sub ShowTempMessage(message)
    On Error Resume Next
    Dim popup
    Set popup = Document.getElementById("tempPopup")
    popup.innerHTML = message
    popup.style.display = "block"
    
    ' Автоматически определяем цвет по содержанию
    If InStr(message, "❌") > 0 Or InStr(message, "отменён") > 0 Then
        popup.style.background = "#8B0000"
        popup.style.border = "1px solid #A00000"
    Else
        popup.style.background = "#2d5b4d"
        popup.style.border = "1px solid #3a7a65"
    End If
    
    window.setTimeout "HideTempMessage()", 2000
End Sub

' Скрытие временного сообщения
Sub HideTempMessage()
    On Error Resume Next
    Document.getElementById("tempPopup").style.display = "none"
End Sub

' Проверка placeholder прокси
Sub CheckProxyPlaceholder()
    On Error Resume Next
    Dim proxyInput
    Set proxyInput = Document.getElementById("proxy")
    
    ' Проверяем что поле пустое И НЕ содержит placeholder-текст
    If proxyInput.Value = "" And proxyInput.className <> "proxy-empty" Then
        proxyInput.className = "proxy-empty"
        proxyInput.Value = "http://ip:port или http://логин:пароль@ip:port"
    End If
End Sub

' Очистка placeholder прокси
Sub ClearProxyPlaceholder()
    On Error Resume Next
    Dim proxyInput
    Set proxyInput = Document.getElementById("proxy")
    
    If proxyInput.className = "proxy-empty" Then
        proxyInput.Value = ""
        proxyInput.className = ""
    End If
End Sub

' Очистка всех полей
Sub ClearAllFields()
    Dim i
    For i = 1 To 5
        Document.getElementById("url" & i).Value = ""
        Document.getElementById("status" & i).innerHTML = ""
    Next
End Sub
 
Sub UpdateFiles()
    On Error Resume Next
    Dim shell, fso, currentDir, parentDir, userResponse
    
    ' Запрос подтверждения
    userResponse = MsgBox("Будет запущен процесс обновления." & vbCrLf & _
                         "Приложение будет перезапущено автоматически, а настройки сохранены." & vbCrLf & vbCrLf & _
                         "Продолжить?", vbYesNo + vbInformation, "Актуализировать файлы")
    
    If userResponse <> vbYes Then
        Exit Sub
    End If
    
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    currentDir = fso.GetParentFolderName(window.location.pathname) ' app\
    parentDir = fso.GetParentFolderName(currentDir) ' корневая папка (где update.bat)
    
    ' Ищем update.bat в корневой папке (на уровень выше app)
    If fso.FileExists(fso.BuildPath(parentDir, "update.bat")) Then
        shell.Run Chr(34) & fso.BuildPath(parentDir, "update.bat") & Chr(34), 1, False
    Else
        MsgBox "Файл обновления update.bat не найден!" & vbCrLf & _
               "Ожидаемый путь: " & vbCrLf & _
               fso.BuildPath(parentDir, "update.bat"), vbExclamation
    End If
    
    ExitApp()
End Sub

Sub Authorization_help()
    On Error Resume Next
    Dim shell, fso, currentDir, parentDir, userResponse
        
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    currentDir = fso.GetParentFolderName(window.location.pathname) ' app\
    parentDir = fso.GetParentFolderName(currentDir) 
    
    ' Ищем Authorization_help.md в корневой папке (на уровень выше app)
    If fso.FileExists(fso.BuildPath(parentDir, "update.exe")) Then
        shell.Run Chr(34) & fso.BuildPath(parentDir, "Authorization_help.md") & Chr(34), 1, False
    Else
        MsgBox "Файл Authorization_help.md не найден!" & vbCrLf & _
               "Ожидаемый путь: " & vbCrLf & _
               fso.BuildPath(parentDir, "update.exe"), vbExclamation
    End If

End Sub

Sub ShowSupportedSites()
    On Error Resume Next
    Dim shell, fso, currentDir
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    currentDir = fso.GetParentFolderName(window.location.pathname) ' app\
    
    ' Ищем supportedsites.md в папке app (рядом с HTA)
    If fso.FileExists(fso.BuildPath(currentDir, "supportedsites.md")) Then
        shell.Run Chr(34) & fso.BuildPath(currentDir, "supportedsites.md") & Chr(34), 1, False
    Else
        MsgBox "Файл supportedsites.md не найден в папке с программой!", vbExclamation
    End If
End Sub

