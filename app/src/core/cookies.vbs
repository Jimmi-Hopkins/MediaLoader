' =========================
' cookies.vbs (автоопределение профиля)
' =========================
Option Explicit

Const BROWSER_PROFILES_FILE = "config\browser_profiles.txt"

Sub StartAutoAuth()
    Dim shell, fso, resultPath, testUrl

    Set shell = CreateObject("WScript.Shell")
    Set fso   = CreateObject("Scripting.FileSystemObject")
    
    testUrl = InputBox("Введите ссылку с ОГРАНИЧЕННЫМ ДОСТУПОМ для проверки авторизации:")
    If Trim(testUrl) = "" Then Exit Sub

    resultPath = "temp\auth_result.txt"

    ' Удаляем перед запуском
    If fso.FileExists(resultPath) Then fso.DeleteFile resultPath, True

    ' Запуск батника
    shell.Run "bin\auth_check.bat " & testUrl, 1, True

    ' Читаем результат
    Dim found, tf, line
   found = ""

    If fso.FileExists(resultPath) Then
        Set tf = fso.OpenTextFile(resultPath, 1)
        If Not tf.AtEndOfStream Then
            line = Trim(tf.ReadLine)
            If line <> "" Then found = line
        End If
        tf.Close
    Else
        found = "0"
    End If

  Dim statusEl
Set statusEl = Document.getElementById("authBrowserStatus")

If found <> "0" And found <> "" Then
    ' Получаем состояние чекбокса авторизации
    Dim authCheckbox
    Set authCheckbox = Document.getElementById("useBrowserAuth")
    
    If Not authCheckbox Is Nothing And authCheckbox.Checked Then
        statusEl.innerText = found & " вкл "
        statusEl.style.color = "lime"
    Else
        statusEl.innerText = found & " выкл" 
        statusEl.style.color = "red"
    End If
    
    ' Сохраняем только если браузер найден
    SaveDetectedBrowser found
Else
    statusEl.innerText = "Не авторизован"
    statusEl.style.color = "red"
    ' ★★★ СБРАСЫВАЕМ БРАУЗЕР ПРИ ОШИБКЕ ★★★
    detectedBrowser = ""
    SaveDetectedBrowser ""  ' Сохраняем пустую строку
End If

Set shell = Nothing
Set fso = Nothing
End Sub
' ================================
'  Сохранение найденного профиля
' ================================
Sub SaveDetectedBrowser(browserName)
    On Error Resume Next

    Dim fso, settingsPath, txt, arr, i
    Set fso = CreateObject("Scripting.FileSystemObject")
    settingsPath = fso.BuildPath(fso.GetParentFolderName(window.location.pathname), _
                                 "config\downloader_settings.txt")

    If Not fso.FileExists(settingsPath) Then Exit Sub

    Dim f
    Set f = fso.OpenTextFile(settingsPath, 1)
    txt = f.ReadAll
    f.Close

    arr = Split(txt, "|")
    ' 7-й параметр = detectedBrowser
    If UBound(arr) < 6 Then ReDim Preserve arr(6)

    arr(6) = browserName  ' ★★★ ТЕПЕРЬ ЗДЕСЬ МОЖЕТ БЫТЬ ПУСТАЯ СТРОКА ★★★

    Set f = fso.OpenTextFile(settingsPath, 2, True)
    f.Write Join(arr, "|")
    f.Close
    
    detectedBrowser = browserName  ' ★★★ ОБНОВЛЯЕМ ГЛОБАЛЬНУЮ ПЕРЕМЕННУЮ ★★★
End Sub
