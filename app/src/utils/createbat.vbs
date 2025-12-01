' createbat.vbs - скрипт инициализации MultiLoader (работает в HTA)
Option Explicit

Dim fso, shell, appPath, binPath, basePath

Sub InitializeBatFiles()
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    
    ' Определяем пути ОТ HTA ФАЙЛА
    appPath = Left(document.location.pathname, InStrRev(document.location.pathname, "\"))
    binPath = appPath & "bin\"
    basePath = fso.GetParentFolderName(appPath)
    
    ' 1. Проверяем и создаем BAT-файлы
    Dim needUpdate
    needUpdate = CheckAndCreateBats()
    
    ' 2. Если создавали батники ИЛИ нет EXE файлов - запускаем update.bat
    If needUpdate Or Not AllExeFilesExist() Then
        DownloadExeFiles
    End If
End Sub

Function CheckAndCreateBats()
    Dim createdAny
    createdAny = False
    
    ' 1. auth_check.bat
    If Not fso.FileExists(binPath & "auth_check.bat") Then
        CreateAuthCheckBat
        createdAny = True
    End If
    
    ' 2. cookies-from-browser.bat
    If Not fso.FileExists(binPath & "cookies-from-browser.bat") Then
        CreateCookiesBat
        createdAny = True
    End If
    
    ' 3. update.bat (на уровень выше app)
    If Not fso.FileExists(basePath & "\update.bat") Then
        CreateUpdateBat
        createdAny = True
    End If
    
    CheckAndCreateBats = createdAny
End Function

Function AllExeFilesExist()
    Dim exeFiles, exeFile
    exeFiles = Array("ffmpeg.exe", "ffplay.exe", "ffprobe.exe", "yt-dlp.exe")
    
    For Each exeFile in exeFiles
        If Not fso.FileExists(binPath & exeFile) Then
            AllExeFilesExist = False
            Exit Function
        End If
    Next
    
    AllExeFilesExist = True
End Function

Sub DownloadExeFiles()
    ' Запускаем update.bat для скачивания недостающих EXE файлов
    shell.Run "cmd /c """ & basePath & "\update.bat""", 1, True
    
    ' После завершения update.bat проверяем скачались ли файлы
    CheckExeFilesAfterUpdate
End Sub

Sub CheckExeFilesAfterUpdate()
    Dim exeFiles, missingExes, exeFile
    exeFiles = Array("ffmpeg.exe", "ffplay.exe", "ffprobe.exe", "yt-dlp.exe")
    missingExes = ""
    
    For Each exeFile in exeFiles
        If Not fso.FileExists(binPath & exeFile) Then
            If missingExes <> "" Then missingExes = missingExes & ", "
            missingExes = missingExes & exeFile
        End If
    Next
    
    If missingExes <> "" Then
        MsgBox "Не удалось скачать: " & missingExes & "." & vbCrLf & _
               "Запустите update.bat вручную для повторной попытки.", _
               vbExclamation, "Ошибка скачивания"
    Else
    ShowTempMessage "✅ Все компоненты установлены. Комфортной работы!"
    End If
End Sub

' ==================== СОЗДАНИЕ BAT-ФАЙЛОВ ====================
Sub CreateAuthCheckBat()
    Dim content
    content = "@echo off" & vbCrLf & _
              "chcp 65001 >nul" & vbCrLf & _
              "" & vbCrLf & _
              "set " & Chr(34) & "URL=%1" & Chr(34) & vbCrLf & _
              "set " & Chr(34) & "BASE=%2" & Chr(34) & "  ← получаем BASE из параметра" & vbCrLf & _
              "" & vbCrLf & _
              "rem === Настройка путей ===" & vbCrLf & _
              "for %%I in (" & Chr(34) & "%~dp0.." & Chr(34) & ") do set " & Chr(34) & "BASE=%%~fI" & Chr(34) & vbCrLf & _
              "set " & Chr(34) & "URL=%1" & Chr(34) & vbCrLf & _
              "set " & Chr(34) & "RESULT=%BASE%\temp\auth_result.txt" & Chr(34) & vbCrLf & _
              "set " & Chr(34) & "PROFILES=%BASE%\config\browser_profiles.txt" & Chr(34) & vbCrLf & _
              "set " & Chr(34) & "YT=%~dp0yt-dlp.exe" & Chr(34) & vbCrLf & _
              "echo." & vbCrLf & _
              "" & vbCrLf & _
              "rem === Удаляем старый результат ===" & vbCrLf & _
              "" & vbCrLf & _
              "if exist " & Chr(34) & "%RESULT%" & Chr(34) & " (" & vbCrLf & _
              "     del " & Chr(34) & "%RESULT%" & Chr(34) & vbCrLf & _
              ")" & vbCrLf & _
              "rem === Основной цикл ===" & vbCrLf & _
              "echo Ищем подходящий профиль...." & vbCrLf & _
              "setlocal enabledelayedexpansion" & vbCrLf & _
              "for /f " & Chr(34) & "usebackq tokens=* delims=" & Chr(34) & " %%P in (" & Chr(34) & "%PROFILES%" & Chr(34) & ") do (" & vbCrLf & _
              "    set " & Chr(34) & "PROFILE=%%P" & Chr(34) & vbCrLf & _
              "    if not " & Chr(34) & "!PROFILE!" & Chr(34) & "==" & Chr(34) & Chr(34) & " (" & vbCrLf & _
              "        echo Проверяем: !PROFILE!" & vbCrLf & _
              "        " & Chr(34) & "%YT%" & Chr(34) & " --cookies-from-browser " & Chr(34) & "!PROFILE!" & Chr(34) & " --get-title " & Chr(34) & "%URL%" & Chr(34) & " 2>nul" & vbCrLf & _
              "        if !ERRORLEVEL! EQU 0 (" & vbCrLf & _
              "            echo SUCCESS: Writing to result file" & vbCrLf & _
              "            echo !PROFILE! > " & Chr(34) & "%RESULT%" & Chr(34) & vbCrLf & _
              "            exit /b 0" & vbCrLf & _
              "        ) else (" & vbCrLf & _
              "            echo Не удалось: !PROFILE!" & vbCrLf & _
              "        )" & vbCrLf & _
              "    )" & vbCrLf & _
              ")" & vbCrLf & _
              "" & vbCrLf & _
              "echo 0 > " & Chr(34) & "%RESULT%" & Chr(34) & vbCrLf & _
              "echo." & vbCrLf & _
              "echo ===============================================" & vbCrLf & _
              "echo           АВТОРИЗАЦИЯ НЕ НАЙДЕНА" & vbCrLf & _
              "echo ===============================================" & vbCrLf & _
              "echo." & vbCrLf & _
              "echo ВОЗМОЖНЫЕ ПРИЧИНЫ:" & vbCrLf & _
              "echo 1. Браузер с авторизацией не полностью закрыт" & vbCrLf & _
              "echo 2. Браузер работает в фоновом режиме" & vbCrLf & _
              "echo 3. Вы не авторизовались в YouTube" & vbCrLf & _
              "echo." & vbCrLf & _
              "echo ВАЖНО:" & vbCrLf & _
              "echo - Закройте браузер ПOЛНОСТЬЮ (даже из системного трея)" & vbCrLf & _
              "echo - Для копирования ссылок используйте ЛЮБОЙ другой браузер" & vbCrLf & _
              "echo - Авторизуйтесь в YouTube в одном из поддерживаемых браузеров" & vbCrLf & _
              "echo - Проверьте правильность ввода ссылки" & vbCrLf & _
              "echo - Ссылка без ограничений на скачивание может дать ложный результат!" & vbCrLf & _
              "echo." & vbCrLf & _
              "echo После закрытия браузера повторите проверку авторизации." & vbCrLf & _
              "echo ===============================================" & vbCrLf & _
              "echo." & vbCrLf & _
              "echo Нажмите любую клавишу для выхода..." & vbCrLf & _
              "pause >nul"
    
    CreateTextFile binPath & "auth_check.bat", content
End Sub

Sub CreateCookiesBat()
    Dim content
    content = "@echo off" & vbCrLf & _
              "chcp 65001 >nul" & vbCrLf & _
              "title Тест cookies из edge" & vbCrLf & _
              "" & vbCrLf & _
              "echo Пк поддерживаемых браузеров yt-dlp с cookies из edge..." & vbCrLf & _
              "echo ----------------------------------------------" & vbCrLf & _
              "" & vbCrLf & _
              "yt-dlp --cookies-from-browser help" & vbCrLf & _
              "" & vbCrLf & _
              "echo." & vbCrLf & _
              "echo ----------------------------------------------" & vbCrLf & _
              "echo Завершено. Нажмите любую клавишу для выхода." & vbCrLf & _
              "pause >nul"
    
    CreateTextFile binPath & "cookies-from-browser.bat", content
End Sub

Sub CreateUpdateBat()
    Dim content
    content = "@echo off" & vbCrLf & _
              "chcp 65001 >nul" & vbCrLf & _
              "cd /d ""%~dp0app""" & vbCrLf & _
              vbCrLf & _
              "echo =========================================" & vbCrLf & _
              "echo        MULTILOADER UPDATE" & vbCrLf & _
              "echo =========================================" & vbCrLf & _
              "echo." & vbCrLf & _
              vbCrLf & _
              "REM === 1. YT-DLP ===" & vbCrLf & _
              "echo [1/3] Проверка/обновление yt-dlp..." & vbCrLf & _
              "if not exist ""bin\yt-dlp.exe"" (" & vbCrLf & _
              "    echo ❌ yt-dlp отсутствует - скачиваем..." & vbCrLf & _
              "    powershell -c ""iwr -outf bin\yt-dlp.exe 'https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe'""" & vbCrLf & _
              "    if exist ""bin\yt-dlp.exe"" (" & vbCrLf & _
              "        echo ✅ yt-dlp скачан" & vbCrLf & _
              "    ) else (" & vbCrLf & _
              "        echo ❌ Ошибка! Скачайте вручную по ссылке https://github.com/yt-dlp/yt-dlp/releases и поместите в app\bin\" & vbCrLf & _
              "        pause >nul" & vbCrLf & _
              "        exit /b" & vbCrLf & _
              "    )" & vbCrLf & _
              ") else (" & vbCrLf & _
              "    echo ✅ yt-dlp найден" & vbCrLf & _
              "    echo Очистка кеша..." & vbCrLf & _
              "    bin\yt-dlp.exe --rm-cache-dir >nul 2>&1" & vbCrLf & _
              "    echo Проверка обновлений..." & vbCrLf & _
              "    bin\yt-dlp.exe -U" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              "echo." & vbCrLf & _
              vbCrLf & _
              "REM === 2. FFMPEG ===" & vbCrLf & _
              "echo [2/3] Проверка/обновление пакета FFmpeg..." & vbCrLf & _
              vbCrLf & _
              "REM --- Определяем разрядность системы ---" & vbCrLf & _
              "for /f ""tokens=2 delims=="" %%I in ('wmic os get osarchitecture /value 2^>nul') do set ""ARCH=%%I""" & vbCrLf & _
              "if ""%ARCH%""=="""" (" & vbCrLf & _
              "    if defined PROCESSOR_ARCHITEW6432 (" & vbCrLf & _
              "        set ""ARCH=64-bit""" & vbCrLf & _
              "    ) else (" & vbCrLf & _
              "        set ""ARCH=32-bit""" & vbCrLf & _
              "    )" & vbCrLf & _
              ") else (" & vbCrLf & _
              "    set ""ARCH=%ARCH:~0,-1%""" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              "if ""%ARCH%""==""32-bit"" (" & vbCrLf & _
              "    set ""ARCH_TYPE=32""" & vbCrLf & _
              "    set ""ARCHIVE_URL=https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win32-gpl.zip""" & vbCrLf & _
              "    set ""ARCHIVE_NAME=ffmpeg-master-latest-win32-gpl.zip""" & vbCrLf & _
              ") else (" & vbCrLf & _
              "    set ""ARCH_TYPE=64""" & vbCrLf & _
              "    set ""ARCHIVE_URL=https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip""" & vbCrLf & _
              "    set ""ARCHIVE_NAME=ffmpeg-master-latest-win64-gpl.zip""" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              "set ""VERSION_FILE=bin\ffmpeg_version.txt""" & vbCrLf & _
              vbCrLf & _
              "REM --- Проверяем наличие файла версии ---" & vbCrLf & _
              "if not exist ""%VERSION_FILE%"" (" & vbCrLf & _
              "    goto :download_ffmpeg" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              "REM --- Читаем размер архива из файла (вторая строка) ---" & vbCrLf & _
              "< ""%VERSION_FILE%"" (" & vbCrLf & _
              "    set /p OLD_ARCH_TYPE=" & vbCrLf & _
              "    set /p OLD_SIZE=" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              "REM --- Проверяем размер архива на GitHub ---" & vbCrLf & _
              "echo [UPDATE] Проверяем обновление на GitHub..." & vbCrLf & _
              "for /f ""tokens=*"" %%I in ('powershell -c ""try {(Invoke-WebRequest '%ARCHIVE_URL%' -Method Head).Headers.'Content-Length'} catch {echo ERROR}"" 2^>nul') do set ""REMOTE_SIZE=%%I""" & vbCrLf & _
              vbCrLf & _
              "if ""%REMOTE_SIZE%""=="""" (" & vbCrLf & _
              "    echo ⚠️ Не удалось проверить актуальность версии, скачиваем заново..." & vbCrLf & _
              "    goto :download_ffmpeg" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              "REM --- Сравниваем размеры ---" & vbCrLf & _
              "if ""%OLD_SIZE%""==""%REMOTE_SIZE%"" (" & vbCrLf & _
              "    echo ✅ Пакет ffmpeg в обновлении не нуждается" & vbCrLf & _
              "    echo." & vbCrLf & _
              "    goto :ffmpeg_end" & vbCrLf & _
              ") else (" & vbCrLf & _
              "    echo 🔄 Требуется обновление" & vbCrLf & _
              "    goto :download_ffmpeg" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              ":download_ffmpeg" & vbCrLf & _
              "echo." & vbCrLf & _
              vbCrLf & _
              "REM --- Скачиваем архив ---" & vbCrLf & _
              "powershell -c ""Invoke-WebRequest -Uri '%ARCHIVE_URL%' -OutFile '%ARCHIVE_NAME%'""" & vbCrLf & _
              vbCrLf & _
              "if not exist ""%ARCHIVE_NAME%"" (" & vbCrLf & _
              "    echo ❌ Ошибка скачивания FFmpeg!" & vbCrLf & _
              "    echo 🔗 Скачайте вручную: %ARCHIVE_URL%" & vbCrLf & _
              "    echo 📁 И распакуйте в app\bin\ файлы: ffmpeg.exe, ffplay.exe, ffprobe.exe" & vbCrLf & _
              "    goto :ffmpeg_end" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              "REM --- Получаем размер скачанного файла ---" & vbCrLf & _
              "for /f %%I in ('powershell -c ""(gi '%ARCHIVE_NAME%').Length""') do set ""DOWNLOADED_SIZE=%%I""" & vbCrLf & _
              vbCrLf & _
              "REM --- Распаковываем архив ---" & vbCrLf & _
              "if exist ""temp_ffmpeg"" rmdir /s /q ""temp_ffmpeg""" & vbCrLf & _
              "mkdir ""temp_ffmpeg""" & vbCrLf & _
              vbCrLf & _
              "powershell -c ""Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('%ARCHIVE_NAME%', 'temp_ffmpeg')""" & vbCrLf & _
              vbCrLf & _
              "for /d %%I in (""temp_ffmpeg\*"") do (" & vbCrLf & _
              "    if exist ""%%I\bin\*.exe"" copy ""%%I\bin\*.exe"" ""bin\""" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              "rmdir /s /q ""temp_ffmpeg""" & vbCrLf & _
              vbCrLf & _
              "REM --- Создаем файл версии ---" & vbCrLf & _
              "> ""%VERSION_FILE%"" echo %ARCH_TYPE%" & vbCrLf & _
              ">> ""%VERSION_FILE%"" echo %DOWNLOADED_SIZE%" & vbCrLf & _
              ">> ""%VERSION_FILE%"" echo %date% %time%" & vbCrLf & _
              vbCrLf & _
              "echo ✅ FFmpeg успешно обновлен!" & vbCrLf & _
              vbCrLf & _
              "REM --- Удаляем архив ---" & vbCrLf & _
              "del ""%ARCHIVE_NAME%"" >nul 2>&1" & vbCrLf & _
              vbCrLf & _
              ":ffmpeg_end" & vbCrLf & _
              "echo === Проверка FFMPEG завершена ===" & vbCrLf & _
              "echo." & vbCrLf & _
              vbCrLf & _
              "cd /d ""%~dp0""" & vbCrLf & _
              "echo [3/3] Обновление списка поддерживаемых сайтов..." & vbCrLf & _
              vbCrLf & _
              "setlocal" & vbCrLf & _
              vbCrLf & _
              "set ""APP_DIR=app""" & vbCrLf & _
              "set ""FILE=%APP_DIR%\supportedsites.md""" & vbCrLf & _
              "set ""OLD_FILE=%APP_DIR%\old_supportedsites.md""" & vbCrLf & _
              "set ""TEMP_USERLIST=temp_userlist.txt""" & vbCrLf & _
              "set ""URL=https://raw.githubusercontent.com/yt-dlp/yt-dlp/master/supportedsites.md""" & vbCrLf & _
              vbCrLf & _
              "if not exist ""%APP_DIR%"" mkdir ""%APP_DIR%""" & vbCrLf & _
              vbCrLf & _
              ":: --- Сохраняем старый файл или скачиваем если файла нет ---" & vbCrLf & _
              "if exist ""%FILE%"" (" & vbCrLf & _
              "   copy /Y ""%FILE%"" ""%OLD_FILE%"" >nul" & vbCrLf & _
              ") else (curl -L -o ""%FILE%"" ""%URL%""" & vbCrLf & _
              "  echo 📋 Добавлен список поддерживаемых сайтов" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              ":: --- Обновляем файл с GitHub ---" & vbCrLf & _
              "curl -L -o ""%FILE%"" ""%URL%""" & vbCrLf & _
              vbCrLf & _
              "if not exist ""%FILE%"" (" & vbCrLf & _
              "    echo ❌ Ошибка: не удалось скачать supportedsites.md" & vbCrLf & _
              "	echo 🔗 Найти файл вручную: %URL%" & vbCrLf & _
              "	echo 📁 Поместите его в папку app\" & vbCrLf & _
              "	goto :end_section" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              ":: --- Извлекаем блок из старого файла, если он есть ---" & vbCrLf & _
              "if exist ""%OLD_FILE%"" (" & vbCrLf & _
              "    echo 🔍 Поиск пользовательского списка..." & vbCrLf & _
              "    powershell -NoProfile -Command ^" & vbCrLf & _
              "      ""$lines = Get-Content -Raw -Encoding UTF8 '%OLD_FILE%';"" ^" & vbCrLf & _
              "      ""$idx = $lines.IndexOf('===user list===');"" ^" & vbCrLf & _
              "      ""if ($idx -ge 0) {"" ^" & vbCrLf & _
              "      ""  $tail = $lines.Substring($idx);"" ^" & vbCrLf & _
              "      ""  $tail | Out-File -Encoding UTF8 '%TEMP_USERLIST%';"" ^" & vbCrLf & _
              "      ""  Write-Host '✔️ Пользовательский список найден и сохранён.';"" ^" & vbCrLf & _
              "      ""} else { Write-Host '💡 У вас еще нет своего списка.' }""" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              ":: --- Добавляем пользовательский блок в конец нового файла ---" & vbCrLf & _
              "if exist ""%TEMP_USERLIST%"" (" & vbCrLf & _
              "    echo.>>""%FILE%""" & vbCrLf & _
              "    type ""%TEMP_USERLIST%"" >>""%FILE%""" & vbCrLf & _
              "    del ""%TEMP_USERLIST%"" >nul 2>&1" & vbCrLf & _
              ")" & vbCrLf & _
              vbCrLf & _
              ":: --- Удаляем временный старый файл ---" & vbCrLf & _
              "if exist ""%OLD_FILE%"" (" & vbCrLf & _
              "    del ""%OLD_FILE%"" >nul 2>&1" & vbCrLf & _
              ")" & vbCrLf & _
              ":end_section" & vbCrLf & _
              "echo ✅ Обновление завершено." & vbCrLf & _
              "echo." & vbCrLf & _
              vbCrLf & _
              "endlocal" & vbCrLf & _
              "echo 🚀 Нажмите любую клавишу для запуска приложения..." & vbCrLf & _
              "pause >nul" & vbCrLf & _
              vbCrLf & _
              "REM === Запуск основного приложения ===" & vbCrLf & _
              "cd /d ""%~dp0""" & vbCrLf & _
              "start """" /D ""%~dp0app"" ""app\MultiLoader.hta""" & vbCrLf & _
              vbCrLf & _
              "exit /b"
    
    CreateTextFile basePath & "\update.bat", content
End Sub

Sub CreateTextFile(filePath, content)
    On Error Resume Next
    Dim stream, bytes
    Set stream = CreateObject("ADODB.Stream")
    
    ' Создаем файл в UTF-8 БЕЗ BOM
    stream.Type = 2 ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    ' Переключаем в binary режим и убираем BOM
    stream.Position = 0
    stream.Type = 1 ' binary
    stream.Position = 3 ' пропускаем BOM
    bytes = stream.Read
    stream.Position = 0
    stream.Write bytes
    stream.SetEOS
    stream.SaveToFile filePath, 2 ' overwrite
    stream.Close
End Sub