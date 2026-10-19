<#
.SYNOPSIS
    GUI-скрипт для переноса файлов баз данных SQL Server (метаданные + физическое перемещение).
.DESCRIPTION
    Windows Forms интерфейс: выбор баз, указание нового пути, генерация SQL,
    опциональное физическое перемещение файлов через Move-Item (OFFLINE -> move -> MODIFY FILE -> ONLINE).
    Системные БД (master, model, msdb, tempdb) исключены.
#>

#region Загрузка сборок
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null
#endregion

#region Функции

function Get-SqlServerConnection {
    param([string]$ServerName, [string]$UserName, [string]$Password, [switch]$UseWindowsAuth)
    $conn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection
    $conn.ServerInstance = $ServerName
    if ($UseWindowsAuth) { $conn.LoginSecure = $true }
    else { $conn.LoginSecure = $false; $conn.Login = $UserName; $conn.Password = $Password }
    return New-Object Microsoft.SqlServer.Management.Smo.Server($conn)
}

function Get-UserDatabases {
    param($Server)
    $systemDbs = @('master', 'model', 'msdb', 'tempdb')
    $Server.Databases | Where-Object { $systemDbs -notcontains $_.Name } | Sort-Object Name
}

function Get-DatabaseFiles {
    param($Database)
    $files = @()
    foreach ($fg in $Database.FileGroups) {
        foreach ($file in $fg.Files) {
            $files += [PSCustomObject]@{
                LogicalName = $file.Name; PhysicalName = $file.FileName
                Type = 'ROWS'; SizeMB = [math]::Round($file.Size / 1024, 2)
            }
        }
    }
    foreach ($logFile in $Database.LogFiles) {
        $files += [PSCustomObject]@{
            LogicalName = $logFile.Name; PhysicalName = $logFile.FileName
            Type = 'LOG'; SizeMB = [math]::Round($logFile.Size / 1024, 2)
        }
    }
    return $files
}

function Get-CurrentDirectories {
    param($Database)
    $dirs = @()
    foreach ($f in (Get-DatabaseFiles -Database $Database)) {
        $dir = Split-Path $f.PhysicalName -Parent
        if ($dir -and ($dirs -notcontains $dir)) { $dirs += $dir }
    }
    return $dirs
}

function Generate-MoveScript {
    param([string]$DatabaseName, [string]$NewPath, [array]$Files)
    $script = @()
    $script += "-- ============================================================"
    $script += "-- Перенос базы данных: $DatabaseName  (файлов: $($Files.Count))"
    $script += "-- ============================================================"
    $allSame = $true
    foreach ($file in $Files) {
        if ((Split-Path $file.PhysicalName -Parent) -ne $NewPath) { $allSame = $false; break }
    }
    if ($allSame) {
        $script += "-- ВНИМАНИЕ: все файлы базы уже в '$NewPath'. Перенос не требуется."
        $script += ""
        return ($script -join "`r`n")
    }
    $script += "USE master;"
    $script += "GO"
    $script += "ALTER DATABASE [$DatabaseName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;"
    $script += "GO"
    foreach ($file in $Files) {
        $newFileName = Join-Path $NewPath (Split-Path $file.PhysicalName -Leaf)
        $script += "-- $($file.Type): $($file.LogicalName) -> $newFileName"
        $script += "ALTER DATABASE [$DatabaseName] MODIFY FILE (NAME = N'$($file.LogicalName)', FILENAME = N'$newFileName');"
    }
    $script += "GO"
    $script += "ALTER DATABASE [$DatabaseName] SET MULTI_USER;"
    $script += "GO"
    $script += ""
    return ($script -join "`r`n")
}

#endregion

#region Форма
$form = New-Object System.Windows.Forms.Form
$form.Text = "Перенос баз данных SQL Server"
$form.Size = New-Object System.Drawing.Size(1200, 780)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(900, 650)

# --- Подключение ---
$gbConnection = New-Object System.Windows.Forms.GroupBox
$gbConnection.Text = "Подключение к SQL Server"
$gbConnection.Location = New-Object System.Drawing.Point(10, 10)
$gbConnection.Size = New-Object System.Drawing.Size(1160, 80)
$gbConnection.Anchor = "Top, Left, Right"

$lblServer = New-Object System.Windows.Forms.Label
$lblServer.Text = "Сервер:"; $lblServer.Location = New-Object System.Drawing.Point(10, 25)
$lblServer.Size = New-Object System.Drawing.Size(60, 20)
$gbConnection.Controls.Add($lblServer)

$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Text = "localhost"; $txtServer.Location = New-Object System.Drawing.Point(80, 22)
$txtServer.Size = New-Object System.Drawing.Size(200, 20)
$gbConnection.Controls.Add($txtServer)

$chkWindowsAuth = New-Object System.Windows.Forms.CheckBox
$chkWindowsAuth.Text = "Windows-аутентификация"; $chkWindowsAuth.Location = New-Object System.Drawing.Point(300, 22)
$chkWindowsAuth.Size = New-Object System.Drawing.Size(180, 20); $chkWindowsAuth.Checked = $true
$gbConnection.Controls.Add($chkWindowsAuth)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Пользователь:"; $lblUser.Location = New-Object System.Drawing.Point(490, 25)
$lblUser.Size = New-Object System.Drawing.Size(80, 20)
$gbConnection.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(570, 22); $txtUser.Size = New-Object System.Drawing.Size(120, 20)
$txtUser.Enabled = $false
$gbConnection.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Пароль:"; $lblPass.Location = New-Object System.Drawing.Point(700, 25)
$lblPass.Size = New-Object System.Drawing.Size(60, 20)
$gbConnection.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(760, 22); $txtPass.Size = New-Object System.Drawing.Size(120, 20)
$txtPass.Enabled = $false; $txtPass.UseSystemPasswordChar = $true
$gbConnection.Controls.Add($txtPass)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Подключиться"; $btnConnect.Location = New-Object System.Drawing.Point(1080, 20)
$btnConnect.Size = New-Object System.Drawing.Size(70, 25)
$gbConnection.Controls.Add($btnConnect)
$form.Controls.Add($gbConnection)

# --- Таблица БД ---
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(10, 100)
$dgv.Size = New-Object System.Drawing.Size(1160, 260)
$dgv.Anchor = "Top, Left, Right"
$dgv.AllowUserToAddRows = $false; $dgv.AllowUserToDeleteRows = $false
$dgv.SelectionMode = "FullRowSelect"; $dgv.MultiSelect = $false
$dgv.AutoSizeColumnsMode = "None"; $dgv.RowHeadersVisible = $false
$dgv.ShowCellToolTips = $true

$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.Name = "Check"; $colCheck.HeaderText = "Выбрать"; $colCheck.Width = 60
$dgv.Columns.Add($colCheck) | Out-Null

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = "Name"; $colName.HeaderText = "База данных"; $colName.ReadOnly = $true; $colName.Width = 180
$dgv.Columns.Add($colName) | Out-Null

$colFiles = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colFiles.Name = "Files"; $colFiles.HeaderText = "Файлов"; $colFiles.ReadOnly = $true; $colFiles.Width = 60
$dgv.Columns.Add($colFiles) | Out-Null

$colCurrentPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCurrentPath.Name = "CurrentPath"; $colCurrentPath.HeaderText = "Текущее расположение"
$colCurrentPath.ReadOnly = $true; $colCurrentPath.Width = 380
$dgv.Columns.Add($colCurrentPath) | Out-Null

$colPurpose = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colPurpose.Name = "Purpose"; $colPurpose.HeaderText = "Назначение"; $colPurpose.ReadOnly = $true; $colPurpose.Width = 220
$dgv.Columns.Add($colPurpose) | Out-Null

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.Name = "Status"; $colStatus.HeaderText = "Состояние"; $colStatus.ReadOnly = $true; $colStatus.Width = 100
$dgv.Columns.Add($colStatus) | Out-Null

$form.Controls.Add($dgv)

# --- Новый путь ---
$gbPath = New-Object System.Windows.Forms.GroupBox
$gbPath.Text = "Новое расположение файлов"
$gbPath.Location = New-Object System.Drawing.Point(10, 370)
$gbPath.Size = New-Object System.Drawing.Size(1160, 60)
$gbPath.Anchor = "Top, Left, Right"

$lblNewPath = New-Object System.Windows.Forms.Label
$lblNewPath.Text = "Путь:"; $lblNewPath.Location = New-Object System.Drawing.Point(10, 25)
$lblNewPath.Size = New-Object System.Drawing.Size(40, 20)
$gbPath.Controls.Add($lblNewPath)

$txtNewPath = New-Object System.Windows.Forms.TextBox
$txtNewPath.Text = "D:\SQLData"; $txtNewPath.Location = New-Object System.Drawing.Point(60, 22)
$txtNewPath.Size = New-Object System.Drawing.Size(700, 20)
$gbPath.Controls.Add($txtNewPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Обзор..."; $btnBrowse.Location = New-Object System.Drawing.Point(770, 20)
$btnBrowse.Size = New-Object System.Drawing.Size(80, 25)
$gbPath.Controls.Add($btnBrowse)

$chkPhysicalMove = New-Object System.Windows.Forms.CheckBox
$chkPhysicalMove.Text = "Физически перемещать файлы (Move-Item)"
$chkPhysicalMove.Location = New-Object System.Drawing.Point(880, 22)
$chkPhysicalMove.Size = New-Object System.Drawing.Size(270, 20)
$chkPhysicalMove.Checked = $false
$gbPath.Controls.Add($chkPhysicalMove)

$form.Controls.Add($gbPath)

# --- Кнопки ---
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Сгенерировать скрипты"
$btnGenerate.Location = New-Object System.Drawing.Point(10, 440)
$btnGenerate.Size = New-Object System.Drawing.Size(170, 30)
$btnGenerate.Anchor = "Top, Left"
$form.Controls.Add($btnGenerate)

$btnExecute = New-Object System.Windows.Forms.Button
$btnExecute.Text = "Выполнить"
$btnExecute.Location = New-Object System.Drawing.Point(190, 440)
$btnExecute.Size = New-Object System.Drawing.Size(150, 30)
$btnExecute.Anchor = "Top, Left"
$btnExecute.Enabled = $false
$form.Controls.Add($btnExecute)

# --- Лог / скрипт ---
$txtScripts = New-Object System.Windows.Forms.TextBox
$txtScripts.Location = New-Object System.Drawing.Point(10, 480)
$txtScripts.Size = New-Object System.Drawing.Size(1160, 250)
$txtScripts.Anchor = "Top, Bottom, Left, Right"
$txtScripts.Multiline = $true; $txtScripts.ScrollBars = "Both"
$txtScripts.ReadOnly = $true; $txtScripts.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtScripts)

#endregion

#region Обработчики

$chkWindowsAuth.Add_CheckedChanged({
    $txtUser.Enabled = -not $chkWindowsAuth.Checked
    $txtPass.Enabled = -not $chkWindowsAuth.Checked
})

$btnConnect.Add_Click({
    try {
        $server = Get-SqlServerConnection -ServerName $txtServer.Text `
            -UseWindowsAuth:$chkWindowsAuth.Checked `
            -UserName $txtUser.Text -Password $txtPass.Text
        $global:SqlServer = $server
        $dbs = Get-UserDatabases -Server $server
        $dgv.Rows.Clear()
        foreach ($db in $dbs) {
            $purpose = ""
            $descProp = $db.ExtendedProperties | Where-Object { $_.Name -eq 'Description' }
            if ($descProp) { $purpose = $descProp.Value }

            $files = Get-DatabaseFiles -Database $db
            $dirs = Get-CurrentDirectories -Database $db
            $tooltipLines = @("База данных: $($db.Name)", "Файлов: $($files.Count)", "")
            foreach ($f in $files) {
                $tooltipLines += "[$($f.Type)] $($f.LogicalName) ($($f.SizeMB) MB)"
                $tooltipLines += "    $($f.PhysicalName)"
            }
            $tooltip = $tooltipLines -join "`r`n"

            $ri = $dgv.Rows.Add($false, $db.Name, $files.Count, ($dirs -join "; "), $purpose, $db.Status)
            $dgv.Rows[$ri].Cells["CurrentPath"].ToolTipText = $tooltip
            $dgv.Rows[$ri].Cells["Name"].ToolTipText = $tooltip
        }
        [System.Windows.Forms.MessageBox]::Show("Подключено. Баз: $($dbs.Count)", "Успех", "OK", "Information")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка подключения:`n$($_.Exception.Message)", "Ошибка", "OK", "Error")
    }
})

$btnBrowse.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = "Выберите папку для новых файлов баз данных"
    if ($fb.ShowDialog() -eq "OK") { $txtNewPath.Text = $fb.SelectedPath }
})

$btnGenerate.Add_Click({
    if (-not $global:SqlServer) {
        [System.Windows.Forms.MessageBox]::Show("Сначала подключитесь.", "Внимание", "OK", "Warning"); return
    }
    $selectedDbs = @()
    for ($i = 0; $i -lt $dgv.Rows.Count; $i++) {
        if ($dgv.Rows[$i].Cells["Check"].Value -eq $true) {
            $selectedDbs += $dgv.Rows[$i].Cells["Name"].Value
        }
    }
    if ($selectedDbs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Не выбрано ни одной базы.", "Внимание", "OK", "Warning"); return
    }
    $newPath = $txtNewPath.Text.Trim()
    if ([string]::IsNullOrEmpty($newPath)) {
        [System.Windows.Forms.MessageBox]::Show("Укажите новый путь.", "Внимание", "OK", "Warning"); return
    }

    # Сохраняем «план переноса» — чтобы при выполнении не парсить текст, а знать точно
    $global:MovePlan = @()
    $allScripts = @()
    foreach ($dbName in $selectedDbs) {
        $db = $global:SqlServer.Databases[$dbName]
        if (-not $db) { continue }
        $files = Get-DatabaseFiles -Database $db
        $filePlan = @()
        foreach ($f in $files) {
            $filePlan += [PSCustomObject]@{
                LogicalName = $f.LogicalName
                OldPath     = $f.PhysicalName
                NewPath     = Join-Path $newPath (Split-Path $f.PhysicalName -Leaf)
                Type        = $f.Type
            }
        }
        $global:MovePlan += [PSCustomObject]@{
            Database = $dbName
            NewPath  = $newPath
            Files    = $filePlan
        }
        $allScripts += Generate-MoveScript -DatabaseName $dbName -NewPath $newPath -Files $files
    }
    $txtScripts.Text = $allScripts -join "`r`n`r`n"
    $btnExecute.Enabled = $true
    [System.Windows.Forms.MessageBox]::Show("Скрипты сгенерированы. Проверьте их в поле ниже.", "Готово", "OK", "Information")
})

$btnExecute.Add_Click({
    if (-not $global:MovePlan -or $global:MovePlan.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Сначала сгенерируйте скрипты.", "Внимание", "OK", "Warning"); return
    }

    $doPhysical = $chkPhysicalMove.Checked
    $msg = if ($doPhysical) {
        "Будет выполнено:`n" +
        "1. OFFLINE баз (с откатом транзакций)`n" +
        "2. ФИЗИЧЕСКОЕ перемещение файлов через Move-Item`n" +
        "3. MODIFY FILE (новые пути)`n" +
        "4. ONLINE`n`nПродолжить?"
    } else {
        "Будет выполнена только замена путей в метаданных (ALTER DATABASE MODIFY FILE).`n" +
        "Файлы физически НЕ перемещаются.`n`nПродолжить?"
    }
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Подтверждение", "YesNo", "Question") -ne "Yes") { return }

    $log = New-Object System.Text.StringBuilder
    [void]$log.AppendLine("=== Запуск $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===  PhysicalMove=$doPhysical`r`n")

    try {
        foreach ($plan in $global:MovePlan) {
            $dbName = $plan.Database
            [void]$log.AppendLine(">>> База: $dbName")

            if ($doPhysical) {
                # 1) OFFLINE — освобождает файловые дескрипторы
                [void]$log.AppendLine("    SET OFFLINE ...")
                $global:SqlServer.ConnectionContext.ExecuteNonQuery(
                    "USE master; ALTER DATABASE [$dbName] SET OFFLINE WITH ROLLBACK IMMEDIATE;")

                try {
                    # 2) Физическое перемещение
                    foreach ($f in $plan.Files) {
                        if ($f.OldPath -eq $f.NewPath) {
                            [void]$log.AppendLine("    [skip] $($f.LogicalName): уже на месте")
                            continue
                        }
                        if (-not (Test-Path -LiteralPath $f.OldPath)) {
                            throw "Файл не найден: $($f.OldPath)"
                        }
                        $destDir = Split-Path $f.NewPath -Parent
                        if (-not (Test-Path -LiteralPath $destDir)) {
                            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                            [void]$log.AppendLine("    [mkdir] $destDir")
                        }
                        if (Test-Path -LiteralPath $f.NewPath) {
                            throw "Целевой файл уже существует: $($f.NewPath)"
                        }
                        Move-Item -LiteralPath $f.OldPath -Destination $f.NewPath -Force -ErrorAction Stop
                        [void]$log.AppendLine("    [move] $($f.OldPath) -> $($f.NewPath)")
                    }

                    # 3) Обновление метаданных
                    [void]$log.AppendLine("    MODIFY FILE ...")
                    $sqlModify = "USE master;`r`n"
                    foreach ($f in $plan.Files) {
                        $sqlModify += "ALTER DATABASE [$dbName] MODIFY FILE (NAME = N'$($f.LogicalName)', FILENAME = N'$($f.NewPath)');`r`n"
                    }
                    $global:SqlServer.ConnectionContext.ExecuteNonQuery($sqlModify)
                }
                finally {
                    # 4) ONLINE в любом случае
                    [void]$log.AppendLine("    SET ONLINE ...")
                    try {
                        $global:SqlServer.ConnectionContext.ExecuteNonQuery(
                            "USE master; ALTER DATABASE [$dbName] SET ONLINE;")
                    } catch {
                        [void]$log.AppendLine("    [WARN] ONLINE: $($_.Exception.Message)")
                    }
                }
            }
            else {
                # Только метаданные (как раньше)
                $sqlOnly = "USE master;`r`nALTER DATABASE [$dbName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;`r`n"
                foreach ($f in $plan.Files) {
                    $sqlOnly += "ALTER DATABASE [$dbName] MODIFY FILE (NAME = N'$($f.LogicalName)', FILENAME = N'$($f.NewPath)');`r`n"
                }
                $sqlOnly += "ALTER DATABASE [$dbName] SET MULTI_USER;`r`n"
                $global:SqlServer.ConnectionContext.ExecuteNonQuery($sqlOnly)
                [void]$log.AppendLine("    MODIFY FILE выполнен (без физического переноса)")
            }
            [void]$log.AppendLine("")
        }
        [void]$log.AppendLine("=== Успешно завершено ===")
        $txtScripts.Text = $log.ToString()
        [System.Windows.Forms.MessageBox]::Show("Готово. Смотрите лог в окне скриптов.", "Успех", "OK", "Information")
        $btnExecute.Enabled = $false
    } catch {
        [void]$log.AppendLine("!!! ОШИБКА: $($_.Exception.Message)")
        $txtScripts.Text = $log.ToString()
        [System.Windows.Forms.MessageBox]::Show("Ошибка:`n$($_.Exception.Message)`n`nСмотрите лог.", "Ошибка", "OK", "Error")
    }
})

#endregion

[void]$form.ShowDialog()
