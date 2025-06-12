Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic  # For Recycle Bin deletion

$form = New-Object Windows.Forms.Form
$form.Text = "Drag to Unzip (Recursive)"
$form.Size = New-Object Drawing.Size(550, 400)
$form.StartPosition = "CenterScreen"
$form.AllowDrop = $true

# Label
$label = New-Object Windows.Forms.Label
$label.Text = "Drag zip-like files here to extract them recursively"
$label.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$label.TextAlign = "MiddleCenter"
$label.Dock = "Top"
$label.Height = 40
$form.Controls.Add($label)

# Password label
$labelPassword = New-Object Windows.Forms.Label
$labelPassword.Text = "Password (leave blank if none):"
$labelPassword.Location = New-Object Drawing.Point(10, 50)
$labelPassword.Size = New-Object Drawing.Size(200, 20)
$form.Controls.Add($labelPassword)

# Password input
$textBoxPassword = New-Object Windows.Forms.TextBox
$textBoxPassword.Location = New-Object Drawing.Point(220, 48)
$textBoxPassword.Size = New-Object Drawing.Size(200, 20)
$form.Controls.Add($textBoxPassword)

# Delete option
$checkDelete = New-Object Windows.Forms.CheckBox
$checkDelete.Text = "Delete original archive after extraction (send to Recycle Bin)"
$checkDelete.Location = New-Object Drawing.Point(10, 75)
$checkDelete.AutoSize = $true
$form.Controls.Add($checkDelete)

# Log textbox
$textBoxLog = New-Object Windows.Forms.TextBox
$textBoxLog.Multiline = $true
$textBoxLog.ScrollBars = "Vertical"
$textBoxLog.Location = New-Object Drawing.Point(10, 105)
$textBoxLog.Size = New-Object Drawing.Size(510, 240)
$textBoxLog.ReadOnly = $true
$form.Controls.Add($textBoxLog)

# 7-Zip path
$sevenZipPath = 'C:\Program Files\7-Zip\7z.exe'
if (-not (Test-Path $sevenZipPath)) {
    [System.Windows.Forms.MessageBox]::Show("7z.exe not found. Please install 7-Zip and update the path.")
    exit
}

# Recursive unzip
function UnzipRecursively($file, $password = "", $deleteAfter = $false) {
    if ($file -eq $MyInvocation.MyCommand.Definition) { return }

    $parentDir = (Get-Item $file).DirectoryName
    $fileName = [IO.Path]::GetFileNameWithoutExtension($file)
    $outputDir = Join-Path $parentDir "$fileName`_unzipped"

    $cmd = "`"$sevenZipPath`" x -y -p$password -o`"$outputDir`" `"$file`""
    cmd /c $cmd | Out-Null

    if ($LASTEXITCODE -eq 0) {
        $textBoxLog.AppendText("[OK] Unzipped: $file -> $outputDir`r`n")

        Get-ChildItem -Path $outputDir -Recurse -File | ForEach-Object {
            UnzipRecursively $_.FullName $password $deleteAfter
        }

        if ($deleteAfter) {
            try {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                    $file,
                    'OnlyErrorDialogs',
                    'SendToRecycleBin'
                )
                $textBoxLog.AppendText("[DELETED] $file`r`n")
            } catch {
                $textBoxLog.AppendText("[WARNING] Could not delete: $file`r`n")
            }
        }
    } else {
        $textBoxLog.AppendText("[FAIL] Could not unzip: $file`r`n")
    }
}

# Drag event handlers
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = "Copy"
    }
})

$form.Add_DragDrop({
    $password = $textBoxPassword.Text
    $deleteChoice = $checkDelete.Checked
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    foreach ($file in $files) {
        UnzipRecursively $file $password $deleteChoice
    }
})

[void]$form.ShowDialog()
