# =====================================================================
#  PowerShell-applikation med Grafiskt Gränssnitt (GUI) v72
#  - Skapar ett anpassat fönster med en fillista och en förhandsgranskning.
#  - Inkluderar: Enter-klick, dubbelklick, dra & släpp, PDF/MSG/Bild-visning.
#  - Inbyggd textredigering för info.txt.
#  - Val av skrivare direkt i menyraden och sparar inställningar.
#  - Kan ta emot ett sexsiffrigt nummer som argument vid uppstart.
#  - Knappar för "Målmapp" och "Nytt fönster".
#  - Robust implementering av FileSystemWatcher för automatisk uppdatering.
#  - info.txt är nu gulmarkerad direkt i fillistan.
#  - NYTT: "CM Kopia" styrs nu helt av checkboxen, ingen automatisk textsökning.
# =====================================================================

param(
    [string[]]$args
)

# --- Självuppdaterande logik ---
function Invoke-SelfUpdate {
    $localDir = "C:\IT"
    $localAppName = "ITMapp.exe"
    $localPath = Join-Path -Path $localDir -ChildPath $localAppName
    $currentExecutable = (Get-Process -Id $PID).Path

    if ($currentExecutable -and -not $currentExecutable.StartsWith($localDir, [System.StringComparison]::InvariantCultureIgnoreCase)) {
        if (-not (Test-Path $localDir)) {
            try { New-Item -Path $localDir -ItemType Directory -Force | Out-Null } catch { return }
        }

        $updateNeeded = $false
        if (-not (Test-Path $localPath)) {
            $updateNeeded = $true
        } else {
            try {
                $localHash = (Get-FileHash -Path $localPath -ErrorAction Stop).Hash
                $networkHash = (Get-FileHash -Path $currentExecutable -ErrorAction Stop).Hash
                if ($localHash -ne $networkHash) {
                    $updateNeeded = $true
                }
            } catch {
                $updateNeeded = $true
            }
        }

        if ($updateNeeded) {
            try {
                Copy-Item -Path $currentExecutable -Destination $localPath -Force
            } catch {
                return
            }
        }
        
        $startParams = @{ FilePath = $localPath }
        if ($args.Count -gt 0) {
            $startParams.ArgumentList = $args
        }
        Start-Process @startParams
        exit
    }
}
Invoke-SelfUpdate


# Ladda nödvändiga .NET-assemblys för att bygga ett GUI.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# --- FÄRGER OCH TYPSNITT ---
$theme = @{
    BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f5f5f5")
    TopPanelColor = [System.Drawing.ColorTranslator]::FromHtml("#e0e0e0")
    TextColor = [System.Drawing.ColorTranslator]::FromHtml("#212121")
    AccentColor = [System.Drawing.ColorTranslator]::FromHtml("#0078d7")
    InfoColor = [System.Drawing.ColorTranslator]::FromHtml("#fff4b8")
    ButtonColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff") 
    ButtonHover = [System.Drawing.ColorTranslator]::FromHtml("#e8e8e8") 
    BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#cccccc") 
    Font = New-Object System.Drawing.Font("Segoe UI", 9)
    MonoFont = New-Object System.Drawing.Font("Consolas", 9)
}

# --- Skapa Huvudfönstret ---
$originalFormTitle = "Öppna projektmapp"
$form = New-Object System.Windows.Forms.Form
$form.Text = $originalFormTitle
$form.Size = New-Object System.Drawing.Size(940, 600)
$form.StartPosition = "CenterScreen" 
$form.MinimumSize = New-Object System.Drawing.Size(840, 400)
$form.BackColor = $theme.BackColor

$script:startEditingAfterLoad = $false
$script:fileWatcher = New-Object System.IO.FileSystemWatcher
$script:tempImageFolder = $null 
$script:isDirty = $false
$script:originalInfoText = ""

# Sökväg för inställningsfil
$settingsPath = Join-Path -Path $env:APPDATA -ChildPath "ÖppnaProjektmapp"
$settingsFile = Join-Path -Path $settingsPath -ChildPath "settings.json"

# --- Skapa Kontroller (Knappar, Textfält etc.) ---
function New-ModernButton {
    param($Text, $Location, $Size)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = $Location
    $button.Size = $Size
    $button.Font = $theme.Font
    $button.FlatStyle = 'Flat'
    $button.FlatAppearance.BorderSize = 1
    $button.FlatAppearance.BorderColor = $theme.BorderColor
    $button.BackColor = $theme.ButtonColor
    $button.FlatAppearance.MouseOverBackColor = $theme.ButtonHover
    return $button
}

$topPanel = New-Object System.Windows.Forms.Panel; $topPanel.Dock = "Top"; $topPanel.Height = 80; $topPanel.BackColor = $theme.TopPanelColor
$label = New-Object System.Windows.Forms.Label; $label.Text = "Ange nummer:"; $label.Location = New-Object System.Drawing.Point(10, 16); $label.AutoSize = $true; $label.Font = $theme.Font
$textBox = New-Object System.Windows.Forms.TextBox; $textBox.Location = New-Object System.Drawing.Point(100, 14); $textBox.Size = New-Object System.Drawing.Size(100, 23); $textBox.Font = $theme.Font
$button = New-ModernButton -Text "Hämta" -Location (New-Object System.Drawing.Point(210, 12)) -Size (New-Object System.Drawing.Size(75, 28))
$buttonInfo = New-ModernButton -Text "Info" -Location (New-Object System.Drawing.Point(290, 12)) -Size (New-Object System.Drawing.Size(60, 28)); $buttonInfo.Enabled = $false
$buttonOpenFolder = New-ModernButton -Text "Målmapp" -Location (New-Object System.Drawing.Point(355, 12)) -Size (New-Object System.Drawing.Size(80, 28)); $buttonOpenFolder.Enabled = $false
$buttonNewWindow = New-ModernButton -Text "Nytt fönster" -Location (New-Object System.Drawing.Point(440, 12)) -Size (New-Object System.Drawing.Size(100, 28))
$labelPrinter = New-Object System.Windows.Forms.Label; $labelPrinter.Text = "Skrivare:"; $labelPrinter.Location = New-Object System.Drawing.Point(10, 50); $labelPrinter.AutoSize = $true; $labelPrinter.Font = $theme.Font
$printerComboBox = New-Object System.Windows.Forms.ComboBox; $printerComboBox.Location = New-Object System.Drawing.Point(65, 46); $printerComboBox.Size = New-Object System.Drawing.Size(180, 25); $printerComboBox.DropDownStyle = 'DropDownList'; $printerComboBox.Font = $theme.Font; $printerComboBox.FlatStyle = 'Flat'
$buttonPrint = New-ModernButton -Text "Skriv ut markerad" -Location (New-Object System.Drawing.Point(255, 44)) -Size (New-Object System.Drawing.Size(125, 28)); $buttonPrint.Enabled = $false
$buttonPrintAllPdf = New-ModernButton -Text "Skriv ut alla PDF" -Location (New-Object System.Drawing.Point(385, 44)) -Size (New-Object System.Drawing.Size(125, 28)); $buttonPrintAllPdf.Enabled = $false
$checkboxCmKopia = New-Object System.Windows.Forms.CheckBox; $checkboxCmKopia.Text = "CM Kopia"; $checkboxCmKopia.Location = New-Object System.Drawing.Point(520, 48); $checkboxCmKopia.Font = $theme.Font; $checkboxCmKopia.Visible = $false; $checkboxCmKopia.AutoSize = $true
$statusBar = New-Object System.Windows.Forms.StatusBar; $statusBar.Text = "Väntar på inmatning..."; $statusBar.Font = $theme.Font
$splitContainer = New-Object System.Windows.Forms.SplitContainer; $splitContainer.Dock = "Fill"; $splitContainer.FixedPanel = "Panel1"; $splitContainer.SplitterDistance = 250
$fileListBox = New-Object System.Windows.Forms.ListBox; $fileListBox.Dock = "Fill"; $fileListBox.Font = $theme.MonoFont; $fileListBox.AllowDrop = $true; $fileListBox.SelectionMode = 'MultiExtended'; $fileListBox.BorderStyle = 'None'; $fileListBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$fileListBox.DrawMode = 'OwnerDrawFixed'
$previewPanel = New-Object System.Windows.Forms.Panel; $previewPanel.Dock = "Fill"
$script:webBrowser = New-Object System.Windows.Forms.WebBrowser; $script:webBrowser.Dock = "Fill"; $script:webBrowser.ScriptErrorsSuppressed = $true 
$txtEditorPanel = New-Object System.Windows.Forms.Panel; $txtEditorPanel.Dock = "Top"; $txtEditorPanel.Height = 30; $txtEditorPanel.Visible = $false; $txtEditorPanel.BackColor = $theme.TopPanelColor
$btnSaveTxt = New-ModernButton -Text "Spara" -Location (New-Object System.Drawing.Point(5, 3)) -Size (New-Object System.Drawing.Size(75, 23))
$btnCancelTxt = New-ModernButton -Text "Avbryt" -Location (New-Object System.Drawing.Point(85, 3)) -Size (New-Object System.Drawing.Size(75, 23))
$txtEditorPanel.Controls.AddRange(@($btnSaveTxt, $btnCancelTxt))

# --- Lägg till Kontroller i Fönstret ---
$topPanel.Controls.AddRange(@($label, $textBox, $button, $buttonInfo, $buttonOpenFolder, $buttonNewWindow, $labelPrinter, $printerComboBox, $buttonPrint, $buttonPrintAllPdf, $checkboxCmKopia))
$previewPanel.Controls.Add($script:webBrowser)
$previewPanel.Controls.Add($txtEditorPanel)
$splitContainer.Panel1.Controls.Add($fileListBox)
$splitContainer.Panel2.Controls.Add($previewPanel)
$form.Controls.AddRange(@($splitContainer, $topPanel, $statusBar))

# --- Funktionsblock ---
function Load-Settings {
    if (Test-Path $settingsFile) {
        try {
            $settings = Get-Content -Path $settingsFile | ConvertFrom-Json
            $form.StartPosition = 'Manual'
            if ($settings.Location) { $form.Location = New-Object System.Drawing.Point($settings.Location.X, $settings.Location.Y) }
            if ($settings.Size) { $form.Size = New-Object System.Drawing.Size($settings.Size.Width, $settings.Size.Height) }
            if ($settings.SplitterDistance) { $splitContainer.SplitterDistance = $settings.SplitterDistance }
            if ($settings.LastPrinter -and $printerComboBox.Items.Contains($settings.LastPrinter)) {
                $printerComboBox.SelectedItem = $settings.LastPrinter
            }
        } catch {}
    }
}

function Save-Settings {
    if (-not (Test-Path $settingsPath)) {
        New-Item -Path $settingsPath -ItemType Directory -Force | Out-Null
    }
    $settings = [PSCustomObject]@{
        Location         = @{ X = $form.Location.X; Y = $form.Location.Y }
        Size             = @{ Width = $form.Size.Width; Height = $form.Size.Height }
        SplitterDistance = $splitContainer.SplitterDistance
        LastPrinter      = $printerComboBox.SelectedItem
    }
    $settings | ConvertTo-Json | Set-Content -Path $settingsFile
}

function Update-FileList {
    $script:fileWatcher.EnableRaisingEvents = $false
    $nummer = $textBox.Text
    $selectedItemsBeforeUpdate = @($fileListBox.SelectedItems)

    $fileListBox.Items.Clear()
    $statusBar.Text = "Bearbetar..."
    
    $form.BackColor = $theme.BackColor
    $buttonInfo.Enabled = $false; $buttonPrint.Enabled = $false; $buttonPrintAllPdf.Enabled = $false; $buttonOpenFolder.Enabled = $false
    $form.Text = $originalFormTitle
    $checkboxCmKopia.Visible = $false

    if ($nummer -match "^\d{6}$") {
        $firstTwo = $nummer.Substring(0, 2)
        $folderPath = "F:\it\data\dokument\$firstTwo\$nummer"

        if (Test-Path -Path $folderPath -PathType Container) {
            $statusBar.Text = "Mapp hittad: $folderPath"
            $form.Text = "$originalFormTitle - $nummer"
            $buttonInfo.Enabled = $true
            $buttonOpenFolder.Enabled = $true
            
            $files = Get-ChildItem -Path $folderPath -Recurse -File
            $hasCmCandidate = $false
            if ($files) {
                $buttonPrint.Enabled = $true
                $buttonPrintAllPdf.Enabled = $true
                $sortedFiles = $files | Sort-Object @{Expression={
                    $isRootInfoTxt = $_.Name -eq 'info.txt' -and $_.DirectoryName -eq $folderPath
                    if ($isRootInfoTxt) { 0 }
                    elseif ($_.Extension -eq '.pdf') { 1 }
                    elseif ($_.Extension -eq '.msg') { 2 }
                    else { 3 }
                }}, FullName

                foreach ($file in $sortedFiles) {
                    $relativePath = $file.FullName.Substring($folderPath.Length + 1)
                    $fileListBox.Items.Add($relativePath) | Out-Null
                    
                    $fileNameOnly = [System.IO.Path]::GetFileName($relativePath)
                    $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileNameOnly)

                    if ($relativePath.EndsWith(".pdf", [System.StringComparison]::InvariantCultureIgnoreCase) -and ($fileNameWithoutExt -match "^\d{8}$") -and $fileNameWithoutExt.EndsWith("15")) {
                        $hasCmCandidate = $true
                    }
                }

                foreach ($item in $selectedItemsBeforeUpdate) {
                    if ($fileListBox.Items.Contains($item)) {
                        $fileListBox.SelectedItems.Add($item)
                    }
                }
            } else {
                 $statusBar.Text = "Mappen är tom."
            }
            
            $infoFilePath = Join-Path -Path $folderPath -ChildPath "info.txt"
            if (Test-Path $infoFilePath) { 
                $form.BackColor = $theme.InfoColor
                if ($fileListBox.SelectedItems.Count -eq 0) {
                    $fileListBox.SelectedItem = "info.txt"
                }
            }
            
            if ($hasCmCandidate) {
                $checkboxCmKopia.Visible = $true
                $checkboxCmKopia.Checked = $true
            }

            $script:fileWatcher.Path = $folderPath
            $script:fileWatcher.EnableRaisingEvents = $true

        } else {
            $statusBar.Text = "Fel: Mappen hittades inte."
        }
    } else {
        $statusBar.Text = "Fel: Ange exakt sex siffror."
    }
}

function Toggle-TxtEditor($enable) {
    $txtEditorPanel.Visible = $enable
    $script:isDirty = $false 
    if ($enable -and $script:webBrowser.Document) {
        $script:originalInfoText = $script:webBrowser.Document.Body.InnerText
    }
    if ($script:webBrowser.Document) { $script:webBrowser.Document.Body.SetAttribute("contentEditable", $enable.ToString()) }
}

function Prompt-SaveChanges {
    if (-not $txtEditorPanel.Visible) { return $true }
    
    $currentText = $script:webBrowser.Document.Body.InnerText
    $script:isDirty = ($currentText -ne $script:originalInfoText)

    if (-not $script:isDirty) {
        Toggle-TxtEditor -enable $false
        return $true
    }

    $result = [System.Windows.Forms.MessageBox]::Show("Vill du spara ändringarna i info.txt?", "Osparade ändringar", 'YesNoCancel', 'Warning')
    
    switch ($result) {
        'Yes' { $btnSaveTxt.PerformClick(); return $true }
        'No' { Toggle-TxtEditor -enable $false; return $true }
        'Cancel' { return $false }
    }
}


function Cleanup-TempImages {
    if ($script:tempImageFolder -and (Test-Path $script:tempImageFolder)) {
        Remove-Item -Path $script:tempImageFolder -Recurse -Force
        $script:tempImageFolder = $null
    }
}

function Print-Files {
    param(
        [System.Collections.IEnumerable]$FilesToPrint,
        [string]$Printer,
        [string]$BaseFolderPath
    )
    
    $originalDefaultPrinter = (Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Default='true'").Name
    
    try {
        (Get-WmiObject -Class Win32_Printer -Filter "Name='$Printer'").SetDefaultPrinter() | Out-Null
        
        foreach ($selectedFile in $FilesToPrint) {
            $fullPath = Join-Path -Path $BaseFolderPath -ChildPath $selectedFile
            $fileNameOnly = [System.IO.Path]::GetFileName($selectedFile)
            $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileNameOnly)

            try {
                $statusBar.Text = "Skriver ut '$selectedFile'..."
                Start-Process -FilePath $fullPath -Verb Print -WindowStyle Hidden -PassThru | Wait-Process -Timeout 30
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Kunde inte skriva ut filen '$selectedFile'.`n`nFel: $($_.Exception.Message)", "Utskriftsfel", "OK", "Error")
            }

            if (($selectedFile.EndsWith(".pdf", [System.StringComparison]::InvariantCultureIgnoreCase)) -and ($fileNameWithoutExt -match "^\d{8}$") -and $fileNameWithoutExt.EndsWith("15") -and $checkboxCmKopia.Checked) {
                $statusBar.Text = "Skapar stämplad kopia för '$selectedFile'..."
                $word = $null; $doc = $null
                try {
                    $copyPath = Join-Path -Path $BaseFolderPath -ChildPath "CM_Kopia.PDF"
                    
                    if (-not (Test-Path $copyPath)) {
                        $word = New-Object -ComObject Word.Application; $word.Visible = $false
                        $doc = $word.Documents.Open($fullPath)
                        $header = $doc.Sections.Item(1).Headers.Item([Microsoft.Office.Interop.Word.WdHeaderFooterIndex]::wdHeaderFooterPrimary)
                        $shape = $header.Shapes.AddTextEffect(0, "COILMATE KOPIA", "Arial", 72, $false, $false, 0, 0)
                        $shape.Select(); $shape.Fill.ForeColor.RGB = 255; $shape.Fill.Transparency = 0.5; $shape.Line.Visible = $false
                        $shape.RelativeHorizontalPosition = 1; $shape.RelativeVerticalPosition = 0; $shape.Left = -999995; $shape.Top = -999995; $shape.Rotation = -45
                        $doc.SaveAs2([ref]$copyPath, [ref]17); $doc.Close([ref]$false); $word.Quit()
                    }
                    
                    $statusBar.Text = "Skriver ut stämplad kopia..."
                    Start-Process -FilePath $copyPath -Verb Print -WindowStyle Hidden -PassThru | Wait-Process -Timeout 30
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Ett fel uppstod vid hantering av COILMATE-PDF.`n`nFel: $($_.Exception.Message)", "PDF-fel", "OK", "Error")
                } finally {
                    if ($doc) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null }
                    if ($word) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null }
                }
            }
        }
        $statusBar.Text = "Utskriftsjobb klara."

    } finally {
        if ($originalDefaultPrinter) {
            (Get-WmiObject -Class Win32_Printer -Filter "Name='$originalDefaultPrinter'").SetDefaultPrinter() | Out-Null
        }
    }
}


# --- Definiera Händelser (Vad som händer när man klickar) ---

$form.Add_Shown({ 
    $textBox.Focus()
    # Fyll skrivarlistan
    $printers = [System.Drawing.Printing.PrinterSettings]::InstalledPrinters
    foreach ($printer in $printers) {
        $printerComboBox.Items.Add($printer) | Out-Null
    }
    $defaultPrinter = (New-Object System.Drawing.Printing.PrinterSettings).PrinterName
    if ($printerComboBox.Items.Contains($defaultPrinter)) {
        $printerComboBox.SelectedItem = $defaultPrinter
    }
    Load-Settings
    
    # Konfigurera FileSystemWatcher händelser
    $script:fileWatcher.SynchronizingObject = $form
    $script:fileWatcher.IncludeSubdirectories = $true
    $script:fileWatcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, DirectoryName'
    $script:fileWatcher.add_Created({ Update-FileList })
    $script:fileWatcher.add_Deleted({ Update-FileList })
    $script:fileWatcher.add_Renamed({ Update-FileList })

    if ($args.Count -gt 0 -and $args[0] -match "^\d{6}$") {
        $textBox.Text = $args[0]
        Update-FileList
    }
})

$form.Add_FormClosing({
    param($sender, $e)
    if (-not (Prompt-SaveChanges)) {
        $e.Cancel = $true
        return
    }
    if ($form.WindowState -eq 'Normal') { Save-Settings }
    if ($script:fileWatcher) { $script:fileWatcher.Dispose() }
    Cleanup-TempImages
})

# Händelse för att rita objekten i fillistan manuellt
$fileListBox.add_DrawItem({
    param($sender, $e)
    if ($e.Index -lt 0) { return }

    $itemText = $fileListBox.Items[$e.Index].ToString()
    $isSelected = ($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected
    $isInfoTxt = $itemText -eq "info.txt"

    $bgColor = $null
    $textColor = $null

    if ($isSelected) {
        $bgColor = [System.Drawing.SystemBrushes]::Highlight
        $textColor = [System.Drawing.SystemColors]::HighlightText
    } elseif ($isInfoTxt) {
        $bgColor = New-Object System.Drawing.SolidBrush($theme.InfoColor)
        $textColor = [System.Drawing.SystemColors]::ControlText
    } else {
        $bgColor = New-Object System.Drawing.SolidBrush($fileListBox.BackColor)
        $textColor = [System.Drawing.SystemColors]::ControlText
    }

    $e.Graphics.FillRectangle($bgColor, $e.Bounds)
    [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $itemText, $theme.MonoFont, $e.Bounds, $textColor, [System.Windows.Forms.TextFormatFlags]::Default)
    $e.DrawFocusRectangle()

    if (-not $isSelected) { $bgColor.Dispose() }
})


$button.Add_Click({ 
    if (-not (Prompt-SaveChanges)) { return }
    Update-FileList 
})
$textBox.Add_KeyDown({ 
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { 
        if (-not (Prompt-SaveChanges)) { return }
        Update-FileList 
    }
})

$buttonOpenFolder.Add_Click({
    if (-not (Prompt-SaveChanges)) { return }
    $nummer = $textBox.Text
    if (-not ($nummer -match "^\d{6}$")) { return }
    $firstTwo = $nummer.Substring(0, 2)
    $folderPath = "F:\it\data\dokument\$firstTwo\$nummer"
    if (Test-Path $folderPath) {
        Invoke-Item -Path $folderPath
    }
})

$buttonNewWindow.Add_Click({
    if (-not (Prompt-SaveChanges)) { return }
    $localPath = "C:\IT\ITMapp.exe"
    try {
        Start-Process -FilePath $localPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Kunde inte starta en ny instans av programmet från '$localPath'.", "Fel", "OK", "Error")
    }
})

$script:webBrowser.Add_DocumentCompleted({
    if ($script:startEditingAfterLoad) {
        $script:startEditingAfterLoad = $false
        Toggle-TxtEditor -enable $true
        $script:webBrowser.Document.Body.Focus()
    }
})

$buttonInfo.Add_Click({
    if (-not (Prompt-SaveChanges)) { return }
    
    $nummer = $textBox.Text
    if (-not ($nummer -match "^\d{6}$")) { return }
    $firstTwo = $nummer.Substring(0, 2)
    $folderPath = "F:\it\data\dokument\$firstTwo\$nummer"
    $infoFilePath = Join-Path -Path $folderPath -ChildPath "info.txt"

    if (-not (Test-Path $infoFilePath)) {
        try {
            "" | Set-Content -Path $infoFilePath
            Update-FileList
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte skapa info.txt.", "Fel", "OK", "Error")
            return
        }
    }
    
    $script:startEditingAfterLoad = $true
    
    # Rensa befintliga val och välj endast info.txt
    $fileListBox.ClearSelected()
    $fileListBox.SelectedItem = "info.txt"
})

$buttonPrint.Add_Click({
    if (-not (Prompt-SaveChanges)) { return }

    if ($fileListBox.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Inga filer är valda för utskrift.", "Information", "OK", "Information")
        return
    }
    
    $chosenPrinter = $printerComboBox.SelectedItem
    if (-not $chosenPrinter) {
        [System.Windows.Forms.MessageBox]::Show("Ingen skrivare är vald.", "Information", "OK", "Information")
        return
    }
    
    $nummer = $textBox.Text; $firstTwo = $nummer.Substring(0, 2)
    $folderPath = "F:\it\data\dokument\$firstTwo\$nummer"

    Print-Files -FilesToPrint $fileListBox.SelectedItems -Printer $chosenPrinter -BaseFolderPath $folderPath
})

$buttonPrintAllPdf.Add_Click({
    if (-not (Prompt-SaveChanges)) { return }
    
    $chosenPrinter = $printerComboBox.SelectedItem
    if (-not $chosenPrinter) {
        [System.Windows.Forms.MessageBox]::Show("Ingen skrivare är vald.", "Information", "OK", "Information")
        return
    }
    
    $allPdfs = $fileListBox.Items | Where-Object { $_.EndsWith(".pdf", [System.StringComparison]::InvariantCultureIgnoreCase) }
    if ($allPdfs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Det finns inga PDF-filer i listan att skriva ut.", "Information", "OK", "Information")
        return
    }

    $nummer = $textBox.Text; $firstTwo = $nummer.Substring(0, 2)
    $folderPath = "F:\it\data\dokument\$firstTwo\$nummer"
    
    Print-Files -FilesToPrint $allPdfs -Printer $chosenPrinter -BaseFolderPath $folderPath
})


$btnSaveTxt.Add_Click({
    $nummer = $textBox.Text; $firstTwo = $nummer.Substring(0, 2)
    $folderPath = "F:\it\data\dokument\$firstTwo\$nummer"
    $infoFilePath = Join-Path -Path $folderPath -ChildPath "info.txt"
    $newContent = $script:webBrowser.Document.Body.InnerText
    try {
        Set-Content -Path $infoFilePath -Value $newContent
        $statusBar.Text = "info.txt har sparats."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Kunde inte spara filen: $($_.Exception.Message)", "Spara-fel", "OK", "Error")
    }
    Toggle-TxtEditor -enable $false
})

$btnCancelTxt.Add_Click({
    Toggle-TxtEditor -enable $false
    $fileListBox.add_SelectedIndexChanged.Target.Invoke($null, $null)
})

$fileListBox.Add_SelectedIndexChanged({
    param($sender, $e)
    if (-not (Prompt-SaveChanges)) { 
        $fileListBox.SelectedIndexChanged.Remove($fileListBox_SelectedIndexChanged)
        $fileListBox.SelectedItem = "info.txt"
        $fileListBox.SelectedIndexChanged.Add($fileListBox_SelectedIndexChanged)
        return 
    }
    Cleanup-TempImages
    
    if ($fileListBox.SelectedItems.Count -ne 1) {
        $script:webBrowser.Navigate("about:blank")
        if ($fileListBox.SelectedItems.Count -gt 1) {
             $statusBar.Text = "$($fileListBox.SelectedItems.Count) filer valda. Ingen förhandsgranskning."
        } else {
             $statusBar.Text = "Ingen fil vald."
        }
        return
    }
    
    $selectedFile = $fileListBox.SelectedItem
    if (-not $selectedFile) { return }

    $nummer = $textBox.Text; $firstTwo = $nummer.Substring(0, 2)
    $folderPath = "F:\it\data\dokument\$firstTwo\$nummer"
    $fullPath = Join-Path -Path $folderPath -ChildPath $selectedFile
    $extension = [System.IO.Path]::GetExtension($selectedFile).ToLower()
    
    $script:webBrowser.Navigate("about:blank")
    
    $imageExtensions = @('.jpg', '.jpeg', '.png', '.gif', '.bmp')

    if ($imageExtensions -contains $extension) {
        $script:webBrowser.DocumentText = "<html><body style='margin:0; background-color:#f0f0f0;'><img src='file:///$fullPath' style='width:100%; height:auto;'></body></html>"
        $statusBar.Text = "Visar bild: $selectedFile"
    } else {
        switch ($extension) {
            '.txt' {
                try {
                    $fileContent = Get-Content -Path $fullPath -Raw
                    $escapedContent = [System.Security.SecurityElement]::Escape($fileContent)
                    $script:webBrowser.DocumentText = "<pre style='font-family: Consolas, monospace; font-size: 10pt; white-space: pre-wrap; word-wrap: break-word;'>$escapedContent</pre>"
                    $statusBar.Text = "Visar: $selectedFile"
                } catch {
                    $script:webBrowser.DocumentText = "Kunde inte läsa textfilen: $($_.Exception.Message)"
                }
            }
            '.pdf' {
                $script:webBrowser.Navigate($fullPath)
                $statusBar.Text = "Visar PDF: $selectedFile"
            }
            '.msg' {
                $statusBar.Text = "Öppnar .msg-fil... (kräver Outlook)"
                $outlook = $null; $mail = $null; $weStartedOutlook = $false; $tempImageFolder = $null
                try {
                    try { $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } 
                    catch { $outlook = New-Object -ComObject Outlook.Application; $weStartedOutlook = $true }
                    $mail = $outlook.Session.OpenSharedItem($fullPath)
                    $htmlBody = $mail.HTMLBody
                    
                    if ($mail.Attachments.Count -gt 0) {
                        $tempImageFolder = Join-Path $env:TEMP ([System.Guid]::NewGuid().ToString())
                        New-Item -ItemType Directory -Path $tempImageFolder | Out-Null
                        $script:tempImageFolder = $tempImageFolder

                        foreach ($attachment in $mail.Attachments) {
                            try {
                                $cid = $attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                                if ($cid) {
                                    $tempFilePath = Join-Path $tempImageFolder $attachment.FileName
                                    $attachment.SaveAsFile($tempFilePath)
                                    $htmlBody = $htmlBody.Replace("cid:$cid", "file:///$tempFilePath")
                                }
                            } catch {}
                        }
                    }

                    $script:webBrowser.DocumentText = $htmlBody
                    $statusBar.Text = "Visar MSG: $selectedFile"
                } catch {
                    $script:webBrowser.DocumentText = "Förhandsgranskning av .msg-filer kräver Microsoft Outlook.`n`nFel: $($_.Exception.Message)"
                } finally {
                    if ($mail) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($mail) | Out-Null }
                    if ($weStartedOutlook -and $outlook) { $outlook.Quit() }
                    if ($outlook) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook) | Out-Null }
                    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
                }
            }
            Default {
                $script:webBrowser.DocumentText = "Förhandsgranskning är inte tillgänglig för filtypen '$extension'.`nDubbelklicka för att försöka öppna."
                $statusBar.Text = "Ingen förhandsgranskning tillgänglig."
            }
        }
    }
})

$fileListBox.Add_MouseDoubleClick({
    if ($fileListBox.SelectedItem) {
        $nummer = $textBox.Text; $firstTwo = $nummer.Substring(0, 2)
        $folderPath = "F:\it\data\dokument\$firstTwo\$nummer"
        $fullPath = Join-Path -Path $folderPath -ChildPath $fileListBox.SelectedItem
        try { Invoke-Item -Path $fullPath } catch {}
    }
})

# Händelse för att ta bort en fil
$fileListBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq 'Delete' -and $fileListBox.SelectedItems.Count -gt 0) {
        $filesToDelete = $fileListBox.SelectedItems -join ", `n"
        $result = [System.Windows.Forms.MessageBox]::Show("Är du säker på att du vill ta bort följande fil(er)?`n`n$filesToDelete", "Bekräfta borttagning", 'YesNo', 'Warning')
        
        if ($result -eq 'Yes') {
            $script:webBrowser.Navigate("about:blank")
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Milliseconds 100

            $itemsToRemove = @($fileListBox.SelectedItems)
            $nummer = $textBox.Text; $firstTwo = $nummer.Substring(0, 2)
            $folderPath = "F:\it\data\dokument\$firstTwo\$nummer"
            
            foreach($selectedFile in $itemsToRemove){
                $fullPath = Join-Path -Path $folderPath -ChildPath $selectedFile
                try {
                    Remove-Item -Path $fullPath -Force -ErrorAction Stop
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Kunde inte ta bort '$selectedFile'.`n`nFel: $($_.Exception.Message)", "Fel vid borttagning", "OK", "Error")
                }
            }
            
            Update-FileList
        }
    }
})

$fileListBox.Add_DragDrop({
    param($s, $e)
    # Hantera standard fil-drop
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $droppedFiles = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($textBox.Text -match "^\d{6}$") {
            $firstTwo = $textBox.Text.Substring(0, 2)
            $destinationFolder = "F:\it\data\dokument\$firstTwo\$($textBox.Text)"
            if (Test-Path $destinationFolder) {
                foreach ($file in $droppedFiles) { try { Copy-Item -Path $file -Destination $destinationFolder -Force } catch {}}
                Update-FileList
            }
        }
    } 
    # Hantera Outlook attachment/email drop
    elseif ($e.Data.GetDataPresent("FileGroupDescriptorW")) {
        $statusBar.Text = "Bearbetar objekt från Outlook..."
        if ($textBox.Text -match "^\d{6}$") {
            $firstTwo = $textBox.Text.Substring(0, 2)
            $destinationFolder = "F:\it\data\dokument\$firstTwo\$($textBox.Text)"
            if (Test-Path $destinationFolder) {
                $descriptorStream = $e.Data.GetData("FileGroupDescriptorW")
                $fileDescriptor = New-Object byte[]($descriptorStream.Length)
                $descriptorStream.Read($fileDescriptor, 0, $descriptorStream.Length) | Out-Null
                $descriptorStream.Close()

                $rawfileName = [System.Text.Encoding]::Unicode.GetString($fileDescriptor, 76, 520).TrimEnd("`0")
                # Rensa filnamn från ogiltiga tecken
                $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() -join ''
                $regex = "[$([regex]::Escape($invalidChars))]"
                $fileName = $rawfileName -replace $regex, '_'
                
                $filePath = Join-Path -Path $destinationFolder -ChildPath $fileName

                $contentStream = $e.Data.GetData("FileContents", $true)
                $fileBytes = New-Object byte[]($contentStream.Length)
                $contentStream.Read($fileBytes, 0, $contentStream.Length) | Out-Null
                [System.IO.File]::WriteAllBytes($filePath, $fileBytes)
                $contentStream.Close()
                Update-FileList
            }
        }
    }
})

# --- Starta Applikationen ---
try {
    # Koppla händelsen här så att den kan referera till sig själv
    $fileListBox_SelectedIndexChanged = [System.EventHandler]{ param($sender, $e)
        $this.Add_SelectedIndexChanged.Target.Invoke($sender, $e)
    }
    $form.ShowDialog() | Out-Null
    $form.Dispose()
} catch {
    # Fånga eventuella fel om $form inte skulle finnas (bör inte hända)
}
