<# ::

    REM Author : Leo Gillet - Freenitial on GitHub

    cls & @echo off & title MSI Properties Viewer
    copy /y "%~f0" "%TEMP%\%~n0.ps1" >NUL && powershell -Nologo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%TEMP%\%~n0.ps1"
    exit /b

#>


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


$loadingForm = New-Object System.Windows.Forms.Form -Property @{Text="MSI Properties Viewer";Size=[System.Drawing.Size]::new(300, 95);StartPosition="CenterScreen";FormBorderStyle="FixedDialog";ControlBox=$false}
$loadingForm.Show()
$loadingLabel = New-Object System.Windows.Forms.Label -Property @{Location=[System.Drawing.Point]::new(10, 5);Size=[System.Drawing.Size]::new(260, 20);Text="Loading interface..."}
$loadingForm.Controls.Add($loadingLabel)
$loadingForm.Refresh()
$launch_progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{Location=[System.Drawing.Point]::new(10, 25);Size=[System.Drawing.Size]::new(260, 20);Style="Continuous";Value=10}
$loadingForm.Controls.Add($launch_progressBar)
$loadingForm.Refresh()


$script:resizePending = $false
$script:lastWidth = @{}
$script:DarkMode = 0
$script:currentTabIndex = 0
$script:TargetPC = ""
$script:stopRequested = $false
$script:sortColumn = -1
$script:sortOrder = [System.Windows.Forms.SortOrder]::Ascending
$script:MSILoaded = $false
$script:fromBrowseButton = $false
$script:tab3Refreshed = $false


Add-Type -ReferencedAssemblies System.Windows.Forms.dll -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Collections;
using System.Runtime.InteropServices;
public class CustomForm : Form
{
    public event EventHandler<Message> OnWindowMessage;
    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);
        if (OnWindowMessage != null)
        {
            OnWindowMessage(this, m);
        }
    }
}
public class ListViewItemComparer : IComparer
{
    private int col;
    private SortOrder order;
    public ListViewItemComparer()
    {
        col = 0;
        order = SortOrder.Ascending;
    }
    public ListViewItemComparer(int column, SortOrder sortOrder)
    {
        col = column;
        order = sortOrder;
    }
    public int Compare(object x, object y)
    {
        int returnVal = -1;
        returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text,
                                   ((ListViewItem)y).SubItems[col].Text);
        if (order == SortOrder.Descending)
            returnVal *= -1;
        return returnVal;
    }
}
public class NativeMethods
{
    public const int SB_HORZ = 0; // Scroll horizontal
    public const int SB_VERT = 1;  // Scroll vertical
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetScrollPos(IntPtr hWnd, int nBar);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SetScrollPos(IntPtr hWnd, int nBar, int nPos, bool bRedraw);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);
    public const int WM_VSCROLL = 0x0115;
    public const int WM_HSCROLL = 0x0114;
    public const int SB_THUMBPOSITION = 4;
}
"@ -Language CSharp


$highlightBrush = [System.Drawing.SolidBrush]::new([System.Drawing.SystemColors]::Highlight)
$highlightTextBrush = [System.Drawing.SolidBrush]::new([System.Drawing.SystemColors]::HighlightText)

# Function that quick-create majority of controls
function gen {
    param(
        [Parameter(Mandatory = $true)] $container,
        [Parameter(Mandatory = $true)][string] $type,
        [Parameter(Mandatory = $true)][AllowEmptyString()] [string] $text,
        [Parameter(Mandatory = $true)][int] $x, [int] $y, [int] $width, [int] $height,
        [Parameter(ValueFromRemainingArguments = $true)] $additionalProps
    )
    $control = New-Object "System.Windows.Forms.$type" -Property @{
        Text     = $text
        Location = New-Object System.Drawing.Point($x, $y)
        Size     = New-Object System.Drawing.Size($width, $height)
    }
    if ($control -is [System.Windows.Forms.Button]) {
        # Additional button properties can be set here if needed
    } elseif ($control -is [System.Windows.Forms.TextBox]) {
        $control.BorderStyle = "FixedSingle"
    }
    if ($additionalProps) {
        foreach ($prop in $additionalProps) {
            $propName, $propValue = $prop -split '=', 2
            $propName  = $propName.Trim() ; $propValue = $propValue.Trim()
            if ($propValue -match '^New-Object' -or $propValue -match '^\@\{') {
                $control.$propName = Invoke-Expression $propValue
            } elseif ($propValue -match '^\[.*\]::') {
                $control.$propName = Invoke-Expression $propValue
            } elseif ($propName -eq "Font") {
                $fontParts = $propValue -split ',\s*'
                $fontName  = $fontParts[0].Trim('"')
                $fontSize  = [float]($fontParts[1].Trim())
                $fontStyle = if ($fontParts.Count -gt 2) {[System.Drawing.FontStyle]($fontParts[2].Trim() -replace '\s', '')} else {[System.Drawing.FontStyle]::Regular}
                $control.Font = New-Object System.Drawing.Font($fontName, $fontSize, $fontStyle)
            } elseif ($propName -eq "Padding") {
                $padValues       = $propValue -split '\s+'
                $control.Padding = New-Object System.Windows.Forms.Padding($padValues[0], $padValues[1], $padValues[2], $padValues[3])
            } elseif ($propName -in @("ForeColor", "BackColor", "BorderColor", "HoverColor")) {
                $colorValues = $propValue -split '\s+'
                $color = if ($colorValues.Count -eq 1) {[System.Drawing.ColorTranslator]::FromHtml($colorValues[0])} else {[System.Drawing.Color]::FromArgb($colorValues[0], $colorValues[1], $colorValues[2])}
                switch ($propName) {
                    "ForeColor"  { $control.ForeColor = $color }
                    "BackColor"  { $control.BackColor = $color }
                    "BorderColor" { $control.FlatAppearance.BorderColor = $color }
                    "HoverColor"  { $control.FlatAppearance.MouseOverBackColor = $color }
                }
            } elseif ($propName -eq "Anchor") {
                $enumType = [System.Windows.Forms.AnchorStyles]
                $enumValues = $propValue -split ',\s*'
                $combinedEnum = 0
                foreach ($value in $enumValues) { $combinedEnum = $combinedEnum -bor [System.Enum]::Parse($enumType, $value) }
                $control.$propName = $combinedEnum
            } elseif ($propName -eq "Dock") {
                $enumType = [System.Windows.Forms.DockStyle]
                $control.$propName = [System.Enum]::Parse($enumType, $propValue)
            } elseif ($propName -eq "BorderStyle") {
                $control.BorderStyle = $propValue
            } else {
                $control.$propName = switch ($propValue) {
                    '$true'  { $true }
                    '$false' { $false }
                    { $_ -match '^\d+$' } { [int]$_ }
                    default { $_ }
                }
            }
        }
    }
    if ($container -is [System.Windows.Forms.TabControl] -and $control -is [System.Windows.Forms.TabPage]) { [void]$container.TabPages.Add($control) } else { [void]$container.Controls.Add($control) }
    return $control
}


function Get-MsiProperty {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateScript({ Test-Path $_ })]
        [string[]] $Path,
        [Parameter(Mandatory = $true)]
        [string[]] $Properties
    )
    $results = @()
    $currentIndex = 0
    foreach ($CurrentPath in $Path) {
        $currentIndex++
        try {
            $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
            $MSIDatabase = $WindowsInstaller.OpenDatabase($CurrentPath, 0)
            foreach ($prop in $Properties) {
                try {
                    $Query = "SELECT Value FROM Property WHERE Property = '$($prop)'"
                    $View = $MSIDatabase.OpenView($Query)
                    $View.Execute()
                    $Record = $View.Fetch()
                    $value = if ($Record) { $Record.StringData(1) } else { "None" }
                    $results += [PSCustomObject]@{ Path=$CurrentPath ; Property=$prop ; Value=$value }
                    $View.Close()
                    if ($Record) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Record) }
                    if ($View) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($View) }
                }
                catch {
                    Write-Warning "Error while reading '$prop' in file '$CurrentPath' : $_"
                    $results += [PSCustomObject]@{ Path=$CurrentPath ; Property=$prop ; Value="None" }
                }
            }
        }
        catch {
            Write-Warning "Error opening msi File : '$CurrentPath' : $_"
            foreach ($prop in $Properties) { $results += [PSCustomObject]@{ Path=$CurrentPath ; Property=$prop ; Value="None" } }
        }
        finally {
            if ($MSIDatabase) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MSIDatabase) }
            if ($WindowsInstaller) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) }
        }
    }
    return $results
}


# Create main form
$screenWidth =  [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
$formWidth =  if ($screenWidth -lt 1250) { $screenWidth }  else { 1250 }
$formHeight = if ($screenHeight -lt 600) { $screenHeight } else { 600 }
$form = New-Object CustomForm
$form.Text = "MSI Properties Viewer"
$form.Size = New-Object System.Drawing.Size($formWidth, $formHeight)


# Create TabPages
$tabControl = gen $form "TabControl" "" 0 0 $form.ClientSize.Width $form.ClientSize.Height 'Anchor=Top,Bottom,Left,Right'
$tabPage1 = gen $tabControl "TabPage" "Drop a MSI" 0 0 0 0
$tabPage2 = gen $tabControl "TabPage" "Explore Folders" 0 0 0 0
$tabPage3 = gen $tabControl "TabPage" "Explore Registry" 0 0 0 0
#$tabPageTheme = gen $tabControl "TabPage" "Switch Theme" 0 0 0 0


$launch_progressBar.Value = 15


# Tab 1: "Drop a file"

$labels      = @()
$textBoxes   = @()
$copyButtons = @()
$properties = [ordered]@{
    "GUID"             = "ProductCode"
    "Product Name"     = "ProductName"
    "Product Version"  = "ProductVersion"
    "Manufacturer"     = "Manufacturer"
    "Upgrade Code"     = "UpgradeCode"
}


$yPos = 35
foreach ($key in $properties.Keys) {
    $label        = gen $tabPage1 "Label" "$key :" 10 $yPos 0 0 'AutoSize=$true'
    $labels      += $label
    $textBox      = gen $tabPage1 "TextBox" "" 101 $yPos 50 25 'ReadOnly=$true'
    $textBoxes   += $textBox
    $copyButton   = gen $tabPage1 "Button" "COPY" 0 0 60 25 'Enabled=$false'
    $copyButtons += $copyButton
    $copyButton.Add_Click({
        $buttonIndex = $copyButtons.IndexOf($this)
        [System.Windows.Forms.Clipboard]::SetText($textBoxes[$buttonIndex].Text)
    })
    $yPos += 30
}


$borderTop =        gen $tabPage1 "Panel" "" 0 0 0 1   'Dock=Top'    'BackColor=Gray' 
$separatorLine =    gen $tabPage1 "Panel" "" 0 ($yPos + 30) $tabPage1.ClientSize.Width 1 'BorderStyle=None' 'BackColor=Gray' 'AutoSize=$true'
$labelMsiPath =     gen $tabPage1 "Label" "MSI Path:" 10 $yPos 0 0 'AutoSize=$true' 
$yPos += 25
$labelDropMessage = gen $tabPage1 "Label" "Drop MSI here, or write path below" 0 0 0 0 'AutoSize=$true' 'Font=Arial,20'
$textBoxPath =      gen $tabPage1 "TextBox" "" 10 $yPos 260 25
$findGuidButton =   gen $tabPage1 "Button" "FIND GUID" 0 0 85 25
$browseButton   =   gen $tabPage1 "Button" "BROWSE" 0 0 85 25
$pictureBoxIcon  =  gen $tabPage1 "PictureBox" "" 0 0 32 32 'SizeMode=StretchImage' 'Visible=$false'
$labelFileName   =  gen $tabPage1 "Label" "" 0 0 0 0 'AutoSize=$true' 'Font=Arial,20' 'Visible=$false'


# Load a generic MSI icon from msiexec.exe
$iconMsi = [System.Drawing.Icon]::ExtractAssociatedIcon((Join-Path $env:windir "System32\msiexec.exe"))
$pictureBoxIcon.Image = $iconMsi.ToBitmap()


function MeasureTextWidth {
    param ([string]$text, [System.Drawing.Font]$font )
    $bitmap = New-Object System.Drawing.Bitmap(1, 1)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $sizeF = $graphics.MeasureString($text, $font)
    $graphics.Dispose()
    $bitmap.Dispose()
    return [int]$sizeF.Width
}


function Update-LabelAndIcon {
    param ([string]$filePath)
    $script:MSILoaded = $true
    $labelFileName.Text = [System.IO.Path]::GetFileName($filePath)
    $labelDropMessage.Visible = $false
    $pictureBoxIcon.Visible = $true
    $labelFileName.Visible = $true
    Update-ButtonPositions | Out-Null
}


function Update-ButtonPositions {
    $formWidth  = [int]$tabPage1.ClientSize.Width
    $formHeight = [int]$tabPage1.ClientSize.Height
    $textBoxWidth = $formWidth - 177
    foreach ($textBox in $textBoxes) { $textBox.Width = $textBoxWidth }
    $separatorLine.Width = $formWidth
    $buttonSpacing       = [int](($formWidth - 2 * $findGuidButton.Width) / 3)
    $browseButtonX       = $buttonSpacing
    $findGuidButtonX     = 2 * $buttonSpacing + $browseButton.Width
    $buttonY             = $formHeight - 40
    $browseButton.Location   = New-Object System.Drawing.Point($browseButtonX, $buttonY)
    $findGuidButton.Location = New-Object System.Drawing.Point($findGuidButtonX, $buttonY)
    $textBoxPath.Location    = New-Object System.Drawing.Point(10, ($buttonY - 30))
    $textBoxPath.Width       = $formWidth - 20
    for ($i = 0; $i -lt $copyButtons.Count; $i++) {
        $copyButtonX = $formWidth - 71
        $copyButtonY = $textBoxes[$i].Location.Y - 4
        $copyButtons[$i].Location = New-Object System.Drawing.Point($copyButtonX, $copyButtonY)
    }
    $labelMsiPathY = ($textBoxPath.Location.Y - 25)
    $labelMsiPath.Location = New-Object System.Drawing.Point(10, $labelMsiPathY)
    if ($pictureBoxIcon.Visible -eq $true -and $labelFileName.Visible -eq $true) {
        $iconWidth = [int]$pictureBoxIcon.Width
        $padding = 30
        $maxLabelWidth = $formWidth - $iconWidth - (2 * $padding) - 40
        $originalText = [System.IO.Path]::GetFileName($textBoxPath.Text)
        $shortenedText = $originalText
        while ((MeasureTextWidth -text $shortenedText -font $labelFileName.Font) -gt $maxLabelWidth) { $shortenedText = $shortenedText.Substring(0, $shortenedText.Length - 1) }
        if ($shortenedText.Length -lt $originalText.Length) { $shortenedText += "..." }
        $labelFileName.Text = $shortenedText
        $totalWidth = $iconWidth + 10 + (MeasureTextWidth -text $labelFileName.Text -font $labelFileName.Font)
        $posX = [int](($formWidth - $totalWidth) / 2)
        $posY = [int](($textBoxPath.Location.Y - $separatorLine.Location.Y - [math]::Max($pictureBoxIcon.Height, $labelFileName.Height)) / 2 + $separatorLine.Location.Y)
        $pictureBoxIcon.Location = New-Object System.Drawing.Point($posX, $posY)
        $calculatedX = $posX + $iconWidth + 10
        $calculatedY = $posY + [int](($pictureBoxIcon.Height - $labelFileName.Height) / 2)
        $labelFileName.Location = New-Object System.Drawing.Point($calculatedX, $calculatedY)
    } else {
        $labelDropMessageX = [int](($formWidth - $labelDropMessage.Width) / 2)
        $labelDropMessageY = [int](($textBoxPath.Location.Y - $separatorLine.Location.Y - $labelDropMessage.Height) / 2 + $separatorLine.Location.Y)
        $labelDropMessage.Location = New-Object System.Drawing.Point($labelDropMessageX, $labelDropMessageY)
    }
}


$tabPage1.Add_Resize({ Update-ButtonPositions })


$launch_progressBar.Value = 20


$tabPage1.AllowDrop = $true
$tabPage1.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [Windows.Forms.DragDropEffects]::Copy
    } else {
        $_.Effect = [Windows.Forms.DragDropEffects]::None
    }
})


$textBoxPath.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $findGuidButton.PerformClick()
    } elseif ($_.Control) {
        if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $textBoxPath.SelectAll()
            $_.Handled = $true
        }
    } else {
        $script:MSILoaded = $false
        $pictureBoxIcon.Visible = $false
        $labelFileName.Visible = $false
        $labelDropMessage.Visible = $true
        Update-ButtonPositions
    }
})


$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Supported files (*.msi;*.msix;*.mst;*.msp)|*.msi;*.msix;*.mst;*.msp|All files (*.*)|*.*"
    $openFileDialog.Title  = "Select a file"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFile = $openFileDialog.FileName
        if ($selectedFile -match "\.(msi|msix|mst|msp)$") {
            if ($selectedFile -ne $textBoxPath.Text) {
                foreach ($textBox in $textBoxes) { $textBox.Text = "" }
                foreach ($copyButton in $copyButtons) { $copyButton.Enabled = $false }
                $job = Start-Job -ScriptBlock { param($path) ; Start-Sleep -Milliseconds 100 ; return $path } -ArgumentList $selectedFile
                while ($job.State -eq 'Running') { Start-Sleep -Milliseconds 50 }
                $result = Receive-Job -Job $job
                Remove-Job -Job $job
                $script:fromBrowseButton = $true
                $textBoxPath.Text        = $result
                $findGuidButton.PerformClick()
            } else { [System.Windows.Forms.MessageBox]::Show("Preaching to the choir", "Same File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
        }
    }
})


$findGuidButton.Add_Click({
    $msiPath = $textBoxPath.Text.Trim('"')
    $propertyArray = $properties.Values
    $results = Get-MsiProperty -Path $msiPath -Properties $propertyArray
    $i = 0
    foreach ($key in $properties.Keys) {
        $value = ($results | Where-Object { $_.Property -eq $properties[$key] }).Value
        if ($value) { $textBoxes[$i].Text = $value.Trim() ; $copyButtons[$i].Enabled = $true }
        else { $textBoxes[$i].Text = "" ; $copyButtons[$i].Enabled = $false }
        $i++
    }
    if ($textBoxes[0].Text -ne "") {
        $findGuidButton.Enabled = $false
        $textBoxPath.Text = $msiPath
        Update-LabelAndIcon -filePath $msiPath
        $findGuidButton.Text = "GUID FOUND"
    } else {
        $findGuidButton.Enabled = $true
        $findGuidButton.Text = "FIND GUID"
    }
})


$tabPage1.Add_DragDrop({
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    foreach ($file in $files) {
        if ($file -match "\.(msi|msix|mst|msp)$") {
            $textBoxPath.Text = $file
            $findGuidButton.PerformClick()
            Update-LabelAndIcon -filePath $file
        }
    }
})


$textBoxPath.Add_TextChanged({
    if (-not $script:fromBrowseButton) { 
        foreach ($textBox in $textBoxes) {$textBox.Text=""}
        foreach ($copyButton in $copyButtons) {$copyButton.Enabled=$false}
        $findGuidButton.Text="FIND GUID"
    }
    $script:fromBrowseButton = $false
    if ($textBoxPath.Text.Trim() -eq "") {
        $findGuidButton.Enabled = $false
        $findGuidButton.Text="FIND GUID" 
    } else { 
        $findGuidButton.Enabled = $true 
    }
})


$launch_progressBar.Value = 25


# Tab 2

$splitContainer2 = New-Object System.Windows.Forms.SplitContainer
$splitContainer2.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer2.FixedPanel = [System.Windows.Forms.FixedPanel]::None
$splitContainer2.SplitterDistance = 250
$splitContainer2.Orientation = [System.Windows.Forms.Orientation]::Vertical
$splitContainer2.Panel1MinSize = 100
$splitContainer2.Panel2MinSize = 100
$tabPage2.Controls.Add($splitContainer2)


# Panels for the left side
$borderTop =         gen $tabPage2                "Panel"     ""                         0 0 0 1    'Dock=Top'    'BackColor=Gray' 
$panelLeftMain =     gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 0    'Dock=Fill'
$treeView =          gen $panelLeftMain           "TreeView"  ""                         0 0 0 0    'Dock=Fill'   'CheckBoxes=$true'
$panelLeftCtrls =    gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 20   'Dock=Top'
$panelLeftCtrls_B =  gen $panelLeftCtrls          "Panel"     ""                         0 0 0 20   'Dock=Bottom'
$panelLeftOptions =  gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 20   'Dock=Top'
$sep1 =              gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 10   'Dock=Top'
$sep2 =              gen $splitContainer2.Panel1  "Panel"     ""                         0 0 1 0    'Dock=Right'  'BackColor=Gray'
$sortComboBoxLabel = gen $panelLeftOptions        "Label"     "Sorting:"                 0 0 43 0   'Dock=Right'
$sortComboBox =      gen $panelLeftOptions        "ComboBox"  ""                         0 0 44 20  'Dock=Right'
   
$openBtn =           gen $panelLeftOptions        "Button"    "Open selected folder"     0 0 130 0  'Dock=Left'
$refreshBtn =        gen $panelLeftOptions        "Button"    "Refresh selected folder"  0 0 130 0  'Dock=Left'
$pathTextBox =       gen $panelLeftCtrls_B        "TextBox"   ""                         0 0 0 0    'Dock=Fill' 
$gotoButton =        gen $panelLeftCtrls_B        "Button"    "Goto"                     0 0 43 0   'Dock=Right'
$pathTextBoxLabel =  gen $panelLeftCtrls_B        "Label"     "Path: "                   0 0 0 0    'Dock=Left'   'Autosize=$true'
$panelLeftBottom =   gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 20   'Dock=Bottom'
   
$sep5 =              gen $panelLeftBottom         "Panel"     ""                         0 0 1 0    'Dock=Right'  'BackColor=Gray' 
$SearchMSIBtn =      gen $panelLeftBottom         "Button"    "Scan selected folder"     0 0 71 0   'Dock=Left'
$SearchCheckedBtn =  gen $panelLeftBottom         "Button"    "Scan checked folders"     0 0 75 0   'Dock=Left'
$recursionComboBox = gen $panelLeftBottom         "ComboBox"  ""                         0 0 36 20  'Dock=Left'
$recursionLabel =    gen $panelLeftBottom         "Label"     "Recursion:"               0 0 58 0   'Dock=Left'


$launch_progressBar.Value = 30

$recursionComboBox.Items.AddRange(@("No", "1", "2", "3", "4", "5", "All"))
$recursionComboBox.SelectedIndex = 6  # default = All
$sortComboBox.Items.AddRange(@("A-Z", "Z-A", "Old", "New"))
$sortComboBox.SelectedIndex = 0
$sortComboBox.Add_SelectedIndexChanged({ refreshTreeViewFolder })


$pathTextBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $gotoButton.PerformClick()
    } elseif ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $pathTextBox.SelectAll()
        $_.Handled = $true
    }
})


# Panels for the right side
$panelRightMain =     gen $splitContainer2.Panel2  "Panel"       ""          0 0 0 0  'Dock=Fill'
$panelRightCtrls =    gen $splitContainer2.Panel2  "Panel"       ""          0 0 0 33 'Dock=Top'
$panelRightCtrls_T =  gen $panelRightCtrls         "Panel"       ""          0 0 0 20 'Dock=Fill'
$showMspCheckbox =    gen $panelRightCtrls_T       "CheckBox"    "Show MSP"  0 0 0 0  'Dock=Right'  'Autosize=$true' 'Checked=$false'
$showMstCheckbox =    gen $panelRightCtrls_T       "CheckBox"    "Show MST"  0 0 0 0  'Dock=Right'  'Autosize=$true' 'Checked=$false'
$panelRightCtrls_B =  gen $panelRightCtrls         "Panel"       ""          0 0 0 20 'Dock=Bottom'
$searchTextBox =      gen $panelRightCtrls_B       "TextBox"     ""          0 0 0 0  'Dock=Fill'
$searchTextBoxLabel = gen $panelRightCtrls_B       "Label"       "Filter:"   0 0 0 0  'Dock=Left'   'Autosize=$true'
$listView_Explore =   gen $panelRightMain          "ListView"    ""          0 0 0 0  'Dock=Fill'   'View=Details'   'FullRowSelect=$true' 'GridLines=$true'  'AllowColumnReorder=$true'  'HideSelection=$false' 
$sep3 =               gen $splitContainer2.Panel2  "Panel"       ""          0 0 1 0  'Dock=Left'   'BackColor=Gray' 
$progressPanel =      gen $splitContainer2.Panel2  "Panel"       ""          0 0 0 20 'Dock=Bottom'


Function Update-ProgressBarWidth {
    $availableWidth = $statusStrip.Width - ($statusLabel.Width + $stopButton.Width + 5)
    if ($availableWidth - 5 -lt 0) { $availableWidth = 0 }
    $progressBar.Width = $availableWidth
}


$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.SizingGrip = $false
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "0 items"
$statusLabel.Font = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Regular)
$statusStrip.Items.Add($statusLabel) | Out-Null
$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Control.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$progressBar.AutoSize = $false
$statusStrip.Items.Add($progressBar) | Out-Null
$stopButton = New-Object System.Windows.Forms.ToolStripButton
$stopButton.Text = "STOP"
$stopButton.Enabled = $false
$stopButton.BackColor = [System.Drawing.Color]::LightGray
$statusStrip.Items.Add($stopButton) | Out-Null
$statusStrip.add_SizeChanged({ Update-ProgressBarWidth })
$statusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom
$progressPanel.Controls.Add($statusStrip)
$sep6 = gen $progressPanel "Panel" "" 0 0 1 0 'Dock=Left' 'BackColor=Gray'


$allListViewItemsExplore = New-Object System.Collections.ArrayList


$columns_listView_Explore = @("File Name", "GUID", "Version", "Path", "Weight", "Modified")
foreach ($col in $columns_listView_Explore) {
    $columnHeader = New-Object System.Windows.Forms.ColumnHeader
    $columnHeader.Text = $col
    [void]$listView_Explore.Columns.Add($columnHeader)
}


$searchTextBox.Add_KeyDown({ if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) { $searchTextBox.SelectAll() ; $_.Handled = $true } })


$launch_progressBar.Value = 35


# Event handler for custom drawing of nodes
$treeView.HideSelection = $false
$treeView.DrawMode = [System.Windows.Forms.TreeViewDrawMode]::OwnerDrawText  # Enable custom drawing
$treeView.Add_DrawNode({
    param($s,$e)
    if($e.Node.IsSelected){
        $bounds=$e.Bounds
        $e.Graphics.FillRectangle($highlightBrush,$bounds)
        $e.Graphics.DrawString($e.Node.Text,$treeView.Font,$highlightTextBrush,$bounds.X,$bounds.Y+1)
        $e.DrawDefault=$false
    }else{ $e.DrawDefault=$true }
})


function FilterListViewItems {
    param(
        [System.Windows.Forms.ListView]$listView,
        [string]$Mode,
        [bool]$showMsp = $false,
        [bool]$showMst = $false,
        [System.Collections.ArrayList]$registryPaths,
        [System.Collections.ArrayList]$allListViewItems
    )
    $visibleItems = @()
    switch ($Mode) {
        "Explore" {
            $listview_NameIndex = -1
            for ($i = 0; $i -lt $listView.Columns.Count; $i++) { if ($listView.Columns[$i].Text -in @("File Name", "DisplayName")) { $listview_NameIndex = $i ; break } }
            if ($listview_NameIndex -eq -1) {
                [System.Windows.Forms.MessageBox]::Show("Column 'File Name' or 'DisplayName' not found in the ListView", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            } 
            foreach ($item in $allListViewItems) {
                $listview_Name = $item.SubItems[$listview_NameIndex].Text.ToLower()
                $isMsp = $listview_Name.EndsWith(".msp", [System.StringComparison]::OrdinalIgnoreCase)
                $isMst = $listview_Name.EndsWith(".mst", [System.StringComparison]::OrdinalIgnoreCase)
                if (($showMsp -or -not $isMsp) -and ($showMst -or -not $isMst)) { $visibleItems += $item.Clone() }
            }
        }
        "Registry" {
            $pathColumnIndex = -1
            for ($i = 0; $i -lt $listView.Columns.Count; $i++) { if ($listView.Columns[$i].Text -eq "Registry Path") { $pathColumnIndex = $i ; break } }
            if ($pathColumnIndex -eq -1) {
                [System.Windows.Forms.MessageBox]::Show("'Registry Path' Column not found", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            foreach ($item in $allListViewItems) {
                $itemPath = $item.SubItems[$pathColumnIndex].Text
                $matchFound = $false
                foreach ($rp in $registryPaths) { if ($itemPath -like "$rp*") { $matchFound = $true ; break } }
                if ($matchFound) { $visibleItems += $item.Clone() }
            }
        }
        default {
            [System.Windows.Forms.MessageBox]::Show("Unknown Mode. Code using 'Explore' or 'Registry'", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
    $listView.Items.Clear()
    foreach ($item in $visibleItems) { [void]$listView.Items.Add($item) }
    $statusLabel.Text = "$($listView_Explore.Items.Count) items"
    Update-ProgressBarWidth
}


$showMspCheckbox.Add_CheckedChanged({ FilterListViewItems -Mode Explore -listView $listView_Explore -showMsp $showMspCheckbox.Checked -showMst $showMstCheckbox.Checked -allListViewItems $allListViewItemsExplore })
$showMstCheckbox.Add_CheckedChanged({ FilterListViewItems -Mode Explore -listView $listView_Explore -showMsp $showMspCheckbox.Checked -showMst $showMstCheckbox.Checked -allListViewItems $allListViewItemsExplore })


function AdjustListViewColumns {
    param([System.Windows.Forms.ListView]$listView)
    $SearchCheckedBtn.Width = ($splitContainer2.Panel1.Width - $recursionLabel.Width - $recursionComboBox.Width) / 2
    $SearchMSIBtn.Width = ($splitContainer2.Panel1.Width - $recursionLabel.Width - $recursionComboBox.Width) / 2
    $totalWidth = $listView.ClientSize.Width
    $columnsInfo = @()
    for ($i = 0; $i -lt $listView.Columns.Count; $i++) {
        $maxWidth = 0
        $headerText = $listView.Columns[$i].Text
        $headerWidth = [System.Windows.Forms.TextRenderer]::MeasureText($headerText, $listView.Font).Width + 8
        foreach ($item in $listView.Items) {
            if ($item.SubItems.Count -gt $i) {
                $textWidth = [System.Windows.Forms.TextRenderer]::MeasureText($item.SubItems[$i].Text, $listView.Font).Width
                if ($textWidth -gt $maxWidth) { $maxWidth = $textWidth }
            }
        }
        $maxWidth = [Math]::Max($maxWidth + 8, $headerWidth)
        $fixedWidth = switch ($headerText) { "GUID" { 250 } ; "Modified" { 80 } ; "InstallDate" { 65 } ; "Version" { 75 } ; "DisplayVersion" { 75 } ; default { $null } }
        $columnsInfo += [PSCustomObject]@{
            Index = $i
            MaxWidth = $maxWidth
            MinWidth = switch ($headerText) { "File Name" { 180 } ; "DisplayName" { 180 } ; default { 0 } }
            FixedWidth = $fixedWidth
            Width = 0
            IsMaxed = $false
        }
    }
    do {
        $remaining = $columnsInfo.Where({ -not $_.IsMaxed -and -not $_.FixedWidth })
        $maxedWidth = ($columnsInfo.Where({ $_.IsMaxed }) | Measure-Object -Property Width -Sum).Sum
        $fixedWidth = ($columnsInfo.Where({ $_.FixedWidth }) | Measure-Object -Property FixedWidth -Sum).Sum
        $remainingWidth = $totalWidth - $maxedWidth - $fixedWidth
        $widthChanged = $false
        foreach ($col in $remaining) {
            $newWidth = [int]($remainingWidth / $remaining.Count) - 4
            $newWidth = [Math]::Max($newWidth, $col.MinWidth)  # Apply minimal width
            if ($newWidth -ge $col.MaxWidth) {
                $col.Width = $col.MaxWidth
                $col.IsMaxed = $true
                $widthChanged = $true
            } else {
                $col.Width = $newWidth
            }
        }
    } while ($widthChanged)
    foreach ($col in $columnsInfo) { if ($col.FixedWidth) { $listView.Columns[$col.Index].Width = $col.FixedWidth } else { $listView.Columns[$col.Index].Width = $col.Width } }
}


function Get-MsiInfo {
    param ([string]$filePath)
    $defaultInfo = @{ GUID="None" ; Version="None" }
    if (-not $filePath.EndsWith(".msi", [System.StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path $filePath)) {
         return $defaultInfo 
    }
    try {
        $msiProperties = Get-MsiProperty -Path $filePath -Properties @("ProductCode", "ProductVersion")
        $guid = ($msiProperties | Where-Object { $_.Property -eq "ProductCode" }).Value
        $version = ($msiProperties | Where-Object { $_.Property -eq "ProductVersion" }).Value
        return @{ GUID = $guid ; Version = $version }
    }
    catch {
        Write-Warning "Error retrieving MSI properties for file '$filePath': $_"
        return $defaultInfo
    }
}


$launch_progressBar.Value = 40


# Initialize the TreeView
$rootNode = New-Object System.Windows.Forms.TreeNode
$rootNode.Text = "This Device"
$rootNode.Tag = "This Device"
$treeView.Nodes.Add($rootNode) | Out-Null

# Add "Fast Access" root node
$fastAccessNode = New-Object System.Windows.Forms.TreeNode
$fastAccessNode.Text = "Fast Access"
$fastAccessNode.Tag = "Fast Access"
$treeView.Nodes.Add($fastAccessNode) | Out-Null

# Populate "Fast Access" node with Quick Access folders
$shell = New-Object -ComObject Shell.Application
$quickAccess = $shell.Namespace('shell:::{679f85CB-0220-4080-B29B-5540CC05AAB6}')
$quickAccessItems = $quickAccess.Items()


foreach ($item in $quickAccessItems) {
    if ($item.IsFolder) {
        $node = New-Object System.Windows.Forms.TreeNode
        $node.Text = $item.Name
        $node.Tag = $item.Path
        $node.Nodes.Add([System.Windows.Forms.TreeNode]::new()) | Out-Null  # Add a dummy child node to enable expansion
        $fastAccessNode.Nodes.Add($node) | Out-Null
    }
}


$drives = Get-PSDrive -PSProvider 'FileSystem'
foreach ($drive in $drives) {
    $driveNode = New-Object System.Windows.Forms.TreeNode
    $driveNode.Text = $drive.Name + " (" + $drive.Root + ")"
    $driveNode.Tag = $drive.Root
    $driveNode.Nodes.Add([System.Windows.Forms.TreeNode]::new()) | Out-Null
    $rootNode.Nodes.Add($driveNode) | Out-Null
}
$rootNode.Expand()


$launch_progressBar.Value = 45


function SortDirectories {
    param (
        [System.IO.DirectoryInfo[]]$dirs,
        [string]$sortOption
    )
    switch ($sortOption) {
        "A-Z" { $dirs = $dirs | Sort-Object -Property Name }
        "Z-A" { $dirs = $dirs | Sort-Object -Property Name -Descending }
        "Old" { $dirs = $dirs | Sort-Object -Property LastWriteTime }
        "New" { $dirs = $dirs | Sort-Object -Property LastWriteTime -Descending }
    }
    return $dirs
}


function PopulateTree {
    param ([System.Windows.Forms.TreeNode]$parentNode, [string]$path)
    try {
        if ($path.StartsWith("\\")) {
            $start = $false
            cmd /c "net view $path" 2>&1 | ForEach-Object {
                if (!$start) { 
                    $start = $_ -match "^-{5,}"
                    return 
                }
                if ($_ -match "^(.+?)\s{2,}") {
                    $shareName = $matches[1].Trim()
                    $shareNode = New-Object System.Windows.Forms.TreeNode
                    $shareNode.Text = $shareName
                    $shareNode.Tag = Join-Path $parentNode.Tag $shareName
                    $shareNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())  # Add empty child to allow expansion
                    $parentNode.Nodes.Add($shareNode)
                }
            }
        }
        $dirs = Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue
        $dirs = SortDirectories -dirs $dirs -sortOption $sortComboBox.SelectedItem
        foreach ($dir in $dirs) {
            $node = New-Object System.Windows.Forms.TreeNode
            $node.Text = $dir.Name
            $node.Tag = $dir.FullName
            $hasReadAccess = $true
            $hasWriteAccess = $true
            $hasSubDirectories = $false
            try {  # Check if there is at least 1 subdolder - CPU optimized
                $enumerator = [System.IO.Directory]::EnumerateDirectories($dir.FullName).GetEnumerator()
                if ($enumerator.MoveNext()) { $hasSubDirectories = $true }
            } catch { $hasReadAccess = $false }
            if ($hasReadAccess) {
                try {
                    $writeAllow = $false
                    $writeDeny = $false
                    $acl = [System.IO.Directory]::GetAccessControl($dir.FullName)
                    $accessRules = $acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])
                    $user = [System.Security.Principal.WindowsIdentity]::GetCurrent()
                    $userSid = $user.User
                    foreach ($rule in $accessRules) {
                        if ($rule.IdentityReference -eq $userSid) {
                            if ($rule.AccessControlType -eq 'Allow') {
                                if (($rule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::Write) -ne 0) { $writeAllow = $true }
                            } elseif ($rule.AccessControlType -eq 'Deny') {
                                if (($rule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::Write) -ne 0) { $writeDeny = $true }
                            }
                        }
                    }
                    if ($writeDeny -and -not $writeAllow) { $hasWriteAccess = $false }
                } catch { $hasWriteAccess = $false }
                if (-not $hasWriteAccess) { $node.ForeColor = 'Orange' }
                if ($hasSubDirectories) { $node.Nodes.Add([System.Windows.Forms.TreeNode]::new()) }  # Add "+" if contains subfolders
            } else { $node.ForeColor = 'Red' } # Not adding "+" because acces denied
            $parentNode.Nodes.Add($node) | Out-Null
        }
    } catch {
        Write-Warning "Error accessing path ${path}: $_"
    }
}


function Expand-TreeViewPath {
    param([System.Windows.Forms.TreeView]$treeView, [string]$path)
    if (-not (Test-Path $path) -and ((cmd /c "net view $path" 2>&1)[0] -match "53")) {
        Show-NonBlockingMessage -message "Path not found" -title "Error" -timeout 2
        return
    }
    # Check if the path is a specific network path or a root
    $path = $path.TrimEnd('\')
    if ($path -match "^\\\\([^\\]+)\\?(.*)") {
        $serverName = $matches[1]
        $specificPath = $matches[2] # May be empty if it's the root
        # Add or find the "Network" node
        $networkRootNode = $treeView.Nodes | Where-Object { $_.Text -eq "Network" }
        if (-not $networkRootNode) {
            $networkRootNode = New-Object System.Windows.Forms.TreeNode
            $networkRootNode.Text = "Network"
            $networkRootNode.Tag = "Network"
            $treeView.Nodes.Add($networkRootNode)
        }
        # Add or find the server node
        $serverNode = $networkRootNode.Nodes | Where-Object { $_.Text -eq $serverName }
        if (-not $serverNode) {
            $serverNode = New-Object System.Windows.Forms.TreeNode
            $serverNode.Text = $serverName
            $serverNode.Tag = "\\$serverName"
            $networkRootNode.Nodes.Add($serverNode)
        }
        # Load all server shares if not already loaded
        if ($serverNode.Nodes.Count -eq 0) {
            $serverNode.Nodes.Clear()
            $start = $false
            cmd /c "net view \\$serverName" 2>&1 | ForEach-Object {
                if (!$start) { $start = $_ -match "^-{5,}"; return }
                if ($_ -match "^(.+?)\s{2,}") {
                    $shareName = $matches[1].Trim()
                    $childNode = New-Object System.Windows.Forms.TreeNode
                    $childNode.Text = $shareName
                    $childNode.Tag = "\\$serverName\$shareName"
                    $childNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
                    $serverNode.Nodes.Add($childNode)
                }
            }
            # Apply sorting to shares
            try {
                $sortedNodes = SortDirectories -dirs $serverNode.Nodes -sortOption $sortComboBox.SelectedItem
                $serverNode.Nodes.Clear()
                foreach ($sortedNode in $sortedNodes) { $serverNode.Nodes.Add($sortedNode) }
            } catch {
                Write-Warning "Sorting not applicable to some nodes under \\$serverName. Ignoring sorting for these nodes."
            }
        }
        # Expand the network root
        $networkRootNode.Expand()
        # If specific path provided, explore subfolders
        if ($specificPath) {
            $segments = $specificPath -split '\\'
            $currentNode = $serverNode
            foreach ($segment in $segments) {
                # Find or add the segment in the current node
                $childNode = $currentNode.Nodes | Where-Object { $_.Text -eq $segment }
                if (-not $childNode) {
                    $childNode = New-Object System.Windows.Forms.TreeNode
                    $childNode.Text = $segment
                    $childNode.Tag = Join-Path -Path $currentNode.Tag -ChildPath $segment
                    $childNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
                    $currentNode.Nodes.Add($childNode)
                }
                $currentNode = $childNode
                # Load subfolders if necessary
                if ($currentNode.Nodes.Count -eq 1 -and -not $currentNode.Nodes[0].Tag) {
                    $currentNode.Nodes.Clear()
                    try {
                        $dirs = Get-ChildItem -Path $currentNode.Tag -Directory -ErrorAction SilentlyContinue
                        $dirs = SortDirectories -dirs $dirs -sortOption $sortComboBox.SelectedItem
                        foreach ($dir in $dirs) {
                            $subChildNode = New-Object System.Windows.Forms.TreeNode
                            $subChildNode.Text = $dir.Name
                            $subChildNode.Tag = $dir.FullName
                            $subChildNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
                            $currentNode.Nodes.Add($subChildNode)
                        }
                    } catch {
                        Write-Warning "Unable to access or sort directories under $($currentNode.Tag)."
                    }
                }
            }
            # Select the last node without expanding
            $treeView.SelectedNode = $currentNode
            $currentNode.EnsureVisible()
        } else {
            # Select the server node if no specific path
            $treeView.SelectedNode = $serverNode
            $serverNode.EnsureVisible()
        }
        return
    }
    # Handle local paths
    $path = [System.IO.Path]::GetFullPath($path).TrimEnd('\')
    $segments = $path -split '\\' | Where-Object { $_ -ne "" }
    $driveRoot = "$($segments[0][0]):\"
    $currentNode = $null
    foreach ($node in $treeView.Nodes[0].Nodes) {
        if ($node.Tag -eq $driveRoot) {
            $currentNode = $node
            break
        }
    }
    if (-not $currentNode) {
        [System.Windows.Forms.MessageBox]::Show("Drive not found in treeview.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    # Build the hierarchy for remaining segments
    for ($i = 1; $i -lt $segments.Count; $i++) {
        $segment = $segments[$i]
        $found = $false
        # Dynamically load subfolders if the node hasn't been expanded
        if ($currentNode.Nodes.Count -eq 1 -and -not $currentNode.Nodes[0].Tag) {
            $currentNode.Nodes.Clear()
            try {
                $dirs = Get-ChildItem -Path $currentNode.Tag -Directory -ErrorAction SilentlyContinue
                $dirs = SortDirectories -dirs $dirs -sortOption $sortComboBox.SelectedItem
                foreach ($dir in $dirs) {
                    $childNode = New-Object System.Windows.Forms.TreeNode
                    $childNode.Text = $dir.Name
                    $childNode.Tag = $dir.FullName
                    $childNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
                    $currentNode.Nodes.Add($childNode)
                }
            } catch {
                Write-Warning "Sorting not applicable or directory inaccessible for path: $($currentNode.Tag)."
            }
        }
        # Search for the segment in loaded subfolders
        foreach ($childNode in $currentNode.Nodes) {
            if ($childNode.Text -eq $segment) {
                $currentNode = $childNode
                $found = $true
                break
            }
        }
        # If the segment doesn't exist yet, create it
        if (-not $found) {
            $newNode = New-Object System.Windows.Forms.TreeNode
            $newNode.Text = $segment
            $newNode.Tag = Join-Path -Path $currentNode.Tag -ChildPath $segment
            $newNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
            $currentNode.Nodes.Add($newNode)
            $currentNode = $newNode
        }
    }
    $treeView.SelectedNode = $currentNode
    $currentNode.EnsureVisible()
}


$launch_progressBar.Value = 50


function OpenTreeViewSelectedFolder {
    $selectedNode = $treeView.SelectedNode
    if ($null -ne $selectedNode) {
        if ($selectedNode.Text -eq "This Device" -or $selectedNode.Text -eq "Fast Access") {
            [System.Windows.Forms.MessageBox]::Show("Cannot open special node: $($selectedNode.Text).", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        if ($selectedNode.Tag -is [string] -and $selectedNode.Tag.StartsWith("\\")) {
            Start-Process -FilePath "explorer.exe" -ArgumentList "`"$($selectedNode.Tag)`""
            return
        }
        if ($selectedNode.Tag -is [string] -and (Test-Path $selectedNode.Tag)) {
            Start-Process -FilePath "explorer.exe" -ArgumentList "`"$($selectedNode.Tag)`""
        } else {
            [System.Windows.Forms.MessageBox]::Show("Path does not exist: $($selectedNode.Tag)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a node.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


$openBtn.Add_Click({ param($s, $e) ; OpenTreeViewSelectedFolder })


function refreshTreeViewFolder {
    $selectedNode = $treeView.SelectedNode
    if ($null -ne $selectedNode) {
        $scrollPosVert = [NativeMethods]::GetScrollPos($treeView.Handle, [NativeMethods]::SB_VERT)
        try {
            if (-not [string]::IsNullOrEmpty($selectedNode.Tag) -and $selectedNode.Parent -ne $null) {
                $selectedNode.Nodes.Clear()
                PopulateTree -parentNode $selectedNode -path $selectedNode.Tag
                $selectedNode.Expand()
            } else { 
                [System.Windows.Forms.MessageBox]::Show("Cannot refresh because node is invalid.") 
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error refreshing tree node: $_") 
        } finally {
            [NativeMethods]::SetScrollPos($treeView.Handle, [NativeMethods]::SB_VERT, $scrollPosVert, $true)
            [NativeMethods]::SendMessage($treeView.Handle, [NativeMethods]::WM_VSCROLL, ([NativeMethods]::SB_THUMBPOSITION -bor ($scrollPosVert -shl 16)), 0)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a node to refresh.")
    }
}


$refreshBtn.Add_Click({ refreshTreeViewFolder })


$launch_progressBar.Value = 55


function Get-FilesRecursive {
    param(
        [string[]]$paths,                          # Accepts one or multiple paths
        [int]$depth,                               # Search Recursion level
        [switch]$foldersOnly,                      # Do not scan files
        $progressBar,                              # Accepts both ProgressBar and ToolStripProgressBar
        [ref]$allItems,                            # Optional: reference to external results array
        [ref]$progressCounter,                     # Tracks progress globally
        [ref]$totalTasks,                          # Total number of tasks for progress calculation
        [System.Diagnostics.Stopwatch]$stopwatch   # Stopwatch to track elapsed time
    )
    if (-not $allItems) { $allItems = [ref](New-Object 'System.Collections.Generic.List[Object]') } # Initialize List
    if (-not $progressCounter) { $progressCounter = [ref]0 }
    if (-not $totalTasks) { $totalTasks = [ref](Get-ChildItem -Path $paths -Recurse -Directory -ErrorAction SilentlyContinue | Measure-Object).Count }
    if (-not $stopwatch) { $stopwatch = [System.Diagnostics.Stopwatch]::StartNew() }
    if ($depth -eq 0) { return }
    $fileFilterRegex = '\.(msi|msix|mst|msp)$'
    $updateFrequency = [Math]::Max([Math]::Ceiling($totalTasks.Value / 100), 1) # Update every ~1% or at least once
    foreach ($path in $paths) {
        if ($script:stopRequested) { break }
        if (-not [System.IO.Directory]::Exists($path)) { continue }
        try {
            if ($foldersOnly) {
                $dirs = [System.IO.Directory]::GetDirectories($path)
                foreach ($dirPath in $dirs) {
                    if ($script:stopRequested) { break }
                    $dirInfo = New-Object System.IO.DirectoryInfo($dirPath)
                    $allItems.Value.Add($dirInfo)
                    $progressCounter.Value++
                    if ($progressCounter.Value % $updateFrequency -eq 0) { UpdateProgressAndStatus -progressBar $progressBar -progressCounter $progressCounter -totalTasks $totalTasks -stopwatch $stopwatch }
                    Get-FilesRecursive -paths @($dirPath) -depth ($depth - 1) -foldersOnly -progressBar $progressBar -allItems $allItems -progressCounter $progressCounter -totalTasks $totalTasks -stopwatch $stopwatch
                }
            } else {
                # Process files in the current directory
                $allFiles = [System.IO.Directory]::GetFiles($path)
                foreach ($filePath in $allFiles) {
                    if ($script:stopRequested) { break }
                    if ([regex]::IsMatch($filePath, $fileFilterRegex, 'IgnoreCase')) {
                        $fileInfo = New-Object System.IO.FileInfo($filePath)
                        $allItems.Value.Add($fileInfo)
                    }
                }
                # Process subdirectories
                $dirs = [System.IO.Directory]::GetDirectories($path)
                foreach ($dirPath in $dirs) {
                    if ($script:stopRequested) { break }
                    $progressCounter.Value++
                    if ($progressCounter.Value % $updateFrequency -eq 0) { UpdateProgressAndStatus -progressBar $progressBar -progressCounter $progressCounter -totalTasks $totalTasks -stopwatch $stopwatch }
                    Get-FilesRecursive -paths @($dirPath) -depth ($depth - 1) -progressBar $progressBar -allItems $allItems -progressCounter $progressCounter -totalTasks $totalTasks -stopwatch $stopwatch
                }
            }
        }
        catch {
            Write-Warning "Error accessing path: $path. $_"
        }
    }

    return $allItems.Value
}


function UpdateProgressAndStatus {
    param($progressBar, [ref]$progressCounter, [ref]$totalTasks, [System.Diagnostics.Stopwatch]$stopwatch )
    Update-ProgressBar -progressBar $progressBar -currentProgress $progressCounter.Value -totalTasks $totalTasks.Value -inverted
    $elapsedTime = $stopwatch.Elapsed.TotalSeconds
    $remainingItems = $totalTasks.Value - $progressCounter.Value
    $itemsProcessedPerSecond = if ($elapsedTime -gt 0) { [Math]::Max($progressCounter.Value / $elapsedTime, 0.01) } else { 1 }
    $estimatedRemainingTime = [Math]::Ceiling($remainingItems / $itemsProcessedPerSecond)
    $timeRemainingText = if ($estimatedRemainingTime -gt 0) { "$([TimeSpan]::FromSeconds($estimatedRemainingTime).ToString('hh\:mm\:ss')) remaining" } else { "Calculating..." }
    $statusLabel.Text = "Calculating - $timeRemainingText"
    Update-ProgressBarWidth
}


function Update-ProgressBar {
    param($progressBar, [int]$currentProgress, [int]$totalTasks, [switch]$inverted)
    if ($progressBar -and $totalTasks -gt 0) {
        $progressValue = if ($inverted) { 100 - [Math]::Min(($currentProgress / $totalTasks) * 100, 100) } else { [Math]::Min(($currentProgress / $totalTasks) * 100, 100) }
        if ($progressBar -is [System.Windows.Forms.ToolStripProgressBar]) {$progressBar.Value=$progressValue;$progressBar.GetCurrentParent().Refresh()} else {$progressBar.Value=$progressValue;$progressBar.Refresh()}
    }
    [System.Windows.Forms.Application]::DoEvents()
}


# Convert ComboBox selection to depth
function Get-RecursionDepth { switch ($recursionComboBox.SelectedItem) { "No" { return 1 } ; "All" { return [int]::MaxValue } ; default { return [int]$recursionComboBox.SelectedItem + 1 } } }


$treeView.Add_BeforeExpand({ 
    $node = $_.Node
    if ($node.Nodes.Count -eq 1 -and -not $node.Nodes[0].Tag) {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        $node.Nodes.Clear()
        PopulateTree $node $node.Tag 
        $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor 
    }
})
#$treeView.Add_AfterCheck({ param($s, $e) ; $node = $e.Node ; foreach ($child in $node.Nodes) { $child.Checked = $node.Checked } })


function Get-CheckedNodes {
    param([System.Windows.Forms.TreeNodeCollection]$nodes)
    foreach ($node in $nodes) {
        if ($node.Checked) { $node }
        if ($node.Nodes.Count -gt 0) { Get-CheckedNodes -nodes $node.Nodes }
    }
}


$launch_progressBar.Value = 60


function Complete-Listview {
    param([bool]$multiSearch = $false)
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $script:stopRequested = $false
    $listView_Explore.Items.Clear()
    $allListViewItemsExplore.Clear()
    $recursionDepth = Get-RecursionDepth
    $finalPathsToSearch = [System.Collections.Generic.List[string]]::new()  # Initialize the list to store paths
    $allItems = [System.Collections.Generic.List[Object]]::new()  # Ensure $allItems is initialized before any calls
    $allItemsRef = [ref]$allItems
    if (-not $multiSearch) {  # Single search mode
        $selectedNode = $treeView.SelectedNode
        if ($null -eq $selectedNode) {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Select a node before search.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $pathTextBox.Text = $selectedNode.Tag
        if (-not [string]::IsNullOrEmpty($selectedNode.Tag)) { $finalPathsToSearch.Add($selectedNode.Tag) } else { Write-Warning "Selected node has no valid tag: $($selectedNode.Text)" }
    } else {  # Multi-search mode
        $checkedNodes = Get-CheckedNodes -nodes $treeView.Nodes
        if (-not $checkedNodes) {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Check some nodes before search", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        foreach ($node in $checkedNodes) { if (-not [string]::IsNullOrEmpty($node.Tag)) { $finalPathsToSearch.Add($node.Tag) } else { Write-Warning "Node has no valid tag: $($node.Text)" } }
    }
    if ($finalPathsToSearch.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No valid paths to search", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }
    $stopButton.Enabled = $true
    [System.Windows.Forms.Application]::DoEvents()
    $pathsArray = $finalPathsToSearch.ToArray()
    $items = Get-FilesRecursive -paths $pathsArray -depth $recursionDepth -allItems $allItemsRef -progressBar $progressBar
    $allFiles = $allItems | Where-Object { -not $_.PSIsContainer }
    if ($allFiles.Count -gt 0) {
        $totalFiles = $allFiles.Count
        $currentFile = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()  # Global stopwatch for total elapsed time
        $lastUpdate = [System.Diagnostics.Stopwatch]::StartNew()  # Timer for 100ms updates
        foreach ($file in $allFiles) {
            if ($script:stopRequested) { break }
            $currentFile++
            $fullPath = $file.FullName
            $sizeBytes = $file.Length
            $sizeMB = [math]::Round($sizeBytes / 1MB, 2).ToString() + " Mo"
            $modified = $file.LastWriteTime.ToString()
            $msiInfo = Get-MsiInfo -filePath $file.FullName
            $dataToAdd = @{ "File Name"=$file.Name ; "GUID"=$msiInfo.GUID ; "Version"=$msiInfo.Version ; "Path"=$fullPath ; "Weight"=$sizeMB ; "Modified"=$modified }
            $columnOrder = $listView_Explore.Columns | ForEach-Object { $_.Text }
            $subItems = @()
            foreach ($columnName in $columnOrder) { if ($dataToAdd.ContainsKey($columnName)) { $subItems += [string]$dataToAdd[$columnName] } else { $subItems += "" } }
            $firstItemText = $subItems[0]  # Create ListView with first sub-item
            $item = New-Object System.Windows.Forms.ListViewItem ($firstItemText)
            for ($i = 1; $i -lt $subItems.Count; $i++) { $item.SubItems.Add($subItems[$i]) | Out-Null }  # Add remaining sub-elements
            $item.Tag = $file.FullName  # Full path in tag
            $listView_Explore.Items.Add($item)
            $allListViewItemsExplore.Add($item)
            if ($lastUpdate.ElapsedMilliseconds -ge 100) { # Update progress and status every 100ms
                $lastUpdate.Restart()
                $progressBar.Value = [Math]::Min(($currentFile / $totalFiles) * 100, 100)
                $elapsedTime = $stopwatch.Elapsed.TotalSeconds
                $itemsProcessedPerSecond = if ($elapsedTime -gt 0) { [Math]::Max($currentFile / $elapsedTime, 0.01) } else { 1 }
                $remainingFiles = $totalFiles - $currentFile
                $estimatedRemainingTime = [Math]::Ceiling($remainingFiles / $itemsProcessedPerSecond)
                $timeRemainingText = if ($estimatedRemainingTime -gt 0) { "$([TimeSpan]::FromSeconds($estimatedRemainingTime).ToString('hh\:mm\:ss')) remaining" } else { "Calculating..." }
                $statusLabel.Text = "$($listView_Explore.Items.Count) items - $timeRemainingText"
                Update-ProgressBarWidth
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    }
    FilterListViewItems -Mode Explore -listView $listView_Explore -showMsp $showMspCheckbox.Checked -showMst $showMstCheckbox.Checked -allListViewItems $allListViewItemsExplore
    AdjustListViewColumns -listView $listView_Explore
    $progressBar.Value = 0
    $stopButton.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}


$SearchCheckedBtn.Add_Click({ Complete-Listview -multiSearch $true })
$SearchMSIBtn.Add_Click({ Complete-Listview -multiSearch $false })


$gotoButton.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $path = $pathTextBox.Text
    Expand-TreeViewPath -treeView $treeView -path $path
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})


$searchTextBox.Add_TextChanged({
    $filter = $searchTextBox.Text.ToLower()
    $listView_Explore.BeginUpdate()
    $listView_Explore.Items.Clear()
    $visibleItems = @()
    foreach ($item in $allListViewItemsExplore) {
        $matchText = $false
        foreach ($subItem in $item.SubItems) {
            if ($subItem.Text.ToLower() -like "*$filter*") {
                $matchText = $true
                break
            }
        }
        if ($matchText) {
            $listview_NameIndex = -1
            for ($i = 0; $i -lt $listView_Explore.Columns.Count; $i++) {
                if ($listView_Explore.Columns[$i].Text -eq "File Name") {
                    $listview_NameIndex = $i
                    break
                }
            }
            if ($listview_NameIndex -ge 0) {
                $listview_Name = $item.SubItems[$listview_NameIndex].Text.ToLower()
                $isMsp = $listview_Name.EndsWith(".msp", [System.StringComparison]::OrdinalIgnoreCase)
                $isMst = $listview_Name.EndsWith(".mst", [System.StringComparison]::OrdinalIgnoreCase)
                if (($showMspCheckbox.Checked -or -not $isMsp) -and ($showMstCheckbox.Checked -or -not $isMst)) { $visibleItems += $item.Clone() }
            }
        }
    }
    foreach ($item in $visibleItems) { [void]$listView_Explore.Items.Add($item) }
    $statusLabel.Text = "$($listView_Explore.Items.Count) items"
    Update-ProgressBarWidth
    $listView_Explore.EndUpdate()
})


$stopButton.Add_Click({ $script:stopRequested = $true })


$listView_Explore.Add_ColumnClick({
    param($s, $e)
    if ($e.Column -eq $script:sortColumn) { # Toggle sort order
        if ($script:sortOrder -eq [System.Windows.Forms.SortOrder]::Ascending) { $script:sortOrder = [System.Windows.Forms.SortOrder]::Descending }
        else { $script:sortOrder = [System.Windows.Forms.SortOrder]::Ascending }
    }
    else { $script:sortColumn=$e.Column ; $script:sortOrder=[System.Windows.Forms.SortOrder]::Ascending } # New column, set to ascending
    $listView_Explore.ListViewItemSorter = New-Object ListViewItemComparer($script:sortColumn, $script:sortOrder)
    $listView_Explore.Sort()
})


$launch_progressBar.Value = 65


function ConfigureTreeViewContextMenu($treeView) {
    # Right click
    $treeView.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            $hitTestInfo = $treeView.HitTest($e.X, $e.Y)
            if ($null -ne $hitTestInfo.Node) { $treeView.SelectedNode = $hitTestInfo.Node }
        }
    })
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $treeView.ContextMenuStrip = $contextMenu
    # Button "Copy Folder Path"
    $copyFolderPathMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyFolderPathMenuItem.Text = "Copy Folder Path"
    $copyFolderPathMenuItem.Add_Click({
        param($s, $e)
        $selectedNode = $treeView.SelectedNode
        if ($null -ne $selectedNode) {
            $path = if ($selectedNode.Tag -match '\s') { "`"$($selectedNode.Tag)`"" } else { $selectedNode.Tag }
            [System.Windows.Forms.Clipboard]::SetText($path)
        }
    })
    $contextMenu.Items.Add($copyFolderPathMenuItem) | Out-Null
    # Button "Open Folder"
    $openFolderMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $openFolderMenuItem.Text = "Open Folder"
    $openFolderMenuItem.Add_Click({
        param($s, $e)
        OpenTreeViewSelectedFolder
    })
    $contextMenu.Items.Add($openFolderMenuItem) | Out-Null
    # Button "Open Parent Folder"
    $openParentFolderMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $openParentFolderMenuItem.Text = "Open Parent Folder"
    $openParentFolderMenuItem.Add_Click({
        param($s, $e)
        $selectedNode = $treeView.SelectedNode
        if ($null -ne $selectedNode -and (Test-Path $selectedNode.Tag)) {
            $parentPath = Split-Path -Path $selectedNode.Tag -Parent
            if (Test-Path $parentPath) {
                Start-Process -FilePath "explorer.exe" -ArgumentList "/select,$parentPath"
            } else {
                [System.Windows.Forms.MessageBox]::Show("Cannot open parent folder: Path does not exist.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $contextMenu.Items.Add($openParentFolderMenuItem) | Out-Null
    # Button "Refresh Folder"
    $refreshFolderMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $refreshFolderMenuItem.Text = "Refresh Folder"
    $refreshFolderMenuItem.Add_Click({
        param($s, $e)
        refreshTreeViewFolder
    })
    $contextMenu.Items.Add($refreshFolderMenuItem) | Out-Null
    # Button "Scan Folder"
    $scanFolderMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $scanFolderMenuItem.Text = "Scan Folder"
    $scanFolderMenuItem.Add_Click({
        param($s, $e)
        $selectedNode = $treeView.SelectedNode
        if ($null -ne $selectedNode -and (Test-Path $selectedNode.Tag)) { Complete-Listview -multiSearch $false }
    })
    $contextMenu.Items.Add($scanFolderMenuItem) | Out-Null
    # Button "CMD here"
    $cmdHereMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $cmdHereMenuItem.Text = "CMD here"
    $cmdHereMenuItem.Add_Click({
        param($s, $e)
        $selectedNode = $treeView.SelectedNode
        if ($null -ne $selectedNode -and (Test-Path $selectedNode.Tag)) {
            Start-Process -FilePath "cmd.exe" -ArgumentList "/k cd /d `"$($selectedNode.Tag)`""
        } else {
            [System.Windows.Forms.MessageBox]::Show("Cannot open CMD here: Invalid path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $contextMenu.Items.Add($cmdHereMenuItem) | Out-Null
    # Button "CMD Admin here"
    $cmdAdminHereMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $cmdAdminHereMenuItem.Text = "CMD Admin here"
    $cmdAdminHereMenuItem.Add_Click({
        param($s, $e)
        $selectedNode = $treeView.SelectedNode
        if ($null -ne $selectedNode -and (Test-Path $selectedNode.Tag)) {
            $script = "cd /d `"$($selectedNode.Tag)`""
            Start-Process -FilePath "cmd.exe" -ArgumentList "/k $script" -Verb RunAs
        } else {
            [System.Windows.Forms.MessageBox]::Show("Cannot open CMD Admin here: Invalid path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $contextMenu.Items.Add($cmdAdminHereMenuItem) | Out-Null
}
ConfigureTreeViewContextMenu -treeView $treeView


$launch_progressBar.Value = 70


# Tab 3: Explore Registry

$panelMain = New-Object System.Windows.Forms.Panel
$panelMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabPage3.Controls.Add($panelMain)

$sep2_tab3                = gen $tabPage3               "Panel"    ""       0 0 1 0    'Dock=Fill' 'BackColor=Gray'
$sep1_tab3                = gen $tabPage3               "Panel"    ""       0 0 0 10   'Dock=Top'
$borderTop_tab3           = gen $tabPage3               "Panel"    ""       0 0 0 1    'Dock=Top'  'BackColor=Gray' 
$listView_Registry        = gen $panelMain              "ListView" ""       0 0 0 0    'Dock=Fill' 'View=Details' 'FullRowSelect=$true' 'GridLines=$true' 'AllowColumnReorder=$true' 'HideSelection=$false'
$panelCtrls_Registry      = gen $panelMain              "Panel"    ""       0 0 0 60   'Dock=Top'
        
$subPanelCtrls_PC         = gen $panelCtrls_Registry    "Panel"    ""       0 0 230 40 'Dock=Left'
$subPanelCtrls_PC_line1   = gen $subPanelCtrls_PC       "Panel"    ""       0 0 0 20   'Dock=Top'
$subPanelCtrls_PC_line2   = gen $subPanelCtrls_PC       "Panel"    ""       0 0 0 20   'Dock=Bottom'
$LoadReg_label            = gen $subPanelCtrls_PC_line1 "Label"    "Device Name Target (empty = This Device)" 0 0 0 0 'Dock=Left' 'Autosize=$true'
$LoadReg_btn              = gen $subPanelCtrls_PC_line2 "Button"   "Target" 0 0 55 0    'Dock=Left'
$LoadReg_textbox          = gen $subPanelCtrls_PC_line2 "TextBox"  ""       0 0 169 20 'Dock=Left' 'Autosize=$true'

$sep6_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"    ""       0 0 25 0   'Dock=Left'
$sep5_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"    ""       0 0 1 0    'Dock=Left' 'BackColor=Gray'
$sep4_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"    ""       0 0 25 0   'Dock=Left'
$subPanelCtrls_HKCU       = gen $panelCtrls_Registry    "Panel"    ""       0 0 450 40 'Dock=Left'
$subPanelCtrls_HKCU_line1 = gen $subPanelCtrls_HKCU     "Panel"    ""       0 0 0 20   'Dock=Top'
$subPanelCtrls_HKCU_line2 = gen $subPanelCtrls_HKCU     "Panel"    ""       0 0 0 20   'Dock=Bottom'

$sep3_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"    ""       0 0 25 0   'Dock=Left'
$sep2_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"    ""       0 0 1 0    'Dock=Left' 'BackColor=Gray'
$sep1_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"    ""       0 0 25 0   'Dock=Left'
$subPanelCtrls_HKLM       = gen $panelCtrls_Registry    "Panel"    ""       0 0 450 40 'Dock=Left'
$subPanelCtrls_HKLM_line1 = gen $subPanelCtrls_HKLM     "Panel"    ""       0 0 0 20   'Dock=Top'
$subPanelCtrls_HKLM_line2 = gen $subPanelCtrls_HKLM     "Panel"    ""       0 0 0 20   'Dock=Bottom'
$subPanelCtrls_Filter     = gen $panelCtrls_Registry    "Panel"    ""       0 0 0 20   'Dock=Bottom'

$HKLM32_btn               = gen $subPanelCtrls_HKLM_line2  "Button"   "Show"      0 0 0 0 'Dock=Right' 'Autosize=$true'
$checkbox_HKLM32          = gen $subPanelCtrls_HKLM_line2  "CheckBox" "HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" 0 0 410 0 'Dock=Left' 'checked=$true'
$HKLM64_btn               = gen $subPanelCtrls_HKLM_line1  "Button"   "Show"      0 0 0 0 'Dock=Right'  'Autosize=$true'
$checkbox_HKLM64          = gen $subPanelCtrls_HKLM_line1  "CheckBox" "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall"             0 0 330 0 'Dock=Left' 'checked=$true'

$HKCU32_btn               = gen $subPanelCtrls_HKCU_line2  "Button"   "Show"      0 0 0 0 'Dock=Right' 'Autosize=$true'
$checkbox_HKCU32          = gen $subPanelCtrls_HKCU_line2  "CheckBox" "HKCU\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" 0 0 410 0 'Dock=Left' 'checked=$true'
$HKCU64_btn               = gen $subPanelCtrls_HKCU_line1  "Button"   "Show"      0 0 0 0 'Dock=Right'  'Autosize=$true'
$checkbox_HKCU64          = gen $subPanelCtrls_HKCU_line1  "CheckBox" "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall"             0 0 330 0 'Dock=Left' 'checked=$true'

$searchTextBox_Registry   = gen $subPanelCtrls_Filter "TextBox"  ""                   0 0 0 20 'Dock=Fill' 
$refreshButton_Registry   = gen $subPanelCtrls_Filter "Button"   "Refresh"            0 0 55 0 'Dock=Right'
$FilterLabel_Registry     = gen $subPanelCtrls_Filter "Label"    "Filter:"            0 0 0 0  'Dock=Left'   'Autosize=$true'

$columnsRegistry = @("DisplayName", "GUID", "InstallSource", "UninstallString", "InstallDate", "DisplayVersion", "Registry Path")
$allListViewItems_Registry = New-Object System.Collections.ArrayList


foreach ($btn in @($HKLM32_btn, $HKLM64_btn, $HKCU32_btn, $HKCU64_btn)) {
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize = 1
}


$LoadReg_textbox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $LoadReg_btn.PerformClick()
    } elseif ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $LoadReg_textbox.SelectAll()
        $_.Handled = $true
    }
})


$searchTextBox_Registry.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $refreshButton_Registry.PerformClick()
    } elseif ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $searchTextBox_Registry.SelectAll()
        $_.Handled = $true
    }
})


$launch_progressBar.Value = 75


Function Initialize-RegistryListView {
    param( [System.Windows.Forms.ListView]$listViewparam )
    foreach ($col in $columnsRegistry) {
        $columnHeader = New-Object System.Windows.Forms.ColumnHeader
        $columnHeader.Text = $col
        [void]$listViewparam.Columns.Add($columnHeader)
    }
}
Initialize-RegistryListView -listViewparam $listView_Registry


$LoadReg_btn.Add_Click({
    $script:TargetPC = $LoadReg_textbox.Text.Trim()
    Update-RegistryListView -refresh $true -restoreScroll $false
})


Function PopulateRegistryListView {
    param(
        [System.Windows.Forms.ListView]$listViewparam,
        [System.Collections.ArrayList]$registryPaths,
        [System.Collections.ArrayList]$allListViewItems
    )
    if ([string]::IsNullOrEmpty($script:TargetPC) -eq $false) {
        if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            $result = [System.Windows.Forms.MessageBox]::Show("Restart as Administrator?", "Permission Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Nologo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`"" -Verb RunAs
                $form.Close()
            } else { return }
        }
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $listViewparam.Items.Clear()
    $allListViewItems.Clear()
    foreach ($registryPath in $registryPaths) {
        $rootKey, $subPath = $registryPath -split ":", 2
        if ($rootKey -eq "HKLM") {$baseKey=[Microsoft.Win32.RegistryHive]::LocalMachine} elseif ($rootKey -eq "HKCU") {$baseKey=[Microsoft.Win32.RegistryHive]::CurrentUser} else {continue}
        if ([string]::IsNullOrEmpty($script:TargetPC)) {
            $registryKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($baseKey, [Microsoft.Win32.RegistryView]::Default).OpenSubKey($subPath)
        } else {
            try { $registryKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($baseKey, $script:TargetPC, [Microsoft.Win32.RegistryView]::Default).OpenSubKey($subPath) } catch { continue }
        }
        if ($null -eq $registryKey) { continue }
        foreach ($subKeyName in $registryKey.GetSubKeyNames()) {
            $subKey = $registryKey.OpenSubKey($subKeyName)
            if ($null -ne $subKey ) {
                $displayName = $subKey.GetValue("DisplayName")
                if ([string]::IsNullOrEmpty($displayName)) { continue }
                $uninstallString = $subKey.GetValue("UninstallString")
                $guid = $null
                if ($uninstallString -match '\{[0-9A-F-]+\}') { $guid = $matches[0] }
                $installSource = $subKey.GetValue("InstallSource")
                $installDate = $subKey.GetValue("InstallDate")
                $displayVersion = $subKey.GetValue("DisplayVersion")
                $regPath = "$registryPath\$subKeyName"
                $dataToAdd = @{"DisplayName"=$displayName;"GUID"=$guid;"InstallSource"=$installSource;"UninstallString"=$uninstallString;"InstallDate"=$installDate;"DisplayVersion"=$displayVersion;"Registry Path"=$regPath}
                $item = New-Object System.Windows.Forms.ListViewItem
                $item.Tag = $regPath
                $sortedKeys = $listViewparam.Columns | Select-Object -ExpandProperty Text
                $firstKey = $sortedKeys[0]
                $item.Text = $dataToAdd[$firstKey]
                foreach ($key in $sortedKeys | Where-Object { $_ -ne $firstKey }) {
                    $value = $dataToAdd[$key]
                    if ($null -eq $value) { $value = "" }
                    $item.SubItems.Add($value)
                }
                $listViewparam.Items.Add($item)
                $allListViewItems.Add($item)
            }
        }
    }
    AdjustListViewColumns -listView $listViewparam
    $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor
}


Function Open-RegeditHere {
    param([System.Windows.Forms.ListView]$listViewparam, [string]$defaultPath, [bool]$parent=$false)
    $selectedItems = $listViewparam.SelectedItems
    if ($selectedItems.Count -gt 0 -and $parent -eq $false) { $regPath = $selectedItems[0].Tag } else { $regPath = $defaultPath }
    $regPath = $regPath -replace 'HKLM:', 'HKEY_LOCAL_MACHINE\' -replace 'HKCU:', 'HKEY_CURRENT_USER\' -replace 'HKCU\\', 'HKEY_CURRENT_USER\' -replace 'HKLM\\', 'HKEY_LOCAL_MACHINE\'
    $cmdAddReg = "REG ADD `"HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit`" /v `"LastKey`" /d `"$regPath`" /f"
    $cmdStartRegedit = "start regedit"
    $isRegeditRunning = Get-WmiObject Win32_Process -Filter "Name = 'regedit.exe'" -ErrorAction SilentlyContinue
    if ($null -ne $isRegeditRunning) {
        $taskKillResult = Start-Process -FilePath "taskkill.exe" -ArgumentList "/f", "/im", "regedit.exe", -NoNewWindow -PassThru -Wait | Out-Null
        if ($taskKillResult.ExitCode -ne 0) {
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c taskkill /f /im regedit.exe & $cmdAddReg & $cmdStartRegedit" -Verb RunAs
            return
        }
    }
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "$cmdAddReg & $cmdStartRegedit"
}


$launch_progressBar.Value = 80


function Update-RegistryListView {
	param([bool]$refresh = $false, [bool]$restoreScroll = $true)
    $selectedPaths = New-Object System.Collections.ArrayList
    if ($checkbox_HKCU64.Checked) { $selectedPaths.Add("HKCU:$($checkbox_HKCU64.Text.Substring(5))") }
    if ($checkbox_HKCU32.Checked) { $selectedPaths.Add("HKCU:$($checkbox_HKCU32.Text.Substring(5))") }
    if ($checkbox_HKLM64.Checked) { $selectedPaths.Add("HKLM:$($checkbox_HKLM64.Text.Substring(5))") }
    if ($checkbox_HKLM32.Checked) { $selectedPaths.Add("HKLM:$($checkbox_HKLM32.Text.Substring(5))") }
    if ($refresh -eq $true) {
        $scrollPosVert = [NativeMethods]::GetScrollPos($listView_Registry.Handle, [NativeMethods]::SB_VERT)
        PopulateRegistryListView -listViewparam $listView_Registry -registryPaths $selectedPaths -allListViewItems $allListViewItems_Registry
    }
    FilterListViewItems -Mode Registry -listView $listView_Registry -registryPaths $selectedPaths -allListViewItems $allListViewItems_Registry
    if ($refresh -eq $true -and $restoreScroll -eq $true) {
        [NativeMethods]::SetScrollPos($listView_Registry.Handle, [NativeMethods]::SB_VERT, $scrollPosVert, $true)
        $listView_Registry.Refresh()
        $scrollMessage = ([NativeMethods]::SB_THUMBPOSITION -bor ($scrollPosVert -shl 16))
        [NativeMethods]::SendMessage($listView_Registry.Handle, [NativeMethods]::WM_VSCROLL, $scrollMessage, 0)

    }
}


$checkbox_HKCU64.Add_CheckedChanged({ Update-RegistryListView})
$checkbox_HKCU32.Add_CheckedChanged({ Update-RegistryListView})
$checkbox_HKLM64.Add_CheckedChanged({ Update-RegistryListView})
$checkbox_HKLM32.Add_CheckedChanged({ Update-RegistryListView})


$HKLM32_btn.Add_Click({ Open-RegeditHere -listViewparam $listView_Registry -defaultPath $checkbox_HKLM32.Text -parent $true })
$HKLM64_btn.Add_Click({ Open-RegeditHere -listViewparam $listView_Registry -defaultPath $checkbox_HKLM64.Text -parent $true })
$HKCU32_btn.Add_Click({ Open-RegeditHere -listViewparam $listView_Registry -defaultPath $checkbox_HKCU32.Text -parent $true })
$HKCU64_btn.Add_Click({ Open-RegeditHere -listViewparam $listView_Registry -defaultPath $checkbox_HKCU64.Text -parent $true })


$refreshButton_Registry.Add_Click({ Update-RegistryListView -refresh $true})


Function Add-SearchFunctionality {
    param ([System.Windows.Forms.TextBox]$searchTextBox, [System.Windows.Forms.ListView]$listViewparam, [System.Collections.ArrayList]$allListViewItems)
    $eventHandler = {  # Capture local variables with closure
        param($s, $e)
        $filter = $searchTextBox.Text
        $listViewparam.BeginUpdate()
        $listViewparam.Items.Clear()
        foreach ($item in $allListViewItems) { if ($item.SubItems | Where-Object { $_.Text -like "*$filter*" }) { $listViewparam.Items.Add($item) } }
        $listViewparam.EndUpdate()
    }.GetNewClosure()
    $searchTextBox.add_TextChanged($eventHandler)
}
Add-SearchFunctionality -searchTextBox $searchTextBox_Registry -listViewparam $listView_Registry -allListViewItems $allListViewItems_Registry


Function HandleColumnClick {
    param ([System.Windows.Forms.ListView]$listViewparam, [System.Windows.Forms.ColumnClickEventArgs]$e)
    if ($e.Column -eq $script:sortColumn) {
        $script:sortOrder = if ($script:sortOrder -eq [System.Windows.Forms.SortOrder]::Ascending) {[System.Windows.Forms.SortOrder]::Descending} else {[System.Windows.Forms.SortOrder]::Ascending}
    } else {
        $script:sortColumn = $e.Column
        $script:sortOrder = [System.Windows.Forms.SortOrder]::Ascending
    }
    $listViewparam.ListViewItemSorter = New-Object ListViewItemComparer($script:sortColumn, $script:sortOrder)
    $listViewparam.Sort()
}
$listView_Registry.Add_ColumnClick({ HandleColumnClick -listViewparam $listView_Registry -e $_ })


function ConfigureListViewContextMenu($listView) {
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $contextMenu.Tag = $listView
    # =========================================== REGISTRY TAB 3
    $registryPathColumn = $listView.Columns | Where-Object { $_.Text -eq "Registry Path" }
    if ($null -ne $registryPathColumn) {
        # Button "Open Regedit here"
        $openRegeditMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $openRegeditMenuItem.Text = "Open Regedit here"
        $openRegeditMenuItem.Add_Click({
            param($s, $e)
            $contextMenu = $s.GetCurrentParent()
            $listView = $contextMenu.Tag
            if ($listView.SelectedItems.Count -gt 0) {
                $registryPathColumn = $listView.Columns | Where-Object { $_.Text -eq "Registry Path" }
                $regPath = $listView.SelectedItems[0].SubItems[$registryPathColumn.Index].Text
                Open-RegeditHere -listViewparam $listView -defaultPath $regPath
            }
        })
        $contextMenu.Items.Add($openRegeditMenuItem) | Out-Null
        # Separator
        $separator = New-Object System.Windows.Forms.ToolStripSeparator
        $contextMenu.Items.Add($separator) | Out-Null
    }
    # "Option Always Include Name"
    $alwaysIncludeNameMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $alwaysIncludeNameMenuItem.Text = "OPTION: Include Name"
    $alwaysIncludeNameMenuItem.CheckOnClick = $true
    $contextMenu.Items.Add($alwaysIncludeNameMenuItem) | Out-Null
    # Button "Copy Full Row"
    $copyAllMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyAllMenuItem.Text = "Copy Full Row"
    $copyAllMenuItem.Add_Click({
        param($s, $e)
        $contextMenu = $s.GetCurrentParent()
        $listView = $contextMenu.Tag
        $selectedItems = $listView.SelectedItems
        $values = @()
        foreach ($item in $selectedItems) {
            $lineValues = @()
            foreach ($column in $listView.Columns) {
                $colValue = $item.SubItems[$column.Index].Text
                if (-not [string]::IsNullOrWhiteSpace($colValue) -and $colValue -ne "None") { $lineValues += $colValue }
            }
            if ($lineValues.Count -gt 0) { $values += ($lineValues -join "`t") }
        }
        $textToCopy = $values -join [Environment]::NewLine
        if (-not [string]::IsNullOrWhiteSpace($textToCopy)) { [System.Windows.Forms.Clipboard]::SetText($textToCopy) }
    })
    $contextMenu.Items.Add($copyAllMenuItem) | Out-Null
    # Buttons for all listview
    foreach ($column in $listView.Columns) {
        $menuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $menuItem.Text = "Copy $($column.Text)"
        $menuItem.Tag = $column.Index
        $menuItem.Add_Click({
            param($s, $e)
            $contextMenu = $s.GetCurrentParent()
            $listView = $contextMenu.Tag
            $alwaysIncludeNameMenuItem = $contextMenu.Items | Where-Object { $_.Text -eq "OPTION: Include Name" }
            if ($listView.SelectedItems.Count -gt 0) {
                $index = $s.Tag
                $values = @()
                $fileNameColumn = $null
                foreach ($col in $listView.Columns) { if ($col.Text -in @("File Name", "DisplayName")) { $fileNameColumn = $col ; break } }
                foreach ($item in $listView.SelectedItems) {
                    $lineValue = $item.SubItems[$index].Text
                    if ($alwaysIncludeNameMenuItem.Checked -and $null -ne $fileNameColumn) {
                        $fileNameIndex = $fileNameColumn.Index
                        $fileName = $item.SubItems[$fileNameIndex].Text
                        if (-not [string]::IsNullOrWhiteSpace($fileName)) { $lineValue = "$fileName - $lineValue" }
                    }
                    $values += $lineValue
                }
                $textToCopy = $values -join [Environment]::NewLine
                if (-not [string]::IsNullOrWhiteSpace($textToCopy)) { [System.Windows.Forms.Clipboard]::SetText($textToCopy) }
            }
        })
        $contextMenu.Items.Add($menuItem) | Out-Null
    }
    $separator2 = New-Object System.Windows.Forms.ToolStripSeparator
    $contextMenu.Items.Add($separator2) | Out-Null
    # Button "Export" only for selected rows
    $exportMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exportMenuItem.Text = "Export"
    $exportMenuItem.Add_Click({ param($s, $e) ; $contextMenu = $s.GetCurrentParent() ; $listView = $contextMenu.Tag ; ListExport -listview $listView -all $false })
    $contextMenu.Items.Add($exportMenuItem) | Out-Null
    # Button "Export All" to export full listview
    $exportAllMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exportAllMenuItem.Text = "Export All"
    $exportAllMenuItem.Add_Click({ param($s, $e) ; $contextMenu = $s.GetCurrentParent() ; $listView = $contextMenu.Tag ; ListExport -listview $listView -all $true })
    $contextMenu.Items.Add($exportAllMenuItem) | Out-Null
    # =========================================== EXPLORE TAB 2
    # If "Path" column exists
    $pathColumn = $listView.Columns | Where-Object { $_.Text -eq "Path" }
    if ($null -ne $pathColumn) {
        $separator = New-Object System.Windows.Forms.ToolStripSeparator
        $contextMenu.Items.Add($separator) | Out-Null
        # Button "Copy File"
        $copyFileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $copyFileMenuItem.Text = "Copy File"
        $copyFileMenuItem.Add_Click({
            param($s, $e)
            $contextMenu = $s.GetCurrentParent()
            $listView = $contextMenu.Tag
            $pathColumn = $listView.Columns | Where-Object { $_.Text -eq "Path" }
            if ($null -ne $pathColumn -and $listView.SelectedItems.Count -gt 0) {
                $paths = $listView.SelectedItems | ForEach-Object { $_.SubItems[$pathColumn.Index].Text }
                [System.Windows.Forms.Clipboard]::SetFileDropList((New-Object System.Collections.Specialized.StringCollection -Property @{ AddRange = $paths }))
            }
        })
        $contextMenu.Items.Add($copyFileMenuItem) | Out-Null
        # Button "Open Parent Folder"
        $openParentMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $openParentMenuItem.Text = "Open Parent Folder"
        $openParentMenuItem.Add_Click({
            param($s, $e)
            $contextMenu = $s.GetCurrentParent()
            $listView = $contextMenu.Tag
            $pathColumn = $listView.Columns | Where-Object { $_.Text -eq "Path" }
            if ($null -ne $pathColumn -and $listView.SelectedItems.Count -gt 0) {
                $paths = $listView.SelectedItems | ForEach-Object { $_.SubItems[$pathColumn.Index].Text }
                foreach ($path in $paths) { Start-Process -FilePath "explorer.exe" -ArgumentList "/select,$path" }
            }
        })
        $contextMenu.Items.Add($openParentMenuItem) | Out-Null
        # Button "Open File"
        $openFileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $openFileMenuItem.Text = "Open File"
        $openFileMenuItem.Add_Click({
            param($s, $e)
            $contextMenu = $s.GetCurrentParent()
            $listView = $contextMenu.Tag
            $pathColumn = $listView.Columns | Where-Object { $_.Text -eq "Path" }
            if ($null -ne $pathColumn -and $listView.SelectedItems.Count -gt 0) {
                $paths = $listView.SelectedItems | ForEach-Object { $_.SubItems[$pathColumn.Index].Text }
                foreach ($path in $paths) { Start-Process -FilePath $path }
            }
        })
        $contextMenu.Items.Add($openFileMenuItem) | Out-Null
        # Separator
        $separator = New-Object System.Windows.Forms.ToolStripSeparator
        $contextMenu.Items.Add($separator) | Out-Null
        # Button "CMD here"
        $cmdHereMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $cmdHereMenuItem.Text = "CMD here"
        $cmdHereMenuItem.Add_Click({
            param($s, $e)
            $contextMenu = $s.GetCurrentParent()
            $listView = $contextMenu.Tag
            $pathColumn = $listView.Columns | Where-Object { $_.Text -eq "Path" }
            if ($null -ne $pathColumn -and $listView.SelectedItems.Count -gt 0) {
                $parentPath = [System.IO.Path]::GetDirectoryName($listView.SelectedItems[0].SubItems[$pathColumn.Index].Text)
                Start-Process -FilePath "cmd.exe" -ArgumentList "/k cd /d `"$parentPath`""
            }
        })
        $contextMenu.Items.Add($cmdHereMenuItem) | Out-Null
        # Button "CMD Admin here"
        $cmdAdminHereMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $cmdAdminHereMenuItem.Text = "CMD Admin here"
        $cmdAdminHereMenuItem.Add_Click({
            param($s, $e)
            $contextMenu = $s.GetCurrentParent()
            $listView = $contextMenu.Tag
            $pathColumn = $listView.Columns | Where-Object { $_.Text -eq "Path" }
            if ($null -ne $pathColumn -and $listView.SelectedItems.Count -gt 0) {
                $parentPath = [System.IO.Path]::GetDirectoryName($listView.SelectedItems[0].SubItems[$pathColumn.Index].Text)
                $script = "cd /d `"$parentPath`""
                Start-Process -FilePath "cmd.exe" -ArgumentList "/k $script" -Verb RunAs
            }
        })
        $contextMenu.Items.Add($cmdAdminHereMenuItem) | Out-Null
    }
    # =========================================== REGISTRY TAB 3
    if ($null -ne $registryPathColumn) {
        # Separator
        $separator3 = New-Object System.Windows.Forms.ToolStripSeparator
        $contextMenu.Items.Add($separator3) | Out-Null
        # Button "Delete Key" double confirmation if several elements
        $deleteKeyMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $deleteKeyMenuItem.Text = "Delete Key"
        $deleteKeyMenuItem.Add_Click({
            param($s, $e)
            $contextMenu = $s.GetCurrentParent()
            $listView = $contextMenu.Tag
            $selectedCount = $listView.SelectedItems.Count
            if ($selectedCount -le 0) { [System.Windows.Forms.MessageBox]::Show("No key selected for deletion.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) ; return}
            $skipConfirmation = [System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift
            if (-not $skipConfirmation) {
                $message = if ($selectedCount -gt 1) { "You are about to delete $selectedCount keys. Are you sure?" } else { "Are you sure you want to delete this key?" }
                $response = [System.Windows.Forms.MessageBox]::Show($message, "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($response -ne [System.Windows.Forms.DialogResult]::Yes) { return }
                if ($selectedCount -gt 1) {
                    $response = [System.Windows.Forms.MessageBox]::Show("Are you absolutely sure you want to delete these $selectedCount keys?", "Final Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    if ($response -ne [System.Windows.Forms.DialogResult]::Yes) { return }
                }
            }
            foreach ($item in $listView.SelectedItems) {
                $regPath = $item.SubItems[($listView.Columns | Where-Object { $_.Text -eq "Registry Path" }).Index].Text
                $normalizedPath = $regPath -replace '^HKLM:', 'HKEY_LOCAL_MACHINE\' -replace '^HKCU:', 'HKEY_CURRENT_USER\'
                try {
                    $subKeys = Get-ChildItem -Path "HKLM:\${normalizedPath}" -ErrorAction Stop
                    if ($subKeys.Count -gt 0) { $subKeys | ForEach-Object { [System.Windows.Forms.MessageBox]::Show("Sub-key: $_", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) } }
                } catch { [System.Windows.Forms.MessageBox]::Show("Error while checking sub-keys: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
                $hasPermission = $true
                try {
                    $acl = Get-Acl -Path $regPath -ErrorAction Stop
                    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
                    $hasPermission = $false
                    $requiredPermissions = [System.Security.AccessControl.RegistryRights]::FullControl -bor [System.Security.AccessControl.RegistryRights]::WriteKey -bor [System.Security.AccessControl.RegistryRights]::Delete
                    foreach ($access in $acl.Access) {
                        if ($access.IdentityReference -eq $currentUser) {
                            $permissions = $access.RegistryRights
                            if ($access.AccessControlType -eq "Deny") { $hasPermission = $false ; break }
                            if (($permissions -band $requiredPermissions) -eq $requiredPermissions) { $hasPermission = $true ; break }
                        }
                    }
                } catch { Write-Host "Error while retrieving ACL for '$normalizedPath': $_" ; $hasPermission = $false }
                try {
                    $processParams = @{ FilePath = "reg.exe" ; ArgumentList = "delete", $normalizedPath, "/f" }
                    if (-not $hasPermission) { $processParams.Verb = "RunAs" }
                    Start-Process @processParams -Wait -PassThru
                    Show-NonBlockingMessage -message "Key successfully deleted via reg.exe." -title "Success" -timeout 2
                    Update-RegistryListView -refresh $true
                } catch { [System.Windows.Forms.MessageBox]::Show("Failed to delete the key with reg.exe: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
            }
        })    
        $contextMenu.Items.Add($deleteKeyMenuItem) | Out-Null
    }
    $listView.ContextMenuStrip = $contextMenu
}

$launch_progressBar.Value = 85


function Show-NonBlockingMessage {
    param([string]$message, [string]$title = "Information", [int]$timeout = 0)
    Add-Type -AssemblyName System.Windows.Forms ; Add-Type -AssemblyName System.Drawing
    $successform = New-Object System.Windows.Forms.Form ; $successform.Text = $title ; $successform.Size = [Drawing.Size]::new(300, 150) ; $successform.StartPosition="CenterScreen" ; $successform.TopMost=$true
    $label = New-Object System.Windows.Forms.Label ; $label.Text=$message ; $label.AutoSize=$true ; $label.Location=[Drawing.Point]::new(20,20) ; $null=$successform.Controls.Add($label)
    $button = New-Object System.Windows.Forms.Button ; $button.Text="OK" ; $button.Size=[Drawing.Size]::new(75,30) ; $button.Location=[Drawing.Point]::new(110,70) ; $button.Add_Click({ param($csender,$cargs) ($csender.FindForm()).Close() }) ; $null=$successform.Controls.Add($button)
    if ($timeout -gt 0) {
        $timer=New-Object System.Windows.Forms.Timer ; $timer.Interval=$timeout*1000 ; $timer.Tag=$successform
        $timer.Add_Tick({ param($tsender,$targs) $tsender.Stop() ; ($tsender.Tag).Close() })
        $timer.Start()
    }
    $successform.Show()
}


function ListExport {
    param(
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.ListView]$listView,
        [Parameter(Mandatory=$true)]
        [bool]$all
    )
    $itemsToExport = if ($all) { $listView.Items } else { $listView.SelectedItems }
    if ($itemsToExport.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Nothing to export.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    $formatForm = New-Object System.Windows.Forms.Form
    $formatForm.Text = "Export" ; $formatForm.Width = 300 ; $formatForm.Height = 150
    $formatForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $formatForm.MinimizeBox = $false ; $formatForm.MaximizeBox = $false ; $formatForm.ShowIcon = $false ; $formatForm.StartPosition = "CenterScreen"
    $label = New-Object System.Windows.Forms.Label ; $label.Text = "Choose an export format :" ; $label.AutoSize = $true ; $label.Top = 10 ; $label.Left = 10
    $formatForm.Controls.Add($label)
    $comboBox = New-Object System.Windows.Forms.ComboBox ; $comboBox.Top = 30 ; $comboBox.Left = 10 ; $comboBox.Width = 100
    $comboBox.Items.Add("OGV Hybrid") ; $comboBox.Items.Add("XLSX") ; $comboBox.Items.Add("CSV") ; 
    $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList ; $comboBox.SelectedIndex = 0
    $formatForm.Controls.Add($comboBox)
    $openafterbtn = New-Object System.Windows.Forms.Checkbox ; $openafterbtn.Text = "Open after export" ; $openafterbtn.AutoSize = $true ; $openafterbtn.Top = 30 ; $openafterbtn.Left = 120
    $formatForm.Controls.Add($openafterbtn)
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Export" ; $okButton.Top = 55 ; $okButton.Left = 10 ; $okButton.Enabled = $true
    $formatForm.Controls.Add($okButton)
    $cancelButton = New-Object System.Windows.Forms.Button ; $cancelButton.Text = "Cancel" ; $cancelButton.Top = 55 ; $cancelButton.Left = 85 ; $cancelButton.Add_Click({ $formatForm.Close() })
    $formatForm.Controls.Add($cancelButton)
    $ExProgress = New-Object System.Windows.Forms.ProgressBar
    $ExProgress.Top = 80 ; $ExProgress.Left = 10 ; $ExProgress.Width = 260 ; $ExProgress.Style = 'Continuous' ; $ExProgress.Minimum = 0 ; $ExProgress.Maximum = 100 ; $ExProgress.Value = 0 ; $ExProgress.Visible = $false
    $formatForm.Controls.Add($ExProgress)
    $okButton.Add_Click({
        $okButton.Enabled = $false
        $comboBox.Enabled = $false
        $columns = $listView.Columns
        $data = New-Object System.Collections.Generic.List[System.Object[]]
        foreach ($item in $itemsToExport) {
            $row = @($null) * $columns.Count
            foreach ($col in $columns) { $index = $col.Index ; if ($item.SubItems.Count -gt $index) { $row[$index] = $item.SubItems[$index].Text } else { $row[$index] = "" } }
            $data.Add($row)
        }
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $dateTimeString = (Get-Date -Format "yyyy-MM-dd_HH-mm")
        switch ($comboBox.SelectedItem) {
            "OGV Hybrid"  { $saveFileDialog.Filter = "Batch (*.bat)|*.bat"; $saveFileDialog.FileName = "Export_MSI_$dateTimeString.bat" }
            "XLSX" { $saveFileDialog.Filter = "Excel (*.xlsx)|*.xlsx"; $saveFileDialog.FileName = "Export_MSI_$dateTimeString.xlsx" }
            "CSV"  { $saveFileDialog.Filter = "CSV (*.csv)|*.csv"; $saveFileDialog.FileName = "Export_MSI_$dateTimeString.csv" }
        }
        $saveFileDialog.Title = "Select a path for export"
        $result = $saveFileDialog.ShowDialog()        
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { $okButton.Enabled=$true ; $comboBox.Enabled=$true ; return }
        $filePath = $saveFileDialog.FileName
        $ExProgress.Visible = $true
        [System.Windows.Forms.Application]::DoEvents()
        $totalSteps = $data.Count
        $progressStep = [int](100 / $totalSteps)
        $currentProgress = 0
        switch ($comboBox.SelectedItem) {
            "OGV Hybrid" {
                $headers = $columns | ForEach-Object { $_.Text }
                $objects = New-Object System.Collections.ArrayList
                foreach ($row in $data) {
                    $obj = [ordered]@{}
                    for ($i = 0; $i -lt $headers.Count; $i++) { $obj[$headers[$i]] = $row[$i] }
                    $objects.Add([PSCustomObject]$obj) | Out-Null
                    $currentProgress += $progressStep
                    $ExProgress.Value = [int][math]::Min($currentProgress, 100)
                    [System.Windows.Forms.Application]::DoEvents()
                }
                $scriptContent = @"
<# ::
    cls & @echo off & title Export_MSI_$dateTimeString
    copy /y "%~f0" "%TEMP%\%~n0.ps1" >NUL && powershell -Nologo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%TEMP%\%~n0.ps1"
    exit /b
#>

`$jsonData = @'
$(($objects | ConvertTo-Json))
'@

`$data = ConvertFrom-Json -InputObject `$jsonData
`$data | Select-Object '$($headers -join "','")' | Out-GridView -Title 'Export_MSI_$dateTimeString' -Wait
"@
                [System.IO.File]::WriteAllLines($filePath, $scriptContent, [System.Text.Encoding]::UTF8)
            }
            "XLSX" {
                $headers = $columns | ForEach-Object { $_.Text }
                $dataArray = $data.ToArray()
                Export-ToXlsx -Path $filePath -Data $dataArray -Columns $headers -ProgressBar $ExProgress
            }
            "CSV" {
                $headers = ($columns | ForEach-Object { $_.Text }) -join ";"
                $csvLines = New-Object System.Collections.Generic.List[string]
                $csvLines.Add($headers)
                foreach ($row in $data) { $csvLines.Add(($row -join ";")) ; $currentProgress+=$progressStep ; $ExProgress.Value=[int][math]::Min($currentProgress, 100) ; [System.Windows.Forms.Application]::DoEvents() }
                [System.IO.File]::WriteAllLines($filePath, $csvLines)
            }
        }
        $ExProgress.Value = 100
        [System.Windows.Forms.Application]::DoEvents()
        if ($openafterbtn.Checked) { Start-Process -FilePath $filePath }
        $formatForm.Close()
        Show-NonBlockingMessage -message "Export Complete" -title "Success" -timeout 2
    })
    $formatForm.ShowDialog()
}


function Export-ToXlsx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][Object[]]$Data,
        [Parameter(Mandatory=$false)][string[]]$Columns,
        [Parameter(Mandatory=$false)]$ProgressBar
    )
    [void][Reflection.Assembly]::LoadWithPartialName("WindowsBase")
    if (Test-Path $Path) {Remove-Item $Path}
    if (-not $Columns -and $Data.Count -gt 0) { 
        $Columns = if ($Data[0] -is [System.Management.Automation.PSObject]) {($Data[0]|Get-Member -MemberType NoteProperty,Property|Select-Object -ExpandProperty Name)} else {"Value"} 
    }
    $Package=[System.IO.Packaging.Package]::Open($Path,[System.IO.FileMode]::Create,[System.IO.FileAccess]::ReadWrite)
    $Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $RelNamespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    $AddCell = {
        param($Doc,$ParentNode,$Namespace,$ColIndex,$RowIndex,$Value)
        if ($null -eq $Value) {$Value = ""}
        $Cell=$Doc.CreateElement("c",$Namespace)
        $Cell.SetAttribute("r",([char](65+$ColIndex-1))+[string]$RowIndex)
        $Cell.SetAttribute("t","inlineStr")
        $InlineStr=$Doc.CreateElement("is",$Namespace)
        $TextNode=$Doc.CreateElement("t",$Namespace)
        $TextNode.InnerText=$Value
        $InlineStr.AppendChild($TextNode)|Out-Null
        $Cell.AppendChild($InlineStr)|Out-Null
        $ParentNode.AppendChild($Cell)|Out-Null
    }
    $WorkbookUri=[Uri]"/xl/workbook.xml"
    $WorkbookPart=$Package.CreatePart($WorkbookUri,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    $WorkbookXml=[xml]"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><workbook xmlns='$Namespace' xmlns:r='$RelNamespace'><sheets><sheet name='Sheet1' sheetId='1' r:id='rId1'/></sheets></workbook>"
    $WorkbookXml.Save($WorkbookPart.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write))
    $Package.CreateRelationship($WorkbookUri,[System.IO.Packaging.TargetMode]::Internal,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument","rId1")|Out-Null
    $SheetUri=[Uri]"/xl/worksheets/sheet1.xml"
    $SheetPart=$Package.CreatePart($SheetUri,"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    $SheetXml=[xml]"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><worksheet xmlns='$Namespace' xmlns:r='$RelNamespace'><sheetData/></worksheet>"
    $NamespaceManager=New-Object System.Xml.XmlNamespaceManager($SheetXml.NameTable)
    $NamespaceManager.AddNamespace("x",$Namespace)
    $SheetData=$SheetXml.SelectSingleNode("//x:sheetData",$NamespaceManager)
    $RowIndex=1
    $HeaderRow=$SheetXml.CreateElement("row",$Namespace)
    $HeaderRow.SetAttribute("r",[string]$RowIndex)
    $SheetData.AppendChild($HeaderRow)|Out-Null
    $ColIndex=1
    $Columns|ForEach-Object{&$AddCell $SheetXml $HeaderRow $Namespace $ColIndex $RowIndex $_;$ColIndex++}
    $RowIndex++
    $progressIncrement=100/$Data.Count
    $currentProgress=0
    foreach($RowData in $Data){
        $Row=$SheetXml.CreateElement("row",$Namespace)
        $Row.SetAttribute("r",[string]$RowIndex)
        $SheetData.AppendChild($Row)|Out-Null
        for($ColIndex=1;$ColIndex -le $Columns.Count;$ColIndex++){
            $Header=$Columns[$ColIndex-1]
            $CellValue=if($RowData -is [System.Management.Automation.PSObject]){($RowData|Select-Object -ExpandProperty $Header -ErrorAction SilentlyContinue)}else{$RowData[$ColIndex-1]}
            &$AddCell $SheetXml $Row $Namespace $ColIndex $RowIndex $CellValue
        }
        $RowIndex++
        if ($ProgressBar) {
            $currentProgress+=$progressIncrement
            $ProgressBar.Value=[int][math]::Min($currentProgress,100)
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    $SheetXml.Save($SheetPart.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write))
    ($Package.GetPart($WorkbookUri)).CreateRelationship($SheetUri,[System.IO.Packaging.TargetMode]::Internal,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet","rId1")|Out-Null
    $Package.Close()
}


$launch_progressBar.Value = 90


ConfigureListViewContextMenu -listView $listView_Explore
ConfigureListViewContextMenu -listView $listView_Registry


$form.add_OnWindowMessage({
    param($s, $m)
    $updateColumns = {
        param($lv)
        $currentWidth = $lv.ClientSize.Width
        if (-not $script:lastWidth[$lv] -or $script:lastWidth[$lv] -ne $currentWidth) {
            $script:lastWidth[$lv] = $currentWidth
            AdjustListViewColumns -listView $lv
        }
    }
    $processAction = {
        foreach ($lv in @($listView_Explore, $listView_Registry)) { if ($lv.IsHandleCreated) { & $updateColumns $lv } }
    }
    switch ($m.Msg) {
        0x231 { $script:resizePending = $true }
        0x232 { if ($script:resizePending) { $script:resizePending = $false; & $processAction } }
        0x0005 { if (-not $script:resizePending) { & $processAction } }
    }
})


$tabControl.Add_SelectedIndexChanged({
    if ($tabControl.SelectedTab -eq $tabPage3 -and -not $script:tab3Refreshed) {
		$script:tab3Refreshed = $true
        $script:currentTabIndex = 2
        $selectedPaths = New-Object System.Collections.ArrayList
        if ($checkbox_HKCU64.Checked) { $selectedPaths.Add("HKLM:$($checkbox_HKCU64.Text.Substring(5))") }
        if ($checkbox_HKCU32.Checked) { $selectedPaths.Add("HKLM:$($checkbox_HKCU32.Text.Substring(5))") }
        if ($checkbox_HKLM64.Checked) { $selectedPaths.Add("HKCU:$($checkbox_HKLM64.Text.Substring(5))") }
        if ($checkbox_HKLM32.Checked) { $selectedPaths.Add("HKCU:$($checkbox_HKLM32.Text.Substring(5))") }
        if ($selectedPaths.Count -gt 0) {
			PopulateRegistryListView -listViewparam $listView_Registry -registryPaths $selectedPaths -allListViewItems $allListViewItems_Registry
            for ($i=0 ; $i -lt 2 ; $i++) {
                $row = $listView_Registry.Items[$i]
                $rowContent = @()
                foreach ($subItem in $row.SubItems) { $rowContent += $subItem.Text }
            }
			return
        }
    }
    if ($tabControl.SelectedTab -eq $tabPage1) { Update-ButtonPositions ; $script:currentTabIndex = 0 }
    if ($tabControl.SelectedTab -eq $tabPage2) { AdjustListViewColumns -listView $listView_Explore ; $script:currentTabIndex = 1 }
    if ($tabControl.SelectedTab -eq $tabPage3) { AdjustListViewColumns -listView $listView_Registry ; $script:currentTabIndex = 2 }
    if ($tabControl.SelectedTab -eq $tabPageTheme) { 
        if ($script:DarkMode -eq 0) {$script:DarkMode = 1} else {$script:DarkMode = 0}
        $tabControl.SelectedIndex = $script:currentTabIndex
    }
})


$form.Add_Load({
    $launch_progressBar.Value = 95
    $loadingLabel.Text       = "Finalizing..."
    $form.MinimumSize        = New-Object System.Drawing.Size(600, 400)
})


$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen


$form.Add_Shown({
	$tabControl.Size = $form.ClientSize
	Update-ButtonPositions
    $launch_progressBar.Value = 100
    $loadingLabel.Text       = "Complete"
    $loadingForm.Close()
    
    $form.Activate()
})


[System.Windows.Forms.Application]::Run($form)
