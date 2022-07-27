Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the window form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Create Rack CSVs'
$form.Size = New-Object System.Drawing.Size(280,220)
$form.StartPosition = 'CenterScreen'

# OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(50,140)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

# Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,140)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

# Time frame label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(200,30)
$label.Text = 'Get files from the past number of days:'
$form.Controls.Add($label)

# Time frame input box
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(210,15)
$textBox.Size = New-Object System.Drawing.Size(30,20)
$textBox.TextAlign = 'Right'
$textBox.Text = 10
$form.Controls.Add($textBox)

# Source directory label
$srclabel = New-Object System.Windows.Forms.Label
$srclabel.Location = New-Object System.Drawing.Point(10,60)
$srclabel.Size = New-Object System.Drawing.Size(100,30)
$srclabel.Text = 'Source Folder:'
$form.Controls.Add($srclabel)

#Source directory input box
$srctextBox = New-Object System.Windows.Forms.TextBox
$srctextBox.Location = New-Object System.Drawing.Point(120,55)
$srctextBox.Size = New-Object System.Drawing.Size(120,20)
$srctextBox.TextAlign = 'Right'
$srctextBox.Text = 'racks-json'
$form.Controls.Add($srctextBox)

#Source directory input box
$srctextBox = New-Object System.Windows.Forms.TextBox
$srctextBox.Location = New-Object System.Drawing.Point(120,55)
$srctextBox.Size = New-Object System.Drawing.Size(120,20)
$srctextBox.TextAlign = 'Right'
$srctextBox.Text = 'racks-json'
$form.Controls.Add($srctextBox)

# Destination directory label
$destlabel = New-Object System.Windows.Forms.Label
$destlabel.Location = New-Object System.Drawing.Point(10,100)
$destlabel.Size = New-Object System.Drawing.Size(100,30)
$destlabel.Text = 'Destination Folder:'
$form.Controls.Add($destlabel)

# Destination directory input box
$desttextBox = New-Object System.Windows.Forms.TextBox
$desttextBox.Location = New-Object System.Drawing.Point(120,95)
$desttextBox.Size = New-Object System.Drawing.Size(120,20)
$desttextBox.TextAlign = 'Right'
$desttextBox.Text = 'rackcsvs'
$form.Controls.Add($desttextBox)

$form.Topmost = $true
$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $FilePaths = @()
    Get-ChildItem -Path "C:\Users\user\Desktop\DSKTP\pwrshlltst\$($srctextBox.Text)\" -Recurse -Include @("*.json")|`
    Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-$textBox.Text)} |
    foreach{
        $Path = $_.FullName
        $FilePaths += $Path
    }

    $RackFiles = @()
    Get-ChildItem -Path "C:\Users\user\Desktop\DSKTP\pwrshlltst\$($desttextBox.Text)\" -Recurse -Include @("*.csv")|`
    Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-$textBox.Text)} |
    foreach {
        $FileName = $_.BaseName
        $RackFiles += $FileName
    }

    $ExistingRacks = @()
    $NewRacks = @()
    foreach ($i in $FilePaths) 
    {
        $Obj = Get-Content -Raw -Path $i | ConvertFrom-Json
        $containers = $Obj.sampleContainers
        foreach ($cont in $containers)
        {
            $rackid = $cont.barcode
            $sampleid = $cont.samples
            $AllSamples = @()
            $sampleid | ForEach-Object {
                $AllSamples += [pscustomobject]@{
                    barcode = $_.barcode
                    limsid = $_.limsid
                }
            }
            if ( $rackid -inotin $RackFiles)
            {
                $NewRacks += $rackid
                $AllSamples | Export-Csv -Path "C:\Users\user\Desktop\DSKTP\pwrshlltst\$($desttextBox.Text)\$($rackid).csv" -NoTypeInformation
            }
            else 
            {
                $ExistingRacks += $rackid
                Write-Output "Rack $($rackid) already saved in directory."
            }
        }
    }

    if($ExistingRacks) {
        $racksnotadded = $ExistingRacks -join "`n"
    } else {
        $racksnotadded = 'N/A'
    }

    if($NewRacks) {
        $racksadded = $NewRacks -join "`n"
    } else {
        $racksadded = 'N/A'
    }

    Write-Output "The following racks already exist: $($racksnotadded)"
    [System.Windows.MessageBox]::Show("The following racks already exist:`n$($racksnotadded)`n`nThe following racks were added:`n$($racksadded)")
    Write-Output "The following racks were added: $($racksadded)"
}
