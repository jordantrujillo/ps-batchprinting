#Set time between each process launch
$seconds = 3

#Initial Default printer to revert back to
$InitPrinter = Get-WmiObject -Query " SELECT * FROM Win32_Printer WHERE Default=$true"

#List of processes to check
$checkproccess = 'notepad,winword,excel,powerpnt,AcroRd32,Acrobat,chrome,firefox,msedge' -split ','

#Loop for checking open processes and helping to close them
do {
    $array = @()
    foreach ($item in $checkproccess) {
        $proc = Get-Process | Where-Object ProcessName -Match "$($item)" | Select -Unique
        $array += New-Object psobject -Property @{ 'Description' = $proc }
}
if ($array.Description.Description -eq $null) {
  break
} 
if ($array.Description.Description -ne $null) {
    $wshell = New-Object -ComObject Wscript.Shell
    $Output = $wshell.Popup("The following application(s) must be closed before the program can proceed:`n`n" + ($array.Description.Description -join "`n") + "`n`nPlease close the application(s) and press 'Retry' to continue with the program.",0,"PS-BatchPrinting: Problem",5+48)
} 
if ($Output -eq 4) {
  #Close Programs Gracefully
  foreach ($item in $checkproccess) {
  Get-Process | Where-Object ProcessName -Match "$($item)" | % { $_.CloseMainWindow() | Out-Null } | stop-process -force
}
}
} until (($array.Description.Description -eq $null) -and ($Output -eq 4) -or ($Output -eq 2))

if ($Output -eq 2) {
    Exit
}

#Select Printer as Default
$PrinterList = @()

$GetPrinters = Get-WMIObject Win32_Printer -ComputerName $env:COMPUTERNAME

foreach ($Printer in $GetPrinters) {
    $PrinterList += New-Object psobject -Property @{'PrinterName' = $Printer.Name}
}

#UI (Select Printer Default)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a Printer'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select a Printer:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

foreach ($object in $PrinterList) {
    [void] $listBox.Items.Add($object.PrinterName)
}

$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $listBox.SelectedItem
    $x
} else {
  Exit
}

#Set Default Printer
(New-Object -ComObject WScript.Network).SetDefaultPrinter($x)

#Prompt for Multi-File selection
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    Multiselect = $true # Multiple files can be chosen
    Filter = 'Documents|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.csv;*.txt' # Specified file types
}
 
[void]$FileBrowser.ShowDialog()

#Check for Input, exit if null
if ($FileBrowser.FileNames) { $FileBrowser.FileNames } 
    elseif (!$FileBrowser.FileNames) { 
      $wshell = New-Object -ComObject Wscript.Shell
      $Output = $wshell.Popup("You did not select any files, application will now close.",0,"PS-BatchPrinting: Problem",0+64)
      #Revert Default Printer Change
      (New-Object -ComObject WScript.Network).SetDefaultPrinter($InitPrinter.Name)
        Exit 
    }
  
$path = $FileBrowser.FileNames

#Parent Directory
$directory = $FileBrowser.FileName | Get-ChildItem | Split-Path -Parent

#Log
Start-Transcript -Path "$($directory)\PS-BatchPrinting-Log.txt" -Append

#Time Stamp
$timestamp = get-date
Write-Host "##############################################################"
Write-Host "$($timestamp)"
Write-Host "##############################################################"

#Get Default Printer
$DefaultPrinter = Get-WmiObject -Query " SELECT * FROM Win32_Printer WHERE Default=$true"

#Open Printer Preferences
rundll32.exe printui.dll,PrintUIEntry /e /n "$($DefaultPrinter.Name)"

Start-Sleep 1

#Warning user that batch job will go to selected printer
$wshell = New-Object -ComObject Wscript.Shell
$Output = $wshell.Popup("All selected files will be printed to $($DefaultPrinter.Name), at their default settings. Make sure your default settings are applied in the 'Printing Preferences' Window. Do you wish to Continue?",0,"PS-BatchPrinting: Warning - Default Printer",4+48)


  switch  ($Output) {

  '6' {

    Write-Host "All Files will be Printed to $($DefaultPrinter.Name)"
    If($FileBrowser.FileNames -like "*\*") {

	  #Printing Selected Documents in Directory
	  foreach($file in Get-ChildItem $path){
	  Get-ChildItem ($file) |
		  ForEach-Object {
            Start-Process -FilePath $file.Fullname -Verb Print -PassThru | Wait-Process -Name notepad, winword, excel, powerpnt, AcroRd32, Acrobat, chrome, firefox, msedge, opera  -Timeout 60 -ErrorAction SilentlyContinue | Stop-Process
            Start-Sleep $seconds
            Write-Host "Printing $($file.name)"
	  	}
  }
  Start-Sleep 10
  }
  }
    '7' {
    Stop-Transcript
    Exit
  }
  }



#Change Directory so that the powershell session is in the dirctory specified by user
Set-Location $directory

#Open Print Queue
rundll32.exe printui.dll,PrintUIEntry /o /n "$($DefaultPrinter.Name)"

#Prompt Move files to Printed Folder
$wshell = New-Object -ComObject Wscript.Shell
$Output = $wshell.Popup("Please VERIFY that the files have printed. Would you like to move the files to a 'Printed Archive' Folder, located in the current directory?",0,"PS-BatchPrinting: Complete!",4+32)
 
  switch  ($Output) {

  '6' {

    #Creates Printed Folder
    $printed = "./Printed Archive"
    $testedpath = $printed
    If(!(test-path $testedpath))
    {
      New-Item -ItemType Directory -Force -Path $testedpath
    }
    Write-Host "##############################################################"
    Write-Host "Moving Files..."
    Write-Host "##############################################################"
    Set-Location $printed
    $pwd = Get-Location
    foreach ($file in Get-childitem $path){
        Move-Item -Path $file.fullname -Destination $pwd
        Write-Host "$($file.name) moved to $($pwd.Path)"
        }
    
        #Open File Explorer to Moved Items
        Invoke-Item .

        #Revert Default Printer Change
        (New-Object -ComObject WScript.Network).SetDefaultPrinter($InitPrinter.Name)

          #Close Programs Gracefully
          foreach ($item in $checkproccess) {
          Get-Process | Where-Object ProcessName -Match "$($item)" | % { $_.CloseMainWindow() | Out-Null } | stop-process -force
          }
        
        #Balloon Notification
        Add-Type -AssemblyName System.Windows.Forms
        $global:balmsg = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
        $balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $balmsg.BalloonTipText = 'All done! Thank you for using PS-BatchPrinting.'
        $balmsg.BalloonTipTitle = "PS-BatchPrinting"
        $balmsg.Visible = $true
        $balmsg.ShowBalloonTip(20000)
      }
  '7' {
    #Revert Default Printer Change
    (New-Object -ComObject WScript.Network).SetDefaultPrinter($InitPrinter.Name)

    #Close Programs Gracefully
    foreach ($item in $checkproccess) {
    Get-Process | Where-Object ProcessName -Match "$($item)" | % { $_.CloseMainWindow() | Out-Null } | stop-process -force
    }

    #Balloon Notification
    Add-Type -AssemblyName System.Windows.Forms
    $global:balmsg = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $balmsg.BalloonTipText = 'All done! Thank you for using PS-BatchPrinting.'
    $balmsg.BalloonTipTitle = "PS-BatchPrinting"
    $balmsg.Visible = $true
    $balmsg.ShowBalloonTip(20000)
  }

  }

Stop-Transcript
Exit
