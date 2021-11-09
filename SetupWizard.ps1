#====================================================#
#############    Hide Console Window     #############
#====================================================#
Import-Module ScheduledTasks
Import-Module PowerShellGet
Import-Module SecureBoot
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)


Function Show-SetupWizard {
	
	#====================================================#
	#############    Imported Assemblies    ##############
	#====================================================#
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
	[void][reflection.assembly]::Load('System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	
	
	#====================================================#
	##############  Create Form Objects    ###############
	#====================================================#
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form2 = New-Object 'System.Windows.Forms.Form'
	$panel2 = New-Object 'System.Windows.Forms.Panel'
	$textbox1 = New-Object 'System.Windows.Forms.TextBox'
	$panel1 = New-Object 'System.Windows.Forms.Panel'
	$cbPolicy = New-Object 'System.Windows.Forms.CheckBox'
	$cbPsVersion = New-Object 'System.Windows.Forms.CheckBox'
	$cbGoogle = New-Object 'System.Windows.Forms.CheckBox'
	$cbGitHub = New-Object 'System.Windows.Forms.CheckBox'
	$progressbar1 = New-Object 'System.Windows.Forms.ProgressBar'
	$cbExcel = New-Object 'System.Windows.Forms.CheckBox'
	$cbDatabase = New-Object 'System.Windows.Forms.CheckBox'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	
	
	
	#====================================================#
	#############    Form Loading Events    ##############
	#====================================================#
	
	$form2_Load = {
		$okCancel = [System.Windows.Forms.MessageBox]::Show("This wizard will install all the necassary modules and updates required for the parts database. `Click OK to continue or Cancel to exit the setup", "System Update", "OKCancel", "Information")
		If ($okCancel -eq [System.Windows.Forms.DialogResult]::OK) {
			$form2.Visible = $true
			$form2.TopLevel = $true
			$form2.TopMost = $true
			$textbox1.Text = "Setting Execution Policy for PowerShell..."
			$progressbar1.pers
			Start-Sleep -Seconds 2
			(Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -Force)
			$cbPolicy.Checked = $True
			$textbox1.Text = "Local machine's PS Execution Policy Change -- Complete!"
			Start-Sleep -Seconds '3'
			$g = (Get-Host).Version
			If ($g.Major -le 5) {
				$textbox1.Text = "Updating PowerShell to version 7.1.5"
				(Invoke-Expression "& { $(Invoke-RestMethod https://aka.ms/install-powershell.ps1) } -UseMSI")
			}
			$cbPsVersion.Checked = $true
			$textbox1.Text = "Install Powershell 7.1.5 -- Complete!"
			Start-Sleep -Seconds '3'
			$githubCheck = (Get-Module -ListAvailable -Name 'PowerShellForGitHub')
			If ($null -eq $githubCheck) {
				$textbox1.Text = "Installing GitHub API Module..."
				(Install-Module 'PowerShellForGitHub' -Scope CurrentUser -AllowClobber -Force)
			}
			$cbGitHub.Checked = $true
			$textbox1.Text = "Install GitHub Rest API Module -- Complete!"
			Start-Sleep -Seconds '3'
			$googleCheck = (Get-Module -ListAvailable -Name 'UMN-Google')
			If ($null -eq $googleCheck) {
				$textbox1.Text = "Installing Google API Module..."
				(Install-Module 'UMN-Google' -Scope CurrentUser -AllowClobber -Force)
			}
			$cbGoogle.Checked = $true
			$textbox1.Text = "Install Google Rest API Module -- Complete!"
			Start-Sleep -Seconds '3'
			$excelCheck = (Get-Module -ListAvailable -Name 'ImportExcel')
			If ($null -eq $excelCheck) {
				$textbox1.Text = "Installing MS Excel Module..."
				(Install-Module 'ImportExcel' -Scope CurrentUser -AllowClobber -Force)
			}
			$cbExcel.Checked = $true
			$textbox1.Text = "Install Microsoft Excel .NET Module -- Complete!"
			Start-Sleep -Seconds '3'
			$dpcPath = Join-Path "$($env:APPDATA)" "DPC_PartsDB"
			$dpcPathExists = Test-Path $dpcPath
			$textbox1.Text = "Installing Database Components..."
			If ($dpcPathExists -eq $false) {
				(New-Item -Path "$($env:APPDATA)" -Name "DPC_PartsDB" -ItemType "Directory")
			}
			$changeDir = Join-Path "$($env:APPDATA)" "DPC_PartsDB/Change_Log"
			$changeExists = Test-Path $changeDir
			If ($changeExists -eq $false) {
				(New-Item -Path "$($env:APPDATA)/DPC_PartsDB" -Name "Change_Log" -ItemType "Directory")
			}
			$textbox1.Text = "Create Database directories in local File System -- Complete!"
			Start-Sleep -Seconds '3'
			$icoPath = "$($env:APPDATA)/DPC_PartsDB/DataTable.ico"
			$icoExists = Test-Path $icoPath
			If ($icoExists -ne $True) {
				(Invoke-WebRequest -Uri https://raw.githubusercontent.com/Mediocre-Software/DPC_PartsDB/main/DataTable.ico -OutFile "$($env:APPDATA)/DPC_PartsDB/DataTable.ico")
			}
			$logoPath = "$($env:APPDATA)/DPC_PartsDB/DPCLogo.png"
			$logoExists = Test-Path $logoPath
			If ($logoExists -ne $True) {
				(Invoke-WebRequest -Uri https://raw.githubusercontent.com/Mediocre-Software/DPC_PartsDB/main/DPCLogo.png -OutFile "$($env:APPDATA)/DPC_PartsDB/DPCLogo.png")
			}
			$firstRun = "$($env:APPDATA)/DPC_PartsDB/FirstRun.Check"
			$firstRunExists = Test-Path $firstRun
			If ($firstRunExists -ne $true) {
				(Invoke-WebRequest -Uri https://raw.githubusercontent.com/Mediocre-Software/DPC_PartsDB/main/FirstRun.Check -OutFile "$($env:APPDATA)/DPC_PartsDB/FirstRun.Check")
			}
			$licPath = "$($env:APPDATA)/DPC_PartsDB/GPL3.0.txt"
			$licExists = Test-Path $licPath
			If ($licExists -ne $true) {
				(Invoke-WebRequest -Uri https://raw.githubusercontent.com/Mediocre-Software/DPC_PartsDB/main/GPL3.0.txt -OutFile "$($env:APPDATA)/DPC_PartsDB/GPL3.0.txt")
			}
			$aboutPath = "$($env:APPDATA)/DPC_PartsDB/About.ps1"
			$aboutExists = Test-Path $aboutPath
			If ($aboutExists -ne $true) {
				(Invoke-WebRequest -Uri https://raw.githubusercontent.com/Mediocre-Software/DPC_PartsDB/main/About.ps1 -OutFile "$($env:APPDATA)/DPC_PartsDB/About.ps1")
			}
			$cbDatabase.Checked = $true
			$textbox1.Text = "Download and install database dependencies -- Complete!"
			Start-Sleep -Seconds '3'
		} Else {
			$form2.Close
		}
		
		$okExit = [System.Windows.Forms.MessageBox]::Show("All updates have been successfully completed. Press OK to exit this wizard and launch the Parts Database.", "Updates Complete", 'OK')
		If ($okExit -eq [System.Windows.Forms.DialogResult]::OK) {
			$scriptPath = $MyInvocation.PSScriptRoot
			$partsDB = (Join-Path $scriptPath '/DPC_PartsDB.ps1')
			(Invoke-Item -Path $partsDB) | Out-Null
			$form2.Close()
		}
	}
	
	$Form_StateCorrection_Load =
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form2.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed =
	{
		#Remove all event handlers from the controls
		Try {
			$form2.remove_Load($form2_Load)
			$form2.remove_Load($Form_StateCorrection_Load)
			$form2.remove_FormClosed($Form_Cleanup_FormClosed)
		} Catch {
			Out-Null <# Prevent PSScriptAnalyzer warning #>
		}
	}
	
	#endregion Generated Events
	
	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$form2.SuspendLayout()
	$panel1.SuspendLayout()
	$panel2.SuspendLayout()
	#
	# formDPCSetupWizard
	#
	$form2.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$($env:APPDATA)/DPC_PartsDB/DataTable.ico")
	$form2.Controls.Add($panel2)
	$form2.Controls.Add($panel1)
	$form2.AutoScaleDimensions = '9, 19'
	$form2.AutoScaleMode = 'Font'
	$form2.BackColor = 'ControlLight'
	$form2.ClientSize = '420,150'
	$form2.Cursor = 'Default'
	$form2.Font = 'Consolas, 9pt'
	$form2.FormBorderStyle = 'FixedSingle'
	$form2.Margin = '3, 3, 3, 3'
	$form2.MaximizeBox = $False
	$form2.MinimizeBox = $true
	$form2.WindowState = 'Normal'
	$form2.TopMost = $true
	$form2.Name = 'form2'
	$form2.ShowIcon = $True
	$form2.ShowInTaskbar = $True
	$form2.SizeGripStyle = 'Hide'
	$form2.StartPosition = 'CenterScreen'
	$form2.Text = 'DiscountPC DB-Setup Wizard'
	$form2.Enabled = $true
	$form2.add_Load($form2_Load)
	#
	# panel2
	#
	$panel2.Controls.Add($textbox1)
	$panel2.BorderStyle = 'FixedSingle'
	$panel2.Location = '9, 1'
	$panel2.Margin = '4, 4, 4, 4'
	$panel2.Name = 'panel2'
	$panel2.Size = '400, 50'
	$panel2.TabIndex = 3
	#
	# textbox1
	#
	$textbox1.BackColor = 'ControlLight'
	$textbox1.BorderStyle = 'None'
	$textbox1.Cursor = 'Default'
	$textbox1.Font = 'Lucida Console, 8pt'
	$textbox1.Location = '13, 12'
	$textbox1.Margin = '4, 4, 4, 4'
	$textbox1.Multiline = $True
	$textbox1.Name = 'textbox1'
	$textbox1.ReadOnly = $True
	$textbox1.ShortcutsEnabled = $False
	$textbox1.Size = '377, 28'
	$textbox1.TabIndex = 4
	$textbox1.Text = "System check completed -- Updates Required. Awaiting user-input."
	$textbox1.TextAlign = 'Center'
	#
	# panel1
	#
	$panel1.Controls.Add($cbPolicy)
	$panel1.Controls.Add($cbPsVersion)
	$panel1.Controls.Add($cbGoogle)
	$panel1.Controls.Add($cbGitHub)
	$panel1.Controls.Add($progressbar1)
	$panel1.Controls.Add($cbExcel)
	$panel1.Controls.Add($cbDatabase)
	$panel1.BorderStyle = 'Fixed3D'
	$panel1.Location = '9, 56'
	$panel1.Margin = '4, 4, 4, 4'
	$panel1.Name = 'panel1'
	$panel1.Size = '400, 100'
	$panel1.TabIndex = 12
	#
	# checkboxPowerShellSetExecuti
	#
	$cbPolicy.AutoCheck = $false
	$cbPolicy.AutoSize = $True
	$cbPolicy.Checked = $false
	$cbPolicy.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbPolicy.FlatStyle = 'System'
	$cbPolicy.Font = 'Lucida Console, 7pt'
	$cbPolicy.Location = '217, 13'
	$cbPolicy.Margin = '4, 4, 4, 4'
	$cbPolicy.Name = 'cbPolicy'
	$cbPolicy.Size = '184,16'
	$cbPolicy.TabIndex = 5
	$cbPolicy.Text = 'Local Execution Policy'
	$cbPolicy.UseCompatibleTextRendering = $True
	$cbPolicy.UseVisualStyleBackColor = $True
	#
	# checkboxPowerShellVersion7Or
	#
	$cbPsVersion.AutoCheck = $False
	$cbPsVersion.AutoSize = $True
	$cbPsVersion.Checked = $False
	$cbPsVersion.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbPsVersion.FlatStyle = 'System'
	$cbPsVersion.Font = 'Lucida Console, 7pt'
	$cbPsVersion.ImageAlign = 'TopRight'
	$cbPsVersion.Location = '217, 35'
	$cbPsVersion.Margin = '4, 4, 4, 4'
	$cbPsVersion.Name = 'cbPsVersion'
	$cbPsVersion.Size = '184, 16'
	$cbPsVersion.TabIndex = 6
	$cbPsVersion.Text = 'PowerShell 7.1.5'
	$cbPsVersion.UseCompatibleTextRendering = $True
	$cbPsVersion.UseVisualStyleBackColor = $True
	#
	# checkboxGoogleRESTAPIModule
	#
	$cbGoogle.AutoCheck = $false
	$cbGoogle.AutoSize = $True
	$cbGoogle.Checked = $False
	$cbGoogle.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbGoogle.FlatStyle = 'System'
	$cbGoogle.Font = 'Lucida Console, 7pt'
	$cbGoogle.Location = '17, 13'
	$cbGoogle.Margin = '4, 4, 4, 4'
	$cbGoogle.Name = 'cbGoogle'
	$cbGoogle.Size = '184, 16'
	$cbGoogle.TabIndex = 7
	$cbGoogle.Text = 'Google API Module'
	$cbGoogle.UseCompatibleTextRendering = $True
	$cbGoogle.UseVisualStyleBackColor = $True
	#
	# checkboxGitHubRESTAPIModule
	#
	$cbGitHub.AutoCheck = $false
	$cbGitHub.AutoSize = $True
	$cbGitHub.Checked = $False
	$cbGitHub.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbGitHub.FlatStyle = 'System'
	$cbGitHub.Font = 'Lucida Console, 7pt'
	$cbGitHub.Location = '17, 35'
	$cbGitHub.Margin = '4, 4, 4, 4'
	$cbGitHub.Name = 'cbGitHub'
	$cbGitHub.Size = '184, 16'
	$cbGitHub.TabIndex = 8
	$cbGitHub.Text = 'GitHub API Module'
	$cbGitHub.UseCompatibleTextRendering = $True
	$cbGitHub.UseVisualStyleBackColor = $True
	#
	# checkboxMicrosoftExcelNETPac
	#
	$cbExcel.AutoCheck = $false
	$cbExcel.AutoSize = $True
	$cbExcel.Checked = $False
	$cbExcel.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbExcel.FlatStyle = 'System'
	$cbExcel.Font = 'Lucida Console, 7pt'
	$cbExcel.Location = '17, 59'
	$cbExcel.Margin = '4, 4, 4, 4'
	$cbExcel.Name = 'cbExcel'
	$cbExcel.Size = '184, 16'
	$cbExcel.TabIndex = 9
	$cbExcel.Text = 'MS Excel Module'
	$cbExcel.UseCompatibleTextRendering = $True
	$cbExcel.UseVisualStyleBackColor = $True
	#
	# checkboxMicrosoftExcelNETPac
	# 
	$cbDatabase.AutoCheck = $false
	$cbDatabase.AutoSize = $True
	$cbDatabase.Checked = $False
	$cbDatabase.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbDatabase.FlatStyle = 'System'
	$cbDatabase.Font = 'Lucida Console, 7pt'
	$cbDatabase.Location = '217, 59'
	$cbDatabase.Margin = '4, 4, 4, 4'
	$cbDatabase.Name = 'cbDatabase'
	$cbDatabase.Size = '184, 16'
	$cbDatabase.TabIndex = 10
	$cbDatabase.Text = 'DPC Parts Database'
	$cbDatabase.UseCompatibleTextRendering = $True
	$cbDatabase.UseVisualStyleBackColor = $True
	#
	# progressbar1
	#
	$progressbar1.Dock = 'Bottom'
	$progressbar1.Margin = '4, 4, 4, 4'
	$progressbar1.Name = 'progressbar1'
	$progressbar1.Size = '401, 20'
	$progressbar1.TabIndex = 0
	$progressbar1.Style = 'Continuous'
	$progressbar1.MarqueeAnimationSpeed = '250'
	$progressbar1.Enabled = $true
	$progressbar1.Visible = $true
	
	$panel2.ResumeLayout()
	$panel1.ResumeLayout()
	$form2.ResumeLayout()
	
	#----------------------------------------------
	#----------------------------------------------
	
	#Save the initial state of the form
	$InitialFormWindowState = $form2.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form2.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$form2.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	Return $form2.ShowDialog()
	
} #End Function

#Call the form
Show-SetupWizard | Out-Null