#====================================================#
#############    Hide Console Window     #############
#====================================================#
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
	$picturebox1 = New-Object 'System.Windows.Forms.PictureBox'
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
			$okUpdate = [System.Windows.Forms.MessageBox]::Show("Follow the prompts that may appear after successful downloads to complete the installation of those components.", "System Update", "OK", "Information")
			If ($okUpdate -eq [System.Windows.Forms.DialogResult]::OK) {
				$progressbar1.ForeColor = 'CornflowerBlue'
				$progressbar1.Enabled = $true
				Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -Force | Out-Null
				$cbPolicy.Checked = $True
				$textbox1.Text = "Local machine's PS Execution Policy Change -- Complete!"
				$g = (Get-Host).Version
				If ($g.Major -le 5) {
					Invoke-Exp "& { $(Invoke-RestMethod https://aka.ms/install-powershell.ps1) } -UseMSI" | Out-Null
				}
				$cbPsVersion.Checked = $true
				$textbox1.Text = "Install Powershell 7.1.5 -- Complete!"
				$githubCheck = (Get-Module -ListAvailable -Name 'PowerShellForGitHub')
				If ($null -eq $githubCheck) {
					Install-Module 'PowerShellForGitHub' -Scope CurrentUser -AllowClobber -Force | Out-Null
				}
				$cbGitHub.Checked = $true
				$textbox1.Text = "Install GitHub Rest API Module -- Complete!"
				$googleCheck = (Get-Module -ListAvailable -Name 'UMN-Google')
				If ($null -eq $googleCheck) {
					Install-Module 'UMN-Google' -Scope CurrentUser -AllowClobber -Force | Out-Null
				}
				$cbGoogle.Checked = $true
				$textbox1.Text = "Install Google Rest API Module -- Complete!"
				$excelCheck = (Get-Module -ListAvailable -Name 'ImportExcel')
				If ($null -eq $excelCheck) {
					Install-Module 'ImportExcel' -Scope CurrentUser -AllowClobber -Force | Out-Null
				}
				$cbExcel.Checked = $true
				$textbox1.Text = "Install Microsoft Excel .NET Module -- Complete!"
				$dpcPath = Join-Path "$($env:APPDATA)" "DPC_PartsDB"
				$dpcPathExists = Test-Path $dpcPath
				If ($dpcPathExists -eq $false) {
					New-Item -Path "$($env:APPDATA)" -Name "DPC_PartsDB" -ItemType "Directory" | Out-Null
				}
				$changeDir = Join-Path "$($env:APPDATA)" "DPC_PartsDB/Change_Log"
				$changeExists = Test-Path $changeDir
				If ($changeExists -eq $false) {
					New-Item -Path "$($env:APPDATA)/DPC_PartsDB" -Name "Change_Log" -ItemType "Directory" | Out-Null
				}
				$textbox1.Text = "Create Database directories in local File System -- Complete!"
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
					(Invoke-WebRequest -Uri https://raw.githubusercontent.com/Mediocre-Software/DPC_PartsDB/main/FirstRun.Check -OutFile "$($env:APPDATA)/DPC_PartsDB/FirstRun.Check")
				$cbDatabase.Checked = $true
				$textbox1.Text = "Download and install database dependencies -- Complete!"
				Start-Sleep -Seconds '2' 
				$progressbar1.Enabled = $false
				$okExit = [System.Windows.Forms.MessageBox]::Show("All updates have been successfully completed. Press OK to exit this wizard and launch the Parts Database.", "Updates Complete", 'OK')
				If ($okExit -eq [System.Windows.Forms.DialogResult]::OK) {
					$form2.Close
				}
			}
		} Else {
			$form2.Close
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
			$textbox1.remove_TextChanged($textbox1_TextChanged)
			$form2.remove_Load($formSplashScreen_Load)
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
	$form2.Controls.Add($panel2)
	$form2.Controls.Add($panel1)
	$form2.AutoScaleDimensions = '9, 19'
	$form2.AutoScaleMode = 'Font'
	$form2.BackColor = 'ControlLight'
	$form2.ClientSize = '423, 228'
	$form2.Font = 'Consolas, 9pt'
	$form2.FormBorderStyle = 'FixedSingle'
	$form2.Margin = '3, 3, 3, 3'
	$form2.MaximizeBox = $False
	$form2.Name = 'formDPCSetupWizard'
	$form2.SizeGripStyle = 'Hide'
	$form2.StartPosition = 'CenterScreen'
	$form2.Text = 'DiscountPC DB-Setup Wizard'
	$form2.TopMost = $True
	$form2.add_Load($form2_Load)
	#
	# panel2
	#
	$panel2.Controls.Add($textbox1)
	$panel2.BorderStyle = 'FixedSingle'
	$panel2.Location = '9, 5'
	$panel2.Margin = '4, 4, 4, 4'
	$panel2.Name = 'panel2'
	$panel2.Size = '403, 52'
	$panel2.TabIndex = 7
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
	$textbox1.TabIndex = 6
	$textbox1.Text = 'System check completed -- Updates Required `
Awaiting user-input.'
	$textbox1.TextAlign = 'Center'
	#
	# panel1
	#
	$panel1.Controls.Add($cbPolicy)
	$panel1.Controls.Add($cbPsVersion)
	$panel1.Controls.Add($cbGoogle)
	$panel1.Controls.Add($cbGitHub)
	$panel1.Controls.Add($picturebox1)
	$panel1.Controls.Add($progressbar1)
	$panel1.Controls.Add($cbExcel)
	$panel1.Controls.Add($cbDatabase)
	$panel1.BorderStyle = 'Fixed3D'
	$panel1.Location = '9, 66'
	$panel1.Margin = '4, 4, 4, 4'
	$panel1.Name = 'panel1'
	$panel1.Size = '405, 154'
	$panel1.TabIndex = 6
	#
	# checkboxPowerShellSetExecuti
	#
	$cbPolicy.AutoCheck = $false
	$cbPolicy.AutoSize = $True
	$cbPolicy.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbPolicy.FlatStyle = 'System'
	$cbPolicy.Font = 'Lucida Console, 7pt'
	$cbPolicy.Location = '217, 13'
	$cbPolicy.Margin = '4, 4, 4, 4'
	$cbPolicy.Name = 'cbPolicy'
	$cbPolicy.Size = '184,16'
	$cbPolicy.TabIndex = 28
	$cbPolicy.Text = 'Local Execution Policy'
	$cbPolicy.UseCompatibleTextRendering = $True
	$cbPolicy.UseVisualStyleBackColor = $True
	#
	# checkboxPowerShellVersion7Or
	#
	$cbPsVersion.AutoCheck = $false
	$cbPsVersion.AutoCheck = $False
	$cbPsVersion.AutoSize = $True
	$cbPsVersion.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbPsVersion.FlatStyle = 'System'
	$cbPsVersion.Font = 'Lucida Console, 7pt'
	$cbPsVersion.ImageAlign = 'TopRight'
	$cbPsVersion.Location = '217, 35'
	$cbPsVersion.Margin = '4, 4, 4, 4'
	$cbPsVersion.Name = 'cbPsVersion'
	$cbPsVersion.Size = '184, 16'
	$cbPsVersion.TabIndex = 27
	$cbPsVersion.Text = 'PowerShell 7.1.5'
	$cbPsVersion.UseCompatibleTextRendering = $True
	$cbPsVersion.UseVisualStyleBackColor = $True
	#
	# checkboxGoogleRESTAPIModule
	#
	$cbGoogle.AutoCheck = $false
	$cbGoogle.AutoSize = $True
	$cbGoogle.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbGoogle.FlatStyle = 'System'
	$cbGoogle.Font = 'Lucida Console, 7pt'
	$cbGoogle.Location = '17, 13'
	$cbGoogle.Margin = '4, 4, 4, 4'
	$cbGoogle.Name = 'cbGoogle'
	$cbGoogle.Size = '184, 16'
	$cbGoogle.TabIndex = 25
	$cbGoogle.Text = 'Google API Module'
	$cbGoogle.UseCompatibleTextRendering = $True
	$cbGoogle.UseVisualStyleBackColor = $True
	#
	# checkboxGitHubRESTAPIModule
	#
	$cbGitHub.AutoCheck = $false
	$cbGitHub.AutoSize = $True
	$cbGitHub.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbGitHub.FlatStyle = 'System'
	$cbGitHub.Font = 'Lucida Console, 7pt'
	$cbGitHub.Location = '17, 35'
	$cbGitHub.Margin = '4, 4, 4, 4'
	$cbGitHub.Name = 'cbGitHub'
	$cbGitHub.Size = '184, 16'
	$cbGitHub.TabIndex = 24
	$cbGitHub.Text = 'GitHub API Module'
	$cbGitHub.UseCompatibleTextRendering = $True
	$cbGitHub.UseVisualStyleBackColor = $True
	#
	# checkboxMicrosoftExcelNETPac
	#
	$cbExcel.AutoCheck = $false
	$cbExcel.AutoSize = $True
	$cbExcel.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbExcel.FlatStyle = 'System'
	$cbExcel.Font = 'Lucida Console, 7pt'
	$cbExcel.Location = '17, 59'
	$cbExcel.Margin = '4, 4, 4, 4'
	$cbExcel.Name = 'cbExcel'
	$cbExcel.Size = '184, 16'
	$cbExcel.TabIndex = 26
	$cbExcel.Text = 'MS Excel Module'
	$cbExcel.UseCompatibleTextRendering = $True
	$cbExcel.UseVisualStyleBackColor = $True
	#
	# checkboxMicrosoftExcelNETPac
	# 
	$cbDatabase.AutoCheck = $false
	$cbDatabase.AutoSize = $True
	$cbDatabase.FlatAppearance.CheckedBackColor = 'DodgerBlue'
	$cbDatabase.FlatStyle = 'System'
	$cbDatabase.Font = 'Lucida Console, 7pt'
	$cbDatabase.Location = '217, 59'
	$cbDatabase.Margin = '4, 4, 4, 4'
	$cbDatabase.Name = 'cbDatabase'
	$cbDatabase.Size = '184, 16'
	$cbDatabase.TabIndex = 26
	$cbDatabase.Text = 'DPC Parts Database'
	$cbDatabase.UseCompatibleTextRendering = $True
	$cbDatabase.UseVisualStyleBackColor = $True
	#
	# progressbar1
	#
	$progressbar1.Dock = 'Bottom'
	$progressbar1.Location = '0, 130'
	$progressbar1.Margin = '4, 4, 4, 4'
	$progressbar1.Name = 'progressbar1'
	$progressbar1.Size = '401, 20'
	$progressbar1.TabIndex = 29
	$progressbar1.ForeColor = 'CornflowerBlue'
	$progressbar1.Style = 'Marquee'
	$progressbar1.MarqueeAnimationSpeed = '85'
	$progressbar1.Minimum = '0'
	$progressbar1.Maximum = '100'
	$progressbar1.Enabled = $false
	
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
