#=============  DiscountPC Inventory   ==============#
#===============  Database System  ==================#
#====================================================#
#================  By: Nic Lynds  ===================#
#===========  MediocreSoftware (c) 2021  ============#
#====================================================#
#=================  DESCRIPTION:  ===================#
#====================================================#
#= GUI front-end, created to manage the database of =#
#=="Non-Purchased Parts" Inventory, for Discount PC==#
#====================================================#
#=========  Last Updated On: 10/25/2021  ============#
#====================================================#

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

Function Show-PartsDB
{
	
	#====================================================#
	#############    Imported Assemblies    ##############
	#====================================================#
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
	[void][reflection.assembly]::Load('System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	
	#====================================================#
	#############    Create Form Objects    ##############
	#====================================================#
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form1 = New-Object 'System.Windows.Forms.Form'
	$picturebox1 = New-Object 'System.Windows.Forms.PictureBox'
	$buttonSearch = New-Object 'System.Windows.Forms.Button'
	$datagridview1 = New-Object 'System.Windows.Forms.DataGridView'
	$textboxSearch = New-Object 'System.Windows.Forms.TextBox'
	$menustrip1 = New-Object 'System.Windows.Forms.MenuStrip'
	$fileToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$reloadToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolStripSeparator = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$saveToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$saveAsToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolStripSeparator1 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$printToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$printPreviewToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolStripSeparator2 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$exitToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$editToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$undoToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$redoToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolStripSeparator3 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$cutToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$copyToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$pasteToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolStripSeparator4 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$selectAllToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolsToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$helpToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolStripSeparator5 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$aboutToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$addNewRowToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$deleteRowToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$refreshDataToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolstripseparator6 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$openfiledialog1 = New-Object 'System.Windows.Forms.SaveFileDialog'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	
	#====================================================#
	#############    Update-DataGridview    ##############
	#====================================================#
	Function Update-DataGridView
	{
		Param (
			[ValidateNotNull()][Parameter(Mandatory = $true)][System.Windows.Forms.DataGridView]$DataGridView,
			[ValidateNotNull()][Parameter(Mandatory = $true)]$Item,
			[Parameter(Mandatory = $false)][string]$DataMember,
			[System.Windows.Forms.DataGridViewAutoSizeColumnsMode]$AutoSizeColumns = 'None'
		)
		$DataGridView.SuspendLayout()
		$DataGridView.DataMember = $DataMember
		
		If ($null -eq $Item)
		{
			$DataGridView.DataSource = $null
		}
		ElseIf ($Item -is [System.Data.DataSet] -and $Item.Tables.Count -gt 0)
		{
			$DataGridView.DataSource = $Item.Tables[0]
		}
		ElseIf ($Item -is [System.ComponentModel.IListSource]`
			-or $Item -is [System.ComponentModel.IBindingList] -or $Item -is [System.ComponentModel.IBindingListView])
		{
			$DataGridView.DataSource = $Item
		}
		Else
		{
			$array = New-Object System.Collections.ArrayList
			
			If ($Item -is [System.Collections.IList])
			{
				$array.AddRange($Item)
			}
			Else
			{
				$array.Add($Item)
			}
			$DataGridView.DataSource = $array
		}
		
		If ($AutoSizeColumns -ne 'None')
		{
			$DataGridView.AutoResizeColumns($AutoSizeColumns)
		}
		
		$DataGridView.ResumeLayout()
	}
	
	#====================================================#
	#############    ConvertTo-DataTable    ##############
	#====================================================#
	Function ConvertTo-DataTable
	{
		[OutputType([System.Data.DataTable])]
		Param (
			$InputObject,
			[ValidateNotNull()][System.Data.DataTable]$Table,
			[switch]$RetainColumns,
			[switch]$FilterWMIProperties)
		
		If ($null -eq $Table)
		{
			$Table = New-Object System.Data.DataTable
		}
		
		If ($null -eq $InputObject)
		{
			$Table.Clear()
			Return @( ,$Table)
		}
		
		If ($InputObject -is [System.Data.DataTable])
		{
			$Table = $InputObject
		}
		ElseIf ($InputObject -is [System.Data.DataSet] -and $InputObject.Tables.Count -gt 0)
		{
			$Table = $InputObject.Tables[0]
		}
		Else
		{
			If (-not $RetainColumns -or $Table.Columns.Count -eq 0)
			{
				#Clear out the Table Contents
				$Table.Clear()
				
				If ($null -eq $InputObject)
				{
					Return
				} #Empty Data
				
				$object = $null
				#find the first non null value
				ForEach ($item In $InputObject)
				{
					If ($null -ne $item)
					{
						$object = $item
						Break
					}
				}
				
				If ($null -eq $object)
				{
					Return
				} #All null then empty
				
				#Get all the properties in order to create the columns
				ForEach ($prop In $object.PSObject.Get_Properties())
				{
					If (-not $FilterWMIProperties -or -not $prop.Name.StartsWith('__')) #filter out WMI properties
					{
						#Get the type from the Definition string
						$type = $null
						
						If ($null -ne $prop.Value)
						{
							Try
							{
								$type = $prop.Value.GetType()
							}
							Catch
							{
								Out-Null
							}
						}
						
						If ($null -ne $type) # -and [System.Type]::GetTypeCode($type) -ne 'Object')
						{
							[void]$table.Columns.Add($prop.Name, $type)
						}
						Else #Type info not found
						{
							[void]$table.Columns.Add($prop.Name)
						}
					}
				}
				
				If ($object -is [System.Data.DataRow])
				{
					ForEach ($item In $InputObject)
					{
						$Table.Rows.Add($item)
					}
					Return @( ,$Table)
				}
			}
			Else
			{
				$Table.Rows.Clear()
			}
			
			ForEach ($item In $InputObject)
			{
				$row = $table.NewRow()
				
				If ($item)
				{
					ForEach ($prop In $item.PSObject.Get_Properties())
					{
						If ($table.Columns.Contains($prop.Name))
						{
							$row.Item($prop.Name) = $prop.Value
						}
					}
				}
				[void]$table.Rows.Add($row)
			}
		}
		
		Return @( ,$Table)
	}
	#====================================================#
	#############    SearchGrid Function    ##############
	#====================================================#
	Function SearchGrid()
	{
		$RowIndex = 0
		$ColumnIndex = 0
		$seachString = $textboxSearch.Text
		
		If ($seachString -eq "")
		{
			Return
		}
		
		If ($datagridview1.SelectedCells.Count -ne 0)
		{
			$startCell = $datagridview1.SelectedCells[0];
			$RowIndex = $startCell.RowIndex + 1
			$ColumnIndex = $startCell.ColumnIndex
		}
		
		$columnCount = $datagridview1.ColumnCount
		$rowCount = $datagridview1.RowCount
		For (; $RowIndex -lt $rowCount; $RowIndex++)
		{
			$Row = $datagridview1.Rows[$RowIndex]
			
			For (; $ColumnIndex -lt $columnCount; $ColumnIndex++)
			{
				$cell = $Row.Cells[$ColumnIndex]
				
				If ($null -ne $cell.Value -and $cell.Value.ToString().IndexOf($seachString, [StringComparison]::OrdinalIgnoreCase) -ne -1)
				{
					$datagridview1.CurrentCell = $cell
					Return
				}
			}
			
			$ColumnIndex = 0
		}
		
		$datagridview1.CurrentCell = $null
		[void][System.Windows.Forms.MessageBox]::Show("The search has reached the end of the grid.", "String not Found")
		
	}
	
	#====================================================#
	#############     Triggered Events     ###############
	#====================================================#
	$buttonSearch_Click = {
		SearchGrid
	}
	
	$datagridview1_ColumnHeaderMouseClick = [System.Windows.Forms.DataGridViewCellMouseEventHandler]{
		#Event Argument: $_ = [System.Windows.Forms.DataGridViewCellMouseEventArgs]
		If ($datagridview1.DataSource -is [System.Data.DataTable])
		{
			$column = $datagridview1.Columns[$_.ColumnIndex]
			$direction = [System.ComponentModel.ListSortDirection]::Ascending
			
			If ($column.HeaderCell.SortGlyphDirection -eq 'Descending')
			{
				$direction = [System.ComponentModel.ListSortDirection]::Descending
			}
			
			$datagridview1.Sort($datagridview1.Columns[$_.ColumnIndex], $direction)
		}
	}
	
	
	$textboxSearch_KeyUp = [System.Windows.Forms.KeyEventHandler]{
		#Event Argument: $_ = [System.Windows.Forms.KeyEventArgs]
		If ($_.KeyCode -eq 'Enter' -and $buttonSearch.Enabled)
		{
			SearchGrid
			$_.SuppressKeyPress = $true
		}
	}
	
	$datagridview1_CellContentClick = [System.Windows.Forms.DataGridViewCellEventHandler]{
		#Event Argument: $_ = [System.Windows.Forms.DataGridViewCellEventArgs]
		
	}
	
	$menustrip1_ItemClicked = [System.Windows.Forms.ToolStripItemClickedEventHandler]{
		#Event Argument: $_ = [System.Windows.Forms.ToolStripItemClickedEventArgs]
		
	}
	
	$fileToolStripMenuItem_Click = {
		
	}
	
	$ReloadToolStripMenuItem_Click = {
		$confirmReload = $YesNo = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to reload the table? '
Any unsaved changes will be lost.", "Confirm Reload", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		If ($YesNo -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			$datagridview1.Refresh()
		}
	}
	$saveToolStripMenuItem_Click = {
		$csvFile = "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv"
		$YesNo = [System.Windows.Forms.MessageBox]::Show("Update Inventory Database with Your Revisions? '
Saved Changes Made to Database Cannot Be Undone...", "Confirm Update", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		If ($YesNo -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			
			# Log the changes made
			$table = $datagridview1.datasource.GetChanges()
			$table.AcceptChanges()
			$table | Export-Csv -Path "$($env:APPDATA)/DPC_PartsDB/Change_Log/$(((get-date).ToUniversalTime()).ToString("yyyyMMddTHHmmssZ")).csv" -Force -NoTypeInformation
			
			# Export table for upload:
			$datagridview1.Rows |
			Select-Object -ExpandProperty DataBoundItem |
			Export-Csv -UseQuotes Always -Encoding Unicode -Path "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv" -Force
			
			# Query the SHA
			$shaCheck = (Get-GitHubContent -OwnerName Mediocre-Software -RepositoryName DPC_PartsDB -Path PartsDB.csv -AccessToken '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').sha
			# Upload revised table
			$fileContent = (Get-Content -Encoding Unicode -Path "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv" -Raw)
			(Set-GitHubContent -OwnerName Mediocre-Software -RepositoryName DPC_PartsDB -Path PartsDB.csv -CommitMessage 'Inventory Altered By End-User - Updating PartsDB.csv' -Content $fileContent -Sha $shaCheck -AccessToken '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')
			
			[System.Windows.Forms.MessageBox]::Show("Updates to the database were successful!", "Successful Update", [System.Windows.Forms.MessageBoxButtons]::OK)
		}
	}
	
	$saveAsCSVToolStripMenuItem_Click = {
			$csvFile = "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv"
			$table = $datagridview1.datasource.GetChanges()
			$table.AcceptChanges()
			$datagridview1.Rows |
			Select-Object -ExpandProperty DataBoundItem |
			Export-Csv -UseQuotes Always -Encoding Unicode -Path "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv" -Force
			$openfiledialog1.FileName = "$($env:USERPROFILE)/Desktop/PartsDB.csv"
			$openfiledialog1.CheckFileExists = $false
			$openfiledialog1.ShowDialog()
			$openfiledialog1_FileOk = [System.ComponentModel.CancelEventHandler]
			{
				[void][System.Windows.Forms.MessageBox]::Show($openfiledialog1.FileName, "Save As")
				$csvFile | Out-File $openfiledialog1.FileName -Force
			}
		}
		
		$printToolStripMenuItem_Click = {
			$csvFile = "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv"
			# Log the changes made
			$table = $datagridview1.datasource.GetChanges()
			$table.AcceptChanges()
			$table | Export-Csv -Path "$($env:APPDATA)/DPC_PartsDB/Change_Log/Print_$(((get-date).ToUniversalTime()).ToString("yyyyMMddTHHmmssZ")).csv" -Force -NoTypeInformation
			# Export table for upload:
			$datagridview1.Rows |
			Select-Object -ExpandProperty DataBoundItem |
			Format-List |
			Out-Printer
			
		}
		
		$printPreviewToolStripMenuItem_Click = {
			$csvFile = "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv"
			# Log the changes made
			$table = $datagridview1.datasource.GetChanges()
			$table.AcceptChanges()
			$table | Export-Csv -Path "$($env:APPDATA)/DPC_PartsDB/Change_Log/Print_$(((get-date).ToUniversalTime()).ToString("yyyyMMddTHHmmssZ")).csv" -Force -NoTypeInformation
			# Export table for upload:
			$datagridview1.Rows |
			Select-Object -ExpandProperty DataBoundItem |
			Format-List |
			Out-File -FilePath "$($env:APPDATA)/DPC_PartsDB/PartsList.txt"
			
			Invoke-Item -Path "$($env:APPDATA)/DPC_PartsDB/PartsList.txt"
		}
		
		$exitToolStripMenuItem_Click = {
			$YesNo = [System.Windows.Forms.MessageBox]::Show("Unsaved Changes Will Be Lost Forever. Exit Anyway?", "Exit", [System.Windows.Forms.MessageBoxButtons]::YesNo)
			If ($YesNo -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				[void]$form1.Close()
			}
		}
		
		$editToolStripMenuItem_Click = {
			
		}
		
		$undoToolStripMenuItem_Click = {
			
		}
		
		$redoToolStripMenuItem_Click = {
			
		}
		
		$cutToolStripMenuItem_Click = {
			$datagridview1.SelectedCells.ToString($copyCells)
			$datagridview1.SelectedCells.Clear()
			
		}
		
		$copyToolStripMenuItem_Click = {
			$datagridview1.SelectedCells.ToSTrying($copyCells)
			$datagridview1.SelectedCells.Clear()
		}
		
		$pasteToolStripMenuItem_Click = {
			$datagridview1.BeginEdit($true)
			$datagridview1.CurrentCell.Value($copyCells)
		}
		
		$selectAllToolStripMenuItem_Click = {
			$datagridview1.SelectAll()
		}
		
		$toolsToolStripMenuItem_Click = {
			
		}
		
		$helpToolStripMenuItem_Click = {
			
		}
		
		$aboutToolStripMenuItem_Click = {
			
		}
		
		$picturebox1_Click = {
			
		}
		
		$addNewRowToolStripMenuItem_Click = {
			$datagridview1.Rows.Add()
		}
		
		$deleteRowToolStripMenuItem_Click = {
			$datagridview1.SelectionMode = 'FullRowSelect'
			$dgvRow = $datagridview1.SelectedRows
			$confirmDeleteRow = [System.Windows.Forms.MessageBox]::Show("Do you really want to delete the selected Row(s)?", "Confirm Row Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo)
			If ($confirmDeleteRow -eq [System.Windows.forms.DialogResult]::Yes)
			{
				$datagridview1.Rows.Remove($dgvRow)
			}
		}
		
		$refreshDataToolStripMenuItem_Click = {
			$confirmRefresh = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to reload the table? '
Any unsaved changes will be lost.", "Confirm Refresh", [System.Windows.Forms.MessageBoxButtons]::YesNo)
			If ($confirmRefresh -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				$datagridview1.Refresh()
			}
		}
		
		#====================================================#
		#############    Form Loading Events    ##############
		#====================================================#
		
		$form1_Load = {
			
			# Assign Variables
			$icoPath = "$($env:APPDATA)/DPC_PartsDB/DataTable.ico"
			$logoPath = "$($env:APPDATA)/DPC_PartsDB/DPCLogo.png"
			$dpcPath = Join-Path "$($env:APPDATA)" "DPC_PartsDB"
			$changeDir = Join-Path "$($env:APPDATA)" "DPC_PartsDB/Change_Log"
			$icoExists = Test-Path $icoPath
			$logoExists = Test-Path $logoPath
			$dpcPathExists = Test-Path $dpcPath
			$changeExists = Test-Path $changeDir
			
			# Check File System For Main Directory	
			If ($dpcPathExists -eq $false)
			{
				New-Item -Path "$($env:APPDATA)" -Name "DPC_PartsDB" -ItemType "Directory"
			}
			
			# Check for Required Modules
			$modCheck1 = (Get-Module -ListAvailable -Name 'PowerShellForGitHub')
			If ($null -eq $modCheck1)
			{
				[System.Windows.Forms.MessageBox]::Show("Installing (1 of 1) Missing Dependencies.`
This Could Take A Few Moments To Complete...", "Missing Program Dependencies", [System.Windows.Forms.MessageBoxButtons]::OK)
				Install-Module  'PowerShellForGitHub' -Scope CurrentUser -AllowClobber -Force | Out-Null
			}
			
			# Import Modules		
			Import-Module PowerShellForGitHub | Out-Null
			
			# Download Database CSV File...	
			Get-GitHubContent -OwnerName Mediocre-Software -RepositoryName DPC_PartsDB -Path PartsDB.csv -AccessToken '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' -MediaType Raw -ResultAsString | Out-File -Encoding UTF8 -FilePath "$($env:APPDATA)/DPC_PartsDB/Parts.csv" -Force
			
			# Decode Csv from Unicode-16 -> UTF-8
			Get-Content -Encoding Unicode -Path "$($env:APPDATA)/DPC_PartsDB/Parts.csv" | Set-Content -Encoding UTF8 -Path "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv" -Force
			
			# Check for Remaining Dependency Files 
			If ($changeExists -eq $false)
			{
				New-Item -Path "$($env:APPDATA)/DPC_PartsDB" -Name "Change_Log" -ItemType "Directory"
			}
			If ($logoExists -ne $True)
			{
				Invoke-WebRequest -Uri https://raw.githubusercontent.com/Mediocre-Software/DPC_PartsDB/main/DPCLogo.png -OutFile "$($env:APPDATA)/DPC_PartsDB/DPCLogo.png"
			}
			If ($icoExists -ne $True)
			{
				Invoke-WebRequest -Uri https://raw.githubusercontent.com/Mediocre-Software/DPC_PartsDB/main/DataTable.ico -OutFile "$($env:APPDATA)/DPC_PartsDB/DataTable.ico"
			}
			
			# Load table into DGV
			$rows = Import-Csv -Encoding UTF8 -Path "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv"
			$table = ConvertTo-DataTable -InputObject $rows
			$table.AcceptChanges()
			Update-DataGridView -DataGridView $datagridview1 -Item $table
			
			$datagridview1.Columns[0].FillWeight = '8'  # DP/N
			$datagridview1.Columns[1].FillWeight = '55' # DESCRIPTION
			$datagridview1.Columns[2].FillWeight = '15' # LOCATION
			$datagridview1.Columns[3].FillWeight = '6'  # QTY
			$datagridview1.Columns[4].FillWeight = '8'  # $VALUE
			$datagridview1.Columns[5].FillWeight = '8'  # $TOTAL
			
		}
		
		$Form_StateCorrection_Load = {
			#Correct the initial state of the form to prevent the .Net maximized form issue
			$form1.WindowState = $InitialFormWindowState
		}
		
		#====================================================#
		#############    Form Closing Events    ##############
		#====================================================#
		
		$Form_Cleanup_FormClosed = {
			
			# If database has been altered - Prompt User to Save Changes
			$changeCheck = $datagridview1.datasource.GetChanges()
			If ($null -ne $changeCheck)
			{
				$YesNo = [System.Windows.Forms.MessageBox]::Show("Do you want to save your changes before Exiting? `
Saved Changes Cannot Be Undone...", "Save Changes?", [System.Windows.Forms.MessageBoxButtons]::YesNo)
				If ($YesNo -eq [System.Windows.Forms.DialogResult]::Yes)
				{
					# Log Changes to DB
					$table = $datagridview1.datasource.GetChanges()
					$table.AcceptChanges()
					$table | Export-Csv -Path "$($env:APPDATA)/DPC_PartsDB/Change_Log/$(((get-date).ToUniversalTime()).ToString("yyyyMMddTHHmmssZ")).csv" -Force -NoTypeInformation
					
					# Create CSV for Upload to DB
					$datagridview1.Rows |
					Select-Object -ExpandProperty DataBoundItem |
					Export-Csv -UseQuotes Always -Encoding UTF8 -Path "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv" -Force
					
					#ConvertTo-ExcelXlsx -Path "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv" -Force
					
					# Upload the Edited DB
					$shaCheck = (Get-GitHubContent -OwnerName Mediocre-Software -RepositoryName DPC_PartsDB -Path PartsDB.csv -AccessToken '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').sha
					$fileContent = Get-Content -Encoding UTF8 -Path "$($env:APPDATA)/DPC_PartsDB/PartsDB.csv" -Raw
					(Set-GitHubContent -OwnerName Mediocre-Software -RepositoryName DPC_PartsDB -Path PartsDB.csv -CommitMessage 'Inventory Altered By End-User' -Content $fileContent -Sha $shaCheck -AccessToken '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')
					
					# Confirm DB Update
					[System.Windows.Forms.MessageBox]::Show("Succesfully Applied Changes to the Database!", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK)
				}
			}
			#Remove all event handlers from the controls
			Try
			{
				$picturebox1.remove_Click($picturebox1_Click)
				$buttonSearch.remove_Click($buttonSearch_Click)
				$datagridview1.remove_CellContentClick($datagridview1_CellContentClick)
				$textboxSearch.remove_TextChanged($textboxSearch_TextChanged)
				$menustrip1.remove_ItemClicked($menustrip1_ItemClicked)
				$form1.remove_Load($form1_Load)
				$fileToolStripMenuItem.remove_Click($fileToolStripMenuItem_Click)
				$reloadToolStripMenuItem.remove_Click($reloadToolStripMenuItem_Click)
				$saveToolStripMenuItem.remove_Click($saveToolStripMenuItem_Click)
				$saveAsToolStripMenuItem.remove_Click($saveAsCSVToolStripMenuItem_Click)
				$printToolStripMenuItem.remove_Click($printToolStripMenuItem_Click)
				$printPreviewToolStripMenuItem.remove_Click($printPreviewToolStripMenuItem_Click)
				$exitToolStripMenuItem.remove_Click($exitToolStripMenuItem_Click)
				$editToolStripMenuItem.remove_Click($editToolStripMenuItem_Click)
				$undoToolStripMenuItem.remove_Click($undoToolStripMenuItem_Click)
				$redoToolStripMenuItem.remove_Click($redoToolStripMenuItem_Click)
				$cutToolStripMenuItem.remove_Click($cutToolStripMenuItem_Click)
				$copyToolStripMenuItem.remove_Click($copyToolStripMenuItem_Click)
				$pasteToolStripMenuItem.remove_Click($pasteToolStripMenuItem_Click)
				$selectAllToolStripMenuItem.remove_Click($selectAllToolStripMenuItem_Click)
				$toolsToolStripMenuItem.remove_Click($toolsToolStripMenuItem_Click)
				$helpToolStripMenuItem.remove_Click($helpToolStripMenuItem_Click)
				$aboutToolStripMenuItem.remove_Click($aboutToolStripMenuItem_Click)
				$addNewRowToolStripMenuItem.remove_Click($addNewRowToolStripMenuItem_Click)
				$deleteRowToolStripMenuItem.remove_Click($deleteRowToolStripMenuItem_Click)
				$refreshDataToolStripMenuItem.remove_Click($refreshDataToolStripMenuItem_Click)
				$form1.remove_Load($Form_StateCorrection_Load)
				$form1.remove_FormClosed($Form_Cleanup_FormClosed)
			}
			Catch
			{
				Out-Null <# Prevent PSScriptAnalyzer warning #>
			}
		}
		
		#====================================================#
		#############    Assign Form Controls    #############
		#====================================================#
		
		$form1.SuspendLayout()
		$menustrip1.SuspendLayout()
		#################
		####  form1  ####
		#################
		$form1.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Join-Path $env:APPDATA "DPC_PartsDB/DataTable.ico"))
		$form1.Controls.Add($picturebox1)
		$form1.Controls.Add($buttonSearch)
		$form1.Controls.Add($datagridview1)
		$form1.Controls.Add($textboxSearch)
		$form1.Controls.Add($menustrip1)
		$form1.AcceptButton = $buttonSearch
		$form1.ShowIcon = $True
		$form1.ShowInTaskbar = $True
		$form1.AutoScaleDimensions = '9, 19'
		$form1.AutoScaleMode = 'Inherit'
		$form1.BackColor = 'ControlLight'
		$form1.ClientSize = '927, 611'
		$form1.Cursor = 'Default'
		$form1.Font = 'Consolas, 8.25pt'
		$form1.MainMenuStrip = $menustrip1
		$form1.Margin = '5, 3, 5, 3'
		$form1.Name = 'form1'
		$form1.SizeGripStyle = 'Show'
		$form1.StartPosition = 'CenterScreen'
		$form1.Text = 'Discount PC - Parts Database'
		$form1.add_Load($form1_Load)
		#################
		## picturebox1 ##
		#################
		$img = [System.Drawing.Image]::FromFile((Join-Path $env:APPDATA "DPC_PartsDB/dpcLogo.png"))
		$picturebox1.Image = $img
		$picturebox1.BorderStyle = 'None'
		$picturebox1.BackgroundImageLayout = 'Stretch'
		$picturebox1.SizeMode = 'StretchImage'
		$picturebox1.Anchor = 'Top, Left, Right'
		$picturebox1.Location = '0, 20'
		$picturebox1.Margin = '0, 0, 0, 0'
		$picturebox1.Name = 'picturebox1'
		$picturebox1.Size = '927, 80'
		$picturebox1.TabIndex = 7
		$picturebox1.TabStop = $False
		##################
		## buttonSearch ##
		#################
		$buttonSearch.Anchor = 'Bottom, Right'
		$buttonSearch.FlatStyle = 'System'
		$buttonSearch.Location = '828, 577'
		$buttonSearch.Margin = '4, 3, 4, 3'
		$buttonSearch.Name = 'buttonSearch'
		$buttonSearch.Size = '88, 27'
		$buttonSearch.TabIndex = 1
		$buttonSearch.TabIndex = $True
		$buttonSearch.Text = '&Search'
		$buttonSearch.UseCompatibleTextRendering = $True
		$buttonSearch.UseVisualStyleBackColor = $True
		$buttonSearch.add_Click($buttonSearch_Click)
		###################
		## datagridview1 ##
		###################
		$System_Windows_Forms_DataGridViewCellStyle_1 = New-Object 'System.Windows.Forms.DataGridViewCellStyle'
		$System_Windows_Forms_DataGridViewCellStyle_1.Alignment = 'MiddleCenter'
		$System_Windows_Forms_DataGridViewCellStyle_1.BackColor = 'LightGray'
		$System_Windows_Forms_DataGridViewCellStyle_1.Font = 'Consolas, 8.5pt'
		$System_Windows_Forms_DataGridViewCellStyle_1.ForeColor = 'Black'
		$System_Windows_Forms_DataGridViewCellStyle_1.SelectionBackColor = 'LightSteelBlue'
		$System_Windows_Forms_DataGridViewCellStyle_1.SelectionForeColor = 'Black'
		$System_Windows_Forms_DataGridViewCellStyle_1.WrapMode = 'False'
		$datagridview1.AlternatingRowsDefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_1
		$System_Windows_Forms_DataGridViewCellStyle_2 = New-Object 'System.Windows.Forms.DataGridViewCellStyle'
		$System_Windows_Forms_DataGridViewCellStyle_2.Alignment = 'MiddleCenter'
		$System_Windows_Forms_DataGridViewCellStyle_2.BackColor = '0, 0, 64'
		$System_Windows_Forms_DataGridViewCellStyle_2.Font = 'Consolas, 10pt'
		$System_Windows_Forms_DataGridViewCellStyle_2.ForeColor = 'HighlightText'
		$System_Windows_Forms_DataGridViewCellStyle_2.SelectionBackColor = '0, 0, 85'
		$System_Windows_Forms_DataGridViewCellStyle_2.SelectionForeColor = 'DarkGoldenrod'
		$System_Windows_Forms_DataGridViewCellStyle_2.WrapMode = 'False'
		$datagridview1.ColumnHeadersDefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_2
		$System_Windows_Forms_DataGridViewCellStyle_3 = New-Object 'System.Windows.Forms.DataGridViewCellStyle'
		$System_Windows_Forms_DataGridViewCellStyle_3.Alignment = 'MiddleCenter'
		$System_Windows_Forms_DataGridViewCellStyle_3.BackColor = 'DimGray'
		$System_Windows_Forms_DataGridViewCellStyle_3.Font = 'Consolas, 8.5pt'
		$System_Windows_Forms_DataGridViewCellStyle_3.ForeColor = 'White'
		$System_Windows_Forms_DataGridViewCellStyle_3.SelectionBackColor = 'LightSteelBlue'
		$System_Windows_Forms_DataGridViewCellStyle_3.SelectionForeColor = 'Black'
		$System_Windows_Forms_DataGridViewCellStyle_3.WrapMode = 'False'
		$datagridview1.RowsDefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_3
		# 	$System_Windows_Forms_DataGridViewCellStyle_4 = New-Object 'System.Windows.Forms.DataGridViewCellStyle'
		# 	$System_Windows_Forms_DataGridViewCellStyle_4.Alignment = 'MiddleCenter'
		# 	$System_Windows_Forms_DataGridViewCellStyle_4.BackColor = 'DimGray'
		# 	$System_Windows_Forms_DataGridViewCellStyle_4.Font = 'Consolas, 8.25pt'
		# 	$System_Windows_Forms_DataGridViewCellStyle_4.ForeColor = 'White'
		# 	$System_Windows_Forms_DataGridViewCellStyle_4.SelectionBackColor = 'IndianRed'
		# 	$System_Windows_Forms_DataGridViewCellStyle_4.SelectionForeColor = 'White'
		# 	$System_Windows_Forms_DataGridViewCellStyle_4.WrapMode = 'False'
		# 	$datagridview1.DefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_4
		$datagridview1.AllowUserToDeleteRows = $False
		$datagridview1.AllowUserToResizeColumns = $False
		$datagridview1.AllowUserToOrderColumns = $False
		$datagridview1.AllowUserToResizeRows = $False
		$datagridview1.Anchor = 'Top, Bottom, Left, Right'
		$datagridview1.AutoSizeColumnsMode = 'Fill'
		$datagridview1.AutoSizeRowsMode = 'AllCells'
		$datagridview1.BackgroundColor = 'ControlDarkDark'
		$datagridview1.BorderStyle = 'FixedSingle'
		$datagridview1.ClipboardCopyMode = 'EnableWithoutHeaderText'
		$datagridview1.ColumnHeadersBorderStyle = 'Raised'
		$datagridview1.ColumnHeadersHeightSizeMode = 'AutoSize'
		$datagridview1.ColumnHeadersVisible = $True
		$datagridview1.EditMode = 'EditOnKeystroke'
		$datagridview1.EnableHeadersVisualStyles = $False
		$datagridview1.GridColor = 'ControlDarkDark'
		$datagridview1.Location = '7, 122'
		$datagridview1.MultiSelect = $True
		$datagridview1.Margin = '4, 3, 4, 3'
		$datagridview1.Name = 'datagridview1'
		$datagridview1.RowHeadersVisible = $False
		$datagridview1.RowHeadersWidthSizeMode = 'DisableResizing'
		$datagridview1.RowTemplate.Height = 20
		$datagridview1.ScrollBars = 'Vertical'
		$datagridview1.SelectionMode = 'FullRowSelect'
		$datagridview1.ShowEditingIcon = $False
		$datagridview1.StandardTab = $False
		$datagridview1.Size = '914, 450'
		$datagridview1.TabIndex = 2
		$datagridview1.TabStop = $False
		###################
		## textboxSearch ##
		###################
		$textboxSearch.AcceptsReturn = $True
		$textboxSearch.AcceptsTab = $True
		$textboxSearch.Anchor = 'Bottom, Left, Right'
		$textboxSearch.BorderStyle = 'FixedSingle'
		$textboxSearch.Font = 'Consolas, 7.5pt'
		$textboxSearch.HideSelection = $False
		$textboxSearch.Location = '11, 581'
		$textboxSearch.Margin = '5, 3, 5, 3'
		$textboxSearch.Name = 'textboxSearch'
		$textboxSearch.Size = '811, 23'
		$textboxSearch.TabIndex = 0
		$textboxSearch.TabStop = $True
		$textboxSearch.add_TextChanged($textboxSearch_TextChanged)
		##################
		##  menustrip1  ##
		##################
		$menustrip1.AllowMerge = $False
		$menustrip1.AutoSize = $False
		$menustrip1.BackColor = 'ControlLight'
		$menustrip1.ImageScalingSize = '17, 17'
		$menustrip1.Font = 'Consolas, 8.25pt'
		[void]$menustrip1.Items.Add($fileToolStripMenuItem)
		[void]$menustrip1.Items.Add($editToolStripMenuItem)
		[void]$menustrip1.Items.Add($toolsToolStripMenuItem)
		[void]$menustrip1.Items.Add($helpToolStripMenuItem)
		$menustrip1.Location = '0, 0'
		$menustrip1.Name = 'menustrip1'
		$menustrip1.Padding = '0, 0, 0, 0'
		$menustrip1.RenderMode = 'Professional'
		$menustrip1.Size = '927, 25'
		$menustrip1.TabIndex = 3
		$menustrip1.TabStop = $False
		$menustrip1.Text = 'menustrip1'
		$menustrip1.add_ItemClicked($menustrip1_ItemClicked)
		#########################
		# fileToolStripMenuItem #
		#########################
		[void]$fileToolStripMenuItem.DropDownItems.Add($reloadToolStripMenuItem)
		[void]$fileToolStripMenuItem.DropDownItems.Add($toolStripSeparator)
		[void]$fileToolStripMenuItem.DropDownItems.Add($saveToolStripMenuItem)
		[void]$fileToolStripMenuItem.DropDownItems.Add($saveAsToolStripMenuItem)
		[void]$fileToolStripMenuItem.DropDownItems.Add($toolStripSeparator1)
		[void]$fileToolStripMenuItem.DropDownItems.Add($printToolStripMenuItem)
		[void]$fileToolStripMenuItem.DropDownItems.Add($printPreviewToolStripMenuItem)
		[void]$fileToolStripMenuItem.DropDownItems.Add($toolStripSeparator2)
		[void]$fileToolStripMenuItem.DropDownItems.Add($exitToolStripMenuItem)
		$fileToolStripMenuItem.Name = 'fileToolStripMenuItem'
		$fileToolStripMenuItem.Size = '46, 21'
		$fileToolStripMenuItem.Text = '&File'
		$fileToolStripMenuItem.add_Click($fileToolStripMenuItem_Click)
		#
		# reloadToolStripMenuItem
		#
		#region Binary Data
		$reloadToolStripMenuItem.Image = [System.Convert]::FromBase64String('
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACx
jwv8YQUAAAAJcEhZcwAAEzkAABM5AY/CVgEAAAERSURBVDhPrZDbSgJRGIXnpewd6jXsjSQvIrwo
I0RQMChU0iiDPCGiE3ZCRkvR8VzTeBhnyR5/ccaZNnPhB4t9sdf6Ln5hb8QeathNJFVFKF5C8DqL
4ksDVHWGDf7jLHyPg6NjviSaFqlu5yQYR+KpupaIkrMknCxT3Y7v/NYYb0ITK1c3BarbWWhLQ7IR
0cTKReyZ6lZ0XYeiztHpK4bAc+h1FgQijzSxMptrGIxVSO0xX3AaStFki7bUMVFmaMm/eJMGfIH/
MkGzLep0AXn4h/r3CJV3mS9gn2bY4UY/UzQ7E9TqfeTFtnuB+XAfzSHKr11kSl/uBebDiZ89ZCst
3OUkdwL28sIVsE83ock+EIQV2Mz2wxeg6/UAAAAASUVORK5CYII=')
		#endregion
		$reloadToolStripMenuItem.ImageTransparentColor = 'Magenta'
		$reloadToolStripMenuItem.Name = 'reloadToolStripMenuItem'
		$reloadToolStripMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::R -bor [System.Windows.Forms.Keys]::Control
		$reloadToolStripMenuItem.Size = '191, 28'
		$reloadToolStripMenuItem.Text = '&Reload'
		$reloadToolStripMenuItem.ToolTipText = 'Reload Inventory Data'
		$reloadToolStripMenuItem.add_Click($reloadToolStripMenuItem_Click)
		#
		# toolStripSeparator
		#
		$toolStripSeparator.Name = 'toolStripSeparator'
		$toolStripSeparator.Size = '188, 6'
		#
		# saveToolStripMenuItem
		#
		#region Binary Data
		$saveToolStripMenuItem.Image = [System.Convert]::FromBase64String('
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACx
jwv8YQUAAAAJcEhZcwAAEzkAABM5AY/CVgEAAAIySURBVDhPrZLfS5NRGMfff6H7boIuuq2pMZyL
1eAt11CWDcOKsB9vpFmaLtNExco0av6CbIVLJ61Wk3BSkT/AFCkRZSpZmrmiJQ41xSaCwdfznL15
XEUX0Reem5f38znnec4j/Zc8fxYGla91CS3eRTx0z6OpMYS7jmnU1X6B/VYA18snUVoyjsKCt8jL
HcH5c36ouCQR2NUJ1Nas4G9ZXlmFKbULh1Kf8lJxSfI+WeCCyopv6q+/h+DQ/DJ2WV5Ao1FgPegR
AveDOS4oLfmq/h6dn/DH4AJizD4UXJrCAUuzEDgbZrjgou2DiohshIcnQtgme5GTPYbkJKcQ1N8O
ckHW2REVi+RXuM8fxGaDG4oyALPZIQQ11Z+5QDk1oKJ/hjv7P2FTfCMOH3mFxMQ6IbhROYWOdrCn
BI4dfwPr0V4+bRoY9UzXppMjcDdSrC8hy3YhuFI2gTYf2A4Aza4f7N2/o/zaLB8qDYx6zszwr8P7
k1thNFYIweXCMXgeAfedq2xxwjClZUeVJd2GtDNFETiJwfs8MBjKhMCWN8pgoLoqzE8miH1GjE7G
4PsZjE7OQsm9ij2mFg7rdrug1xcJAa2l4w7Wr00Cgk/n38S7wBwC04u4UGxHrMHF4CbEJtyDLj5f
CDIzhljfSxzeavRgyw4Zj9t64GvvQ0d3P3pfD2Kv2QqNvgFxDN6urYdWmyMElJMnevh60obRktA7
01PRtGlg1DOdSkXwzrisaMG/RZLWAE60OMW5fNhvAAAAAElFTkSuQmCC')
		#endregion
		$saveToolStripMenuItem.ImageTransparentColor = 'Magenta'
		$saveToolStripMenuItem.Name = 'saveToolStripMenuItem'
		$saveToolStripMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::S -bor [System.Windows.Forms.Keys]::Control
		$saveToolStripMenuItem.Size = '191, 28'
		$saveToolStripMenuItem.Text = '&Save'
		$saveToolStripMenuItem.add_Click($saveToolStripMenuItem_Click)
		#
		# saveAsToolStripMenuItem
		#
		$saveAsToolStripMenuItem.Name = 'saveAsToolStripMenuItem'
		$saveAsToolStripMenuItem.Size = '191, 28'
		$saveAsToolStripMenuItem.Text = 'Save &As'
		$saveAsToolStripMenuItem.add_Click($saveAsCSVToolStripMenuItem_Click)
		#
		# toolStripSeparator1
		#
		$toolStripSeparator1.Name = 'toolStripSeparator1'
		$toolStripSeparator1.Size = '188, 6'
		#
		# printToolStripMenuItem
		#
		#region Binary Data
		$printToolStripMenuItem.Image = [System.Convert]::FromBase64String('
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACx
jwv8YQUAAAAJcEhZcwAAEzkAABM5AY/CVgEAAAIpSURBVDhPtZL/T1JRGMb5p1itrVZbbRpqZbaw
nBENV1I0jGlByTSyJTXJwq2oKZQb1KAv6JCYWSxvBrkkZUq4CeQEiRABFeLL072Xa0zRra31bO8v
57zP5znnPYf1X+TxhWF6O7VtGYcnwbSWijKPOLzYrPSvLPwLS3huGUMlT7o9wGD9grVUBj+icdid
03S9tDmgNxNwTgVQJ+rA8XNtWwM+uuZATMwxmQVRycuJFNyzIRitDlScugKzjSgFRGJJaIwEsrk8
AsHIhnSL/Ssck37UNipQI5DjtuYV7uksRYhr2kebhx2eP6nrycFIEh5fBA/1Nvru8q5+PDaOovK0
rABwfwugWzcErfkzHhjsePL6E7q1VrTdNUDcrgGvSYlDZHN5XTNOnL8BVe8AJAoNDtZfLgDu9L1B
PJmikzcrk81hlRwodZJwdBXziwnIOrVoaOkiT8C8hKLHBPO7CbywOaE1jeC+bhAd6meQdvZC1KoG
/5IS3MZ2HObLUHZSggvkWq3wOvbWiAqAVpWeyStVfCUNf3AZ4zNhfHCFMEDMgye+hYr6FrDLzxQA
UuVTpr0ocn74mchg5vsKRt1RcHp2Qv9+kZ78UcE17KkWFgHNN/uQzgBkGKLJPBZiecyGchjzrmFw
PIF++xJUbDbUQzEacIArLpopSRSP4CUN1Obf1Abzuqob5KjiXwWH/GVl5HPt5zZh37GL2H1EiF1V
Z7GDI6CNW5r/TSzWbwHYL0mKJ5czAAAAAElFTkSuQmCC')
		#endregion
		$printToolStripMenuItem.ImageTransparentColor = 'Magenta'
		$printToolStripMenuItem.Name = 'printToolStripMenuItem'
		$printToolStripMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::P -bor [System.Windows.Forms.Keys]::Control
		$printToolStripMenuItem.Size = '191, 28'
		$printToolStripMenuItem.Text = '&Print'
		$printToolStripMenuItem.add_Click($printToolStripMenuItem_Click)
		#
		# printPreviewToolStripMenuItem
		#
		#region Binary Data
		$printPreviewToolStripMenuItem.Image = [System.Convert]::FromBase64String('
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACx
jwv8YQUAAAAJcEhZcwAAEzkAABM5AY/CVgEAAAGCSURBVDhPnZK9S0JRGMb9F1xb2gqaq6mhwCGD
tvYIIyLIcJOE1paoIYpMKUjFRDH87lpoakGlIZF9DA2hZJEQhJXl1xPn3HPV29WQfvBwOfA+P95z
uDJ39A6/4wylYOOSMHvOcHGThuwvSKEVRvsR+pQqWD3R1pK98DUbl7Jm5hA8SfESd6S5xH5wycal
rO4E0D8yWQuriLH6E2xcSqlcoRJBxCpiTO5TNi4m/ZgDF4nDsOulsfujyGRzUsmWM8YqdcggKbve
S3A88bEkslRye58RSzZtIVarY/FFaPmlwp+fUaESYRNW5Vm3BPmpBpZNvppACDmTLbS6FbGAPFAj
5OGI4PALOK/yZfIlAlk4j7n5xdaCarWKj0KRXmE2+UklJEJZZ/RCPTPdWvBdLOP1rYD41QNcgRiV
kKJQ1mjGsa2VNxeQb2OWDC7sh47pddQLeoyOTSFiVAAFvVhChsmv2k6Uvd3Icx1UolMNiDdpl4nh
LiohW/xb0tMph2JwCJxjAz9A30JI8zYAtAAAAABJRU5ErkJggg==')
		#endregion
		$printPreviewToolStripMenuItem.ImageTransparentColor = 'Magenta'
		$printPreviewToolStripMenuItem.Name = 'printPreviewToolStripMenuItem'
		$printPreviewToolStripMenuItem.Size = '191, 28'
		$printPreviewToolStripMenuItem.Text = 'Print Pre&view'
		$printPreviewToolStripMenuItem.add_Click($printPreviewToolStripMenuItem_Click)
		#
		# toolStripSeparator2
		#
		$toolStripSeparator2.Name = 'toolStripSeparator2'
		$toolStripSeparator2.Size = '188, 6'
		#
		# exitToolStripMenuItem
		#
		$exitToolStripMenuItem.Name = 'exitToolStripMenuItem'
		$exitToolStripMenuItem.Size = '191, 28'
		$exitToolStripMenuItem.Text = 'E&xit'
		$exitToolStripMenuItem.add_Click($exitToolStripMenuItem_Click)
		#
		# editToolStripMenuItem
		#
		[void]$editToolStripMenuItem.DropDownItems.Add($undoToolStripMenuItem)
		[void]$editToolStripMenuItem.DropDownItems.Add($redoToolStripMenuItem)
		[void]$editToolStripMenuItem.DropDownItems.Add($toolStripSeparator3)
		[void]$editToolStripMenuItem.DropDownItems.Add($cutToolStripMenuItem)
		[void]$editToolStripMenuItem.DropDownItems.Add($copyToolStripMenuItem)
		[void]$editToolStripMenuItem.DropDownItems.Add($pasteToolStripMenuItem)
		[void]$editToolStripMenuItem.DropDownItems.Add($toolStripSeparator4)
		[void]$editToolStripMenuItem.DropDownItems.Add($selectAllToolStripMenuItem)
		$editToolStripMenuItem.Name = 'editToolStripMenuItem'
		$editToolStripMenuItem.Size = '48, 21'
		$editToolStripMenuItem.Text = '&Edit'
		$editToolStripMenuItem.add_Click($editToolStripMenuItem_Click)
		#
		# undoToolStripMenuItem
		#
		$undoToolStripMenuItem.Name = 'undoToolStripMenuItem'
		$undoToolStripMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Z -bor [System.Windows.Forms.Keys]::Control
		$undoToolStripMenuItem.Size = '180, 28'
		$undoToolStripMenuItem.Text = '&Undo'
		$undoToolStripMenuItem.add_Click($undoToolStripMenuItem_Click)
		#
		# redoToolStripMenuItem
		#
		$redoToolStripMenuItem.Name = 'redoToolStripMenuItem'
		$redoToolStripMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Y -bor [System.Windows.Forms.Keys]::Control
		$redoToolStripMenuItem.Size = '180, 28'
		$redoToolStripMenuItem.Text = '&Redo'
		$redoToolStripMenuItem.add_Click($redoToolStripMenuItem_Click)
		#
		# toolStripSeparator3
		#
		$toolStripSeparator3.Name = 'toolStripSeparator3'
		$toolStripSeparator3.Size = '177, 6'
		#
		# cutToolStripMenuItem
		#
		#region Binary Data
		$cutToolStripMenuItem.Image = [System.Convert]::FromBase64String('
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACx
jwv8YQUAAAAJcEhZcwAAEzkAABM5AY/CVgEAAAGDSURBVDhPrZFNSwJRGIX9NYGbFoUlFElY1EJQ
KEYhCJsiLaVsERnRF5iCaSZJO1toCDVGFkgoFpWQWWRR2aIvUxm1BKN1wSnHCFw4TOCzue+9nPNw
4eVVnav4IzzbQfxeGZ5TWaxT/rK3irzmC7CsusvC1G4IkbNLboIiDieF4GGUKeTeClDpppF8eeEu
2PIfwfrzizSdw3HkEnKlFpkMzV2wH77AosOFTV8A+vkl9CiHuJeLJNNZjM8tYWB0FkTvMAwmy/8E
RTR6CwjlGAi1Ccence6C1NsXzN4PKIxJLLgeIJ2MoXvmFraNBKK3eXZRIveJPvs7FIYniEkXZENO
dE+GIZ2Ko10TwLK7tJmKmL0FEEYarYM+NMnt0C1sQzpx/lcSEnZ2gcKY/gs0dlmZuWvmjjmpwA1q
xVp2AWFIMAF/OAGBzMjMI7ZrtJCb4Df3o4Zfxy7QrdxDRFKol5khkpR2H4qmIOzUQNBGwrsXYxcc
nNOQqNbQ0KGGZ+eEPVwdeLxvqqrf4wGhTNAAAAAASUVORK5CYII=')
		#endregion
		$cutToolStripMenuItem.ImageTransparentColor = 'Magenta'
		$cutToolStripMenuItem.Name = 'cutToolStripMenuItem'
		$cutToolStripMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::X -bor [System.Windows.Forms.Keys]::Control
		$cutToolStripMenuItem.Size = '180, 28'
		$cutToolStripMenuItem.Text = 'Cu&t'
		$cutToolStripMenuItem.add_Click($cutToolStripMenuItem_Click)
		#
		# copyToolStripMenuItem
		#
		#region Binary Data
		$copyToolStripMenuItem.Image = [System.Convert]::FromBase64String('
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACx
jwv8YQUAAAAJcEhZcwAAEzkAABM5AY/CVgEAAAHkSURBVDhPvZHfS1NhHIf3p5QypLr2D4goMwoM
Ci/qIugHXe1Cr7qKDIMkZixwNhfWLGWbnuki0kXKzLU023KubBNPJrbRdOzocm6e2dPOO21mMS+C
HvjcvOf9PF++79H9M+7RT2iRRsIi9sEAXe43yAvf2LpSHq28G9uAnytNT4jMLewtcQ2Ht2pF8ps/
aOt+gccX5lxD694S+1BQFD1RkN5DSFa4Z3uONKbgHE3h8KZ4OJTC1J8UiSzmfhd2uf1CoJHbyKOs
Zokl0kKwm+aeJaov+wjOrpQkVqdXfOz0bWAcVLghfaXxkUz3y2VxvpMGSwL3uMKh+gHezSSLEnNh
X23vtYzKUirDfGyFj/Iy1mdxUWqR8iKhwtQLxjgH659y4EwvVXWPiwJt3/Ws+muywRrlqvkDdx3z
QrCN8l1ldnEd3/QqFmkS/akHJYGSzjLzOUEwEsMf+sLI2zmaOou/93pPGoM5zvk7UU7fnBKxSBPo
T7SXBNW1F/9Io2lKCNTCeomUyrS8xnBAwfUqyf1eP5U1ptJD/o1LzeNCsHPydtqdr6k4aiwvOHvN
Sya3ibU/QIdrEkvfhJislc32MfYfuV1eUGPwFF7bIVJVZ0N/soPK421UHGstlFvYd/hWecF/Qqf7
CR0A5wwgSQA2AAAAAElFTkSuQmCC')
		#endregion
		$copyToolStripMenuItem.ImageTransparentColor = 'Magenta'
		$copyToolStripMenuItem.Name = 'copyToolStripMenuItem'
		$copyToolStripMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::C -bor [System.Windows.Forms.Keys]::Control
		$copyToolStripMenuItem.Size = '180, 28'
		$copyToolStripMenuItem.Text = '&Copy'
		$copyToolStripMenuItem.add_Click($copyToolStripMenuItem_Click)
		#
		# pasteToolStripMenuItem
		#
		#region Binary Data
		$pasteToolStripMenuItem.Image = [System.Convert]::FromBase64String('
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACx
jwv8YQUAAAAJcEhZcwAAEzkAABM5AY/CVgEAAAJSSURBVDhPtZJrSJNRGMdf6IN9KbpQn/pUEH2J
IoLqQ0ZhFqYZRmJG1iKmUqKyLB2pqSm6vC1Nm5GXoeatEsVJ0RASR3eNzegikRq5lrV3857Fr/d9
ddlICoL+8OfAOef/e57zcIT/os7WLMw302muSGJ2689qqi7A44q8IzjtNYzarzHQm8tZtT8FmRqu
6LToMxN+B8qhCbGRKVcDE85ajKUaxoaryEuL4UVXIudPB5Ko2oy98xjDptXERuz3hsgAOTzlqqMk
6yjdllzE90UM9Wp5azlBS1kwkeG+1CSv4mmBQPThfd6Ahqq8GYB4A11yBKmaMLQxoZyLDkGjDiZO
FUhUuB+FsWsUQFiArzegtlzHpFjPpMPA2GA2jucx2KqWK7ZWLqO7dBGP9D5KWLbfto3eAKMhi3FH
BeP9GYy9PMXos4OIrYvJrzSRbWjmwuV6EnVG4tLLiEzSExGf4w0oL05nZEDPaK+akceBuO9v4uPt
FUrYo6npbzhdE/QPOQmNSiPouHYOUpafgvgqA/dDf9wd63G1r2SgUlAqyyq/1anYUGfG2mdXwne7
bOwJUc1AinOS+NxzBpd5HWLbUhyNPvRdF5S2v05/54tbqvzBifWNHUvPOwLC4/CXwrv2HsB3+w6E
wosJOB5ESeElfGpayGD1AmwlArHSm+W2PR1clTooMrbT0mFTVtlbN6xFuJQar3wQz5Q9VksD+7Xy
PctrJdx4p5s605M5gKz8lJPSDwtGFbKboJ1blAN52vKbPdXm80/AfDokTVu+8DfPXv9XCcIPTvjv
LQ8YoakAAAAASUVORK5CYII=')
		#endregion
		$pasteToolStripMenuItem.ImageTransparentColor = 'Magenta'
		$pasteToolStripMenuItem.Name = 'pasteToolStripMenuItem'
		$pasteToolStripMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::V -bor [System.Windows.Forms.Keys]::Control
		$pasteToolStripMenuItem.Size = '180, 28'
		$pasteToolStripMenuItem.Text = '&Paste'
		$pasteToolStripMenuItem.add_Click($pasteToolStripMenuItem_Click)
		#
		# toolStripSeparator4
		#
		$toolStripSeparator4.Name = 'toolStripSeparator4'
		$toolStripSeparator4.Size = '177, 6'
		#
		# selectAllToolStripMenuItem
		#
		$selectAllToolStripMenuItem.Name = 'selectAllToolStripMenuItem'
		$selectAllToolStripMenuItem.Size = '180, 28'
		$selectAllToolStripMenuItem.Text = 'Select &All'
		$selectAllToolStripMenuItem.add_Click($selectAllToolStripMenuItem_Click)
		#
		# toolsToolStripMenuItem
		#
		[void]$toolsToolStripMenuItem.DropDownItems.Add($refreshDataToolStripMenuItem)
		[void]$toolsToolStripMenuItem.DropDownItems.Add($toolstripseparator6)
		[void]$toolsToolStripMenuItem.DropDownItems.Add($addNewRowToolStripMenuItem)
		[void]$toolsToolStripMenuItem.DropDownItems.Add($deleteRowToolStripMenuItem)
		$toolsToolStripMenuItem.Name = 'toolsToolStripMenuItem'
		$toolsToolStripMenuItem.Size = '57, 21'
		$toolsToolStripMenuItem.Text = '&Tools'
		$toolsToolStripMenuItem.add_Click($toolsToolStripMenuItem_Click)
		#
		# helpToolStripMenuItem
		#
		[void]$helpToolStripMenuItem.DropDownItems.Add($toolStripSeparator5)
		[void]$helpToolStripMenuItem.DropDownItems.Add($aboutToolStripMenuItem)
		$helpToolStripMenuItem.Name = 'helpToolStripMenuItem'
		$helpToolStripMenuItem.Size = '54, 21'
		$helpToolStripMenuItem.Text = '&Help'
		$helpToolStripMenuItem.add_Click($helpToolStripMenuItem_Click)
		#
		# toolStripSeparator5
		#
		$toolStripSeparator5.Name = 'toolStripSeparator5'
		$toolStripSeparator5.Size = '135, 6'
		#
		# aboutToolStripMenuItem
		#
		$aboutToolStripMenuItem.Name = 'aboutToolStripMenuItem'
		$aboutToolStripMenuItem.Size = '138, 28'
		$aboutToolStripMenuItem.Text = '&About...'
		$aboutToolStripMenuItem.add_Click($aboutToolStripMenuItem_Click)
		#
		# addNewRowToolStripMenuItem
		#
		$addNewRowToolStripMenuItem.Name = 'addNewRowToolStripMenuItem'
		$addNewRowToolStripMenuItem.Size = '179, 28'
		$addNewRowToolStripMenuItem.Text = '&New Row'
		$addNewRowToolStripMenuItem.add_Click($addNewRowToolStripMenuItem_Click)
		#
		# deleteRowToolStripMenuItem
		#
		$deleteRowToolStripMenuItem.Name = 'deleteRowToolStripMenuItem'
		$deleteRowToolStripMenuItem.Size = '179, 28'
		$deleteRowToolStripMenuItem.Text = '&Delete Row'
		$deleteRowToolStripMenuItem.add_Click($deleteRowToolStripMenuItem_Click)
		#
		# refreshDataToolStripMenuItem
		#
		$refreshDataToolStripMenuItem.Name = 'refreshDataToolStripMenuItem'
		$refreshDataToolStripMenuItem.Size = '179, 28'
		$refreshDataToolStripMenuItem.Text = '&Refresh Table'
		$refreshDataToolStripMenuItem.add_Click($refreshDataToolStripMenuItem_Click)
		#
		# toolstripseparator6
		#
		$toolstripseparator6.Name = 'toolstripseparator6'
		$toolstripseparator6.Size = '176, 6'
		#
		# savefiledialog1
		#
		$openfiledialog1.AutoUpgradeEnabled = $False
		$openfiledialog1.CreatePrompt = $True
		$openfiledialog1.DefaultExt = 'csv'
		$openfiledialog1.FileName = 'DPC_PartsDB'
		$openfiledialog1.Filter = 'CSV Files|*.csv'
		$openfiledialog1.InitialDirectory = "$($env:USERPROFILE)\Downloads"
		$openfiledialog1.Title = 'Save As'
		$openfiledialog1.ValidateNames = $False
		$menustrip1.ResumeLayout()
		$form1.ResumeLayout()
		#endregion Generated Form Code
		
		#----------------------------------------------
		
		#Save the initial state of the form
		$InitialFormWindowState = $form1.WindowState
		#Init the OnLoad event to correct the initial state of the form
		$form1.add_Load($Form_StateCorrection_Load)
		#Clean up the control events
		$form1.add_FormClosed($Form_Cleanup_FormClosed)
		#Show the Form
		Return $form1.ShowDialog()
	
}

#Call the form
Show-PartsDB | Out-Null