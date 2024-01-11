using namespace System.Collections.Generic
using namespace System.Management.Automation
using namespace System.Windows.Forms
using namespace System.Drawing


function Open-Agenda {
<# 
.SYNOPSIS 
   Open agenda application 
.DESCRIPTION
   This function initializes a new instance of the Agenda class and opens the
   application for use.
#> 
	$App = New-Object Agenda
	$App.Open()
}

function Add-EventWrapper {
<# 
.SYNOPSIS 
   Wrap an event handler to preserve class instance 
.DESCRIPTION
   This function invokes a script block or method within a new closure of curly
   braces, preserving the invocation before the event is raised. This prevents
   contamination of automatic variable $this and allows the script block or
   method to reference a class instance. Without this function, $this references
   the event sender.
   -----------------------------------------------------------------------------
   This function is adapted from code written by user mklement0 on Stack Overflow.
   The link is provided below.
.PARAMETER Method 
   Class method that handles an event 
.PARAMETER ScriptBlock 
   Script block that handlees an event
.PARAMETER SendArgs 
   Values from automatic variable $Args are preserved
.LINK 
   https://stackoverflow.com/a/64236498
#> 
	param (
		[Parameter()]
		[PSMethod] $Method,
		[Parameter()]
		[ScriptBlock] $ScriptBlock,
		[Parameter()]
		[Switch] $SendArgs
	)
	$Block = {}
	if ($Method) {
		if ($SendArgs) {
			$Block = { $Method.Invoke( $Args[0], $Args[1] ) }.GetNewClosure()
		} else {
			$Block = { $Method.Invoke() }.GetNewClosure()
		}
	}
	if ($ScriptBlock) {
		if ($SendArgs) {
			$Block = { $ScriptBlock.Invoke($Args) }.GetNewClosure()
		} else {
			$Block = { $ScriptBlock.Invoke() }.GetNewClosure()
		}
	}
	return $Block
}


class Task {
	[String] $Desc
	[String] $Webpage

	Task() {
		$this.Desc = ""
		$this.Webpage = ""
	}
	
	[String] GetDesc() {
		return $this.Desc
	}
		
	[String] GetWebpage() {
		return $this.Webpage
	}
	
	[Void] SetDesc([String] $NewDesc) {
		$this.Desc = $NewDesc
	}
	
	[Void] SetWebpage([String] $NewWebpage) {
		$this.Webpage = $NewWebpage
	}
}


class Category {
	[String] $Name
	[List[Task]] $TaskList

	Category() {
		$this.Name = ""
		$this.TaskList = New-Object List[Task]
	}
	
	[String] GetName() {
		return $this.Name
	}
	
	[List[Task]] GetTaskList() {
		return $this.TaskList
	}
	
	[Void] SetName([String] $newName) {
		$this.Name = $NewName
	}
	
	[Void] SetTaskList([List[Task]] $newList) {
		$this.TaskList = $NewList
	}

	[Int] TaskCount() {
		return $this.TaskList.Count
	}

	[Void] AddTask([Task] $NewTask) {
		$this.TaskList.Add($NewTask)
	}

	[Void] RemoveTask([Task] $Task) {
		$this.TaskList.Remove($Task)
	}

	[Void] RemoveTaskAt([Int] $TaskIndex) {
		$this.TaskList.RemoveAt($TaskIndex)
	}

	[Void] ClearTasks() {
		$this.TaskList.Clear()
	}
}


class Agenda {
	<#
		CONSTANTS
	#>
	[Int] $FormHeight
	[Int] $FormWidth
	[Int] $TabControlHeight
	[Int] $TabControlWidth
	[Int] $TabControlX
	[Int] $TabControlY
	[Int] $ListViewHeight
	[Int] $ListViewWidth
	[Int] $ListViewX
	[Int] $ListViewY
	[Int] $DeleteCategoryButtonWidth
	[Int] $DeleteCategoryButtonHeight
	[Int] $SideButtonHeight
	[Int] $SideButtonWidth
	[Int] $SideButtonX
	[Int] $SideButtonY
	[Int] $SideButtonPadding
	[Int] $AddTaskTextBoxWidth
	[Int] $AddTaskTextBoxHeight
	[Int] $AddTaskTextBoxX
	[Int] $AddTaskTextBoxY
	[String] $SelectedTabWhitespace
	[String] $AddTabPageText
	[String] $SaveDataPath
	[String] $TrashIconPath
	[String] $CheckIconPath
	[String] $UncheckIconPath
	[String] $InvalidNameTitle
	[String] $InvalidNameMsg
	[String] $DuplicateNameTitle
	[String] $DuplicateNameMsg
	[String] $UnsavedTitle
	[String] $UnsavedMsg
	[String] $DeleteCatTitle
	[String] $DeleteCatMsg
	[String] $SaveTitle
	[String] $SaveMsg
	[String] $DefaultCategory
	[String] $UnnamedCategory

	<#
		VARIABLES
	#>
	[Boolean] $IsSaved
	[Boolean] $IsNew
	[List[Category]] $AgendaData

	<#
		CONTROLS
	#>
	[TabPage] $AddTabPage
	[TextBox] $AddTaskTextBox
	[TabControl] $MainTabControl
	[MenuStrip] $MainMenuStrip
	[Button] $DeleteTaskButton
	[Button] $DeleteCategoryButton
	[Button] $CheckAllButton
	[Button] $UncheckAllButton
	[Form] $MainForm
	
	Agenda() {
		<#
			CONSTANTS
		#>
		$this.FormHeight = 800
		$this.FormWidth = 600
		$this.TabControlHeight = $this.FormHeight - 100
		$this.TabControlWidth = $this.FormWidth - 85
		$this.TabControlX = 35
		$this.TabControlY = 40
		$this.ListViewHeight = $this.TabControlHeight - 60
		$this.ListViewWidth = $this.TabControlWidth
		$this.ListViewX = 0
		$this.ListViewY = 32
		$this.DeleteCategoryButtonWidth = 15
		$this.DeleteCategoryButtonHeight = 15
		$this.SideButtonHeight = 30
		$this.SideButtonWidth = 30
		$this.SideButtonX = 5
		$this.SideButtonY = 94
		$this.SideButtonPadding = $this.SideButtonHeight + 4
		$this.AddTaskTextBoxWidth = $this.TabControlWidth - 4
		$this.AddTaskTextBoxHeight = 20
		$this.AddTaskTextBoxX = $this.TabControlX + 2
		$this.AddTaskTextBoxY = $this.TabControlY + 30
		$this.SelectedTabWhitespace = " " * 5
		$this.AddTabPageText = "   +"
		$this.SaveDataPath = "$($PSScriptRoot)\Save.json"
		$this.TrashIconPath = "$($PSScriptRoot)\Images\TrashIcon.png"
		$this.CheckIconPath = "$($PSScriptRoot)\Images\CheckIcon.png"
		$this.UncheckIconPath = "$($PSScriptRoot)\Images\UncheckIcon.png"
		$this.InvalidNameTitle = "Invalid Name"
		$this.InvalidNameMsg = "Please enter a valid name."
		$this.DuplicateNameTitle = "Duplicate Name"
		$this.DuplicateNameMsg = "Name already in use."
		$this.UnsavedTitle = "Unsaved Changes"
		$this.UnsavedMsg = "Unsaved changes will be lost. Continue without saving?"
		$this.DeleteCatTitle = "Delete"
		$this.DeleteCatMsg = "Delete list?"
		$this.SaveTitle = "Save"
		$this.SaveMsg = "Save changes?"
		$this.DefaultCategory = "General"
		$this.UnnamedCategory = "New List"

		<#
			VARIABLES
		#>
		$this.IsSaved = $True
		$this.IsNew = $False
		$this.AgendaData = $this.GetData()

		<#
			CONTROLS
		#>
		$this.AddTabPage = New-Object TabPage
		$this.AddTaskTextBox = $this.SetAddTaskTextBox()
		$this.MainTabControl = $this.SetMainTabControl()
		$this.MainMenuStrip = $this.SetMainMenuStrip()
		$this.DeleteTaskButton = $this.SetDeleteTaskButton()
		$this.DeleteCategoryButton = $this.SetDeleteCategoryButton()
		$this.CheckAllButton = $this.SetCheckAllButton()
		$this.UncheckAllButton = $this.SetUncheckAllButton()
		$this.MainForm = $this.SetMainForm()
	}

	<#
		CLASS METHODS
	#>
	[Form] SetMainForm() {
		$NewForm = New-Object Form
		$NewForm.Name = "MainForm"
		$NewForm.Text = "Agenda"
		$NewForm.Size = New-Object Size($this.FormWidth, $this.FormHeight)
		$NewForm.StartPosition = [FormStartPosition]::CenterScreen
		$NewForm.add_Shown( (Add-EventWrapper -Method $this.MainForm_Shown) )
		$NewForm.Controls.Add($this.MainMenuStrip)
		$NewForm.Controls.Add($this.AddTaskTextBox)
		$NewForm.Controls.Add($this.DeleteTaskButton)
		$NewForm.Controls.Add($this.DeleteCategoryButton)
		$NewForm.Controls.Add($this.CheckAllButton)
		$NewForm.Controls.Add($this.UncheckAllButton)
		$NewForm.Controls.Add($this.MainTabControl)
		$NewForm.add_Click( (Add-EventWrapper -Method $this.BlurredControl_Click) )
		$Newform.add_FormClosing( (Add-EventWrapper -Method $this.MainForm_FormClosing -SendArgs) )
		$this.DeleteCategoryButton.BringToFront()
		return $NewForm
	}

	[TabControl] SetMainTabControl() {
		$NewTabControl = New-Object TabControl
		$NewTabControl.Location = New-Object Point($this.TabControlX, $this.TabControlY)
		$NewTabControl.Size = New-Object Size($this.TabControlWidth, $this.TabControlHeight)
		$NewTabControl.Multiline = $True
		foreach ($Category in $this.AgendaData) {
			$NewTab = New-Object TabPage
			$NewTab.Text = $Category.GetName()
			$NewListView = $this.SetListView( ($Category.GetTaskList()) )
			$NewTab.Controls.Add($NewListView)
			$NewTab.add_Click( (Add-EventWrapper -Method $this.BlurredControl_Click) )
			$NewTabControl.Controls.Add($NewTab)
		}
		$this.AddTabPage.Text = $this.AddTabPageText
		$NewTabControl.Controls.Add($this.AddTabPage)
		$NewTabControl.add_DoubleClick( (Add-EventWrapper -Method $this.MainTabControl_DoubleClick) )
		$NewTabControl.add_Deselected( (Add-EventWrapper -Method $this.MainTabControl_Deselected -SendArgs) )
		$NewTabControl.add_SelectedIndexChanged( (Add-EventWrapper -Method $this.MainTabControl_SelectedIndexChanged) )
		return $NewTabControl
	}

	[MenuStrip] SetMainMenuStrip() {
		$NewMenuStrip = New-Object MenuStrip
		$File = New-Object ToolStripMenuItem
		$Save = New-Object ToolStripMenuItem
		$Close = New-Object ToolStripMenuItem
		$File.Text = "File"
		$Save.Text = "Save"
		$Close.Text = "Close"
		$NewMenuStrip.Items.Add($File)
		$File.DropDownItems.AddRange( @($Save, $Close) )
		$Save.add_Click( (Add-EventWrapper -Method $this.SaveToolStripMenuItem_Click) )
		$Close.add_Click( (Add-EventWrapper -Method $this.CloseToolStripMenuItem_Click) )
		return $NewMenuStrip
	}

	[ListView] SetListView([List[Task]] $NewTaskList) {
		$NewListView = New-Object ListView
		$NewListView.View = "Details"
		$NewListView.HeaderStyle = "None"
		$NewListView.Columns.Add("", -1)
		$NewListView.CheckBoxes = $True
		$NewListView.AllowDrop = $True
		$NewListView.Size = New-Object Size($this.ListViewWidth, $this.ListViewHeight)
		$NewListView.Location = New-Object Point($this.ListViewX, $this.ListViewY)
		$NewListView.Font = New-Object Font("Segoe UI", 12, [FontStyle]::Regular)

		$ItemDragHandler = {
			$s = $Args[0]
			$e = $Args[1]
			$s.DoDragDrop($e.Item, [DragDropEffects]::Copy)
		}
		$DragEnterHandler = {
			$e = $Args[1]
			$e.Effect = $e.AllowedEffect
		}
		$DragDropHandler = {
			$s = $Args[0]
			$e = $Args[1]
			if ($e.Data.GetDataPresent([ListViewItem])) {
				foreach ($TaskDesc in $e.Data.GetData([ListViewItem])) {
					$s.Items.Add($TaskDesc.Text)
				}
			}
		}
		$NewListView.add_ItemDrag( $ItemDragHandler )
		$NewListView.add_DragEnter( $DragEnterHandler )
		$NewListView.add_DragDrop( $DragDropHandler )

		foreach ($Task in $NewTaskList) {
			$NewListView.Items.Add( ($Task.GetDesc()) )
		}
		return $NewListView
	}

	[TextBox] SetAddTaskTextBox() {
		$NewTextBox = New-Object TextBox
		$NewTextBox.Location = New-Object Point( $this.AddTaskTextBoxX, $this.AddTaskTextBoxY )
		$NewTextBox.Size = New-Object Size( $this.AddTaskTextBoxWidth, $this.AddTaskTextBoxHeight )
		$NewTextBox.ContextMenu = New-Object ContextMenu
		$NewTextBox.add_KeyDown( (Add-EventWrapper -Method $this.AddTaskTextBox_KeyDown -SendArgs) )
		return $NewTextBox
	}

	[Button] SetDeleteCategoryButton() {
		$NewBtn = New-Object Button
		$NewBtn.Size = New-Object Size($this.DeleteCategoryButtonWidth, $this.DeleteCategoryButtonHeight)
		$NewBtn.Text = "x"
		$NewBtn.add_Click( (Add-EventWrapper -Method $this.DeleteCategoryButton_Click) )
		return $NewBtn
	}

	[Button] SetDeleteTaskButton() {
		$NewBtn = New-Object Button
		$NewBtn.Size = New-Object Size($this.SideButtonWidth, $this.SideButtonHeight)
		$NewBtn.Location = New-Object Point($this.SideButtonX, $this.SideButtonY)
		$NewImage = [Image]::FromFile($this.TrashIconPath)
		$NewBtn.ImageList = New-Object ImageList
		$NewBtn.ImageList.ImageSize = New-Object Size( ($this.SideButtonWidth - 5), ($this.SideButtonHeight - 5) )
		$NewBtn.ImageList.Images.Add($NewImage)
		$NewBtn.ImageIndex = 0
		$NewBtn.add_Click( (Add-EventWrapper -Method $this.DeleteTaskButton_Click) )
		return $NewBtn
	}

	[Button] SetCheckAllButton() {
		$NewBtn = New-Object Button
		$NewBtn.Size = New-Object Size($this.SideButtonWidth, $this.SideButtonHeight)
		$NewBtn.Location = New-Object Point($this.SideButtonX, ($this.SideButtonY + $this.SideButtonPadding) )
		$NewImage = [Image]::FromFile($this.CheckIconPath)
		$NewBtn.ImageList = New-Object ImageList
		$NewBtn.ImageList.ImageSize = New-Object Size( ($this.SideButtonWidth - 10), ($this.SideButtonHeight - 10) )
		$NewBtn.ImageList.Images.Add($NewImage)
		$NewBtn.ImageIndex = 0
		$NewBtn.add_Click( (Add-EventWrapper -Method $this.CheckButton_Click) )
		return $NewBtn
	}

	[Button] SetUncheckAllButton() {
		$NewBtn = New-Object Button
		$NewBtn.Size = New-Object Size($this.SideButtonWidth, $this.SideButtonHeight)
		$NewBtn.Location = New-Object Point($this.SideButtonX, ($this.SideButtonY + $this.SideButtonPadding * 2) )
		$NewImage = [Image]::FromFile($this.UncheckIconPath)
		$NewBtn.ImageList = New-Object ImageList
		$NewBtn.ImageList.ImageSize = New-Object Size( ($this.SideButtonWidth - 10), ($this.SideButtonHeight - 10) )
		$NewBtn.ImageList.Images.Add($NewImage)
		$NewBtn.ImageIndex = 0
		$NewBtn.add_Click( (Add-EventWrapper -Method $this.UncheckButton_Click) )
		return $NewBtn
	}

	[List[Category]] GetData() {
		$FileExists = Test-Path -Path $this.SaveDataPath
		If ($FileExists -eq $False) {
			$NewCategory = New-Object Category
			$NewCategory.SetName($this.DefaultCategory)
			$NewData = [List[Category]]::New()
			$NewData.Add($NewCategory)
			$NewData | ConvertTo-Json -Depth 3 | Set-Content -Path $this.SaveDataPath
		}
		$RawJSON = Get-Content -Path $this.SaveDataPath -Raw | ConvertFrom-Json
		$NewData = [List[Category]]::New()
		foreach ($Object in $RawJSON) {
			try {
				$NewCategory = [Category] $Object
			}
			catch {
				$NewCategory = New-Object Category
				$NewCategory.SetName($Object.Name)
				foreach ($Task in $Object.TaskList) {
					$Task = [Task] $Task
					$NewCategory.AddTask($Task)
				}
			}
			finally {
				$NewData.Add($NewCategory)
			}
		}
		return $NewData
	}

	[Void] RelocateDeleteCategoryButton() {
		$CurrentIndex = $this.MainTabControl.SelectedIndex
		if ($CurrentIndex -eq -1) {
			$CurrentIndex = 0
		}
		$Rect = $this.MainTabControl.GetTabRect($CurrentIndex)
		$BtnX = $Rect.Right - $this.DeleteCategoryButtonWidth + $this.TabControlX
		$BtnY = $Rect.Y + $this.TabControlY
		$this.DeleteCategoryButton.Location = New-Object Point($BtnX, $BtnY)
	}
	
	[Void] Open() {
		$this.MainForm.ShowDialog() | Out-Null
		$this.MainForm.Dispose()
	}

	<#
		EVENT HANDLERS
	#>
	[Void] ListView_DragDrop([Object] $s, [EventArgs] $e) {
		# Defunct method
		if ($e.Data.GetDataPresent([DataFormats]::Text)) {
			$e.Effect = DragDropEffects.Copy
		} else {
			$e.Effect = DragDropEffects.None
		}
	}

	[Void] ListView_DragEnter() {
		# Defunct method
		$s = $Args[0]
		$e = $Args[1]
		if ($e.Data.GetDataPresent([DataFormats]::Text)) {
			foreach ($TaskDesc in $e.Data.GetData([DataFormats]::Text)) {
				$s.Items.Add($TaskDesc)
			}
		}
	}

	[Void] BlurredControl_Click() {
		$this.MainTabControl.SelectedTab.Select()
	}

	[Void] MainForm_Shown() {
		if ($this.MainTabControl.Controls.Count -gt 2) {
			foreach ($CurrentTab in $this.MainTabControl.Controls) {
				if ($CurrentTab -ne $this.AddTabPage) {
					$this.MainTabControl.SelectedTab = $CurrentTab
				}
			}
			$this.MainTabControl.SelectedIndex = 0
		} else {
			$this.MainTabControl.Controls[0].Text += $this.SelectedTabWhitespace
			$this.RelocateDeleteCategoryButton()
		}
	}

	[Void] MainForm_FormClosing([Object] $s, [EventArgs] $e) {
		if ($this.IsSaved -eq $False) {
			$MsgBtns = [MessageBoxButtons]::YesNo
			$Result = [MessageBox]::Show($this.UnsavedMsg, $this.UnsavedTitle, $MsgBtns)
			if ($Result -eq [DialogResult]::No) {
				$e.Cancel = $True
			}
		}
	}

	[Void] MainTabControl_DoubleClick() {
		$this.DeleteCategoryButton.Visible = $False
		$CurrentTab = $this.MainTabControl.SelectedTab
		$CurrentIndex = $this.MainTabControl.SelectedIndex
		$Rect = $this.MainTabControl.GetTabRect($CurrentIndex)
		$RenameTextBox = New-Object TextBox
		$RenameTextBox.Location = New-Object Point( ($Rect.X + $this.TabControlX), ($Rect.Y + $this.TabControlY) )
		$RenameTextBox.Size = New-Object Size($Rect.Size)
		$RenameTextBox.Text = $CurrentTab.Text.Trim()
		$RenameTextBox.add_KeyDown( (Add-EventWrapper -Method $this.RenameTextBox_KeyDown -SendArgs) )
		$RenameTextBox.add_Leave( (Add-EventWrapper -Method $this.RenameTextBox_Leave -SendArgs) )
		$this.MainForm.Controls.Add($RenameTextBox)
		$RenameTextBox.BringToFront()
		$RenameTextBox.Focus()
		$RenameTextBox.SelectAll()
	}

	[Void] MainTabControl_Deselected([Object] $s, [EventArgs] $e) {
		$PrevTab = $e.TabPage
		if ($PrevTab.Disposing -eq $False -and $PrevTab -ne $this.AddTabPage) {
			$PrevTab.Text = $PrevTab.Text.Trim()
		}
	}

	[Void] MainTabControl_SelectedIndexChanged() {
		$CurrentTab = $this.MainTabControl.SelectedTab
		if ($CurrentTab -eq $this.AddTabPage) {
			$this.DeleteCategoryButton.Visible = $False
			$this.IsNew = $True
			$NewTabPage = New-Object TabPage
			$NewTabPage.Text = $this.UnnamedCategory
			$NewListView = $this.SetListView( @() )
			$NewTabPage.Controls.Add($NewListView)
			$NewTabPage.add_Click( (Add-EventWrapper -Method $this.BlurredControl_Click) )
			$this.MainTabControl.TabPages.Insert( ($this.MainTabControl.Controls.Count - 1), $NewTabPage )
			$this.MainTabControl.SelectedTab = $NewTabPage
			$CurrentIndex = $this.MainTabControl.SelectedIndex
			$Rect = $this.MainTabControl.GetTabRect($CurrentIndex)
			$RenameTextBox = New-Object TextBox
			$RenameTextBox.Location = New-Object Point( ($Rect.X + $this.TabControlX), ($Rect.Y + $this.TabControlY) )
			$RenameTextBox.Size = New-Object Size($Rect.Size)
			$RenameTextBox.Text = $NewTabPage.Text.Trim()
			$RenameTextBox.add_KeyDown( (Add-EventWrapper -Method $this.RenameTextBox_KeyDown -SendArgs) )
			$RenameTextBox.add_Leave( (Add-EventWrapper -Method $this.RenameTextBox_Leave -SendArgs) )
			$this.MainForm.Controls.Add($RenameTextBox)
			$RenameTextBox.BringToFront()
			$RenameTextBox.Focus()
			$RenameTextBox.SelectAll()
		} else {
			$CurrentTab.Text = $CurrentTab.Text + $this.SelectedTabWhitespace
			$this.RelocateDeleteCategoryButton()
		}
	}

	[Void] CloseToolStripMenuItem_Click() {
		$this.MainForm.Close()
	}

	[Void] SaveToolStripMenuItem_Click() {
		$MsgBtns = [MessageBoxButtons]::YesNo
		$Result = [MessageBox]::Show($this.SaveMsg, $this.SaveTitle, $MsgBtns)
		if ($Result -eq [DialogResult]::Yes) {
			$this.AgendaData | ConvertTo-Json -Depth 3 | Set-Content -Path $this.SaveDataPath
		}
		$this.IsSaved = $True
	}

	[Void] DeleteTaskButton_Click() {
		$CurrentTab = $this.MainTabControl.SelectedTab
		$CurrentIndex = $this.MainTabControl.SelectedIndex
		$CurrentListView = $CurrentTab.Controls[0]
		$Checked = $CurrentListView.CheckedIndices
		for ($i = $Checked.Count - 1; $i -ge 0; $i--) {
			$this.AgendaData[$CurrentIndex].RemoveTaskAt($Checked[$i])
			$CurrentListView.Items.RemoveAt($Checked[$i])
			$this.IsSaved = $False
		}
	}

	[Void] CheckButton_Click() {
		$CurrentTab = $this.MainTabControl.SelectedTab
		$CurrentListView = $CurrentTab.Controls[0]
		foreach ($Task in $CurrentListView.Items) {
			$Task.Checked = $True
		}
	}

	[Void] UncheckButton_Click() {
		$CurrentTab = $this.MainTabControl.SelectedTab
		$CurrentListView = $CurrentTab.Controls[0]
		foreach ($Task in $CurrentListView.Items) {
			$Task.Checked = $False
		}
	}

	[Void] AddTaskTextBox_KeyDown([Object] $s, [EventArgs] $e) {	
		if ($e.KeyCode -eq "Enter") {
			$CurrentTab = $this.MainTabControl.SelectedTab
			$TaskText = $this.AddTaskTextBox.Text
			$CurrentListView = $CurrentTab.Controls[0]
			$CurrentListView.Items.Add($TaskText)
			$CurrentListView.AutoResizeColumn(0, "ColumnContent")
			$this.AddTaskTextBox.Clear()
			$e.SuppressKeyPress = $True
			$e.Handled = $True
			foreach ($Category in $this.AgendaData) {
				if ( ($CurrentTab.Text.Trim()) -eq ($Category.GetName()) ) {
					$NewTask = New-Object Task
					$NewTask.SetDesc($TaskText)
					$Category.AddTask($NewTask)
					$this.IsSaved = $False
					Break
				}
			}
		}
	}

	[Void] DeleteCategoryButton_Click() {
		$CurrentIndex = $this.MainTabControl.SelectedIndex
		$MsgBtns = [MessageBoxButtons]::YesNo
		$Result = [MessageBox]::Show($this.DeleteCatMsg, $this.DeleteCatTitle, $MsgBtns)
		if ($Result -eq [DialogResult]::Yes) {
			if ($this.MainTabControl.Controls.Count -eq 2) {
				$CurrentTab = $this.MainTabControl.SelectedTab
				$CurrentTab.Text = $this.DefaultCategory + $this.SelectedTabWhitespace
				$CurrentTab.Controls[0].Clear()
				$this.RelocateDeleteCategoryButton()
				$this.AddTaskTextBox.Select()
				$this.AgendaData[0].SetName($this.DefaultCategory)
				$this.AgendaData[0].ClearTasks()
			} else {
				if ($CurrentIndex -gt 0) {
					$this.MainTabControl.SelectedIndex = ($CurrentIndex - 1)
				}
				$this.MainTabControl.Controls.RemoveAt($CurrentIndex)
				$this.AgendaData.RemoveAt($CurrentIndex)
			}
			$this.IsSaved = $False
		}
	}

	[Void] RenameTextBox_KeyDown([Object] $s, [EventArgs] $e) {
		if ($e.KeyCode -ne "Enter") {
			return
		}
		$e.Handled = $True
		$e.SuppressKeyPress = $True
		$NewCategoryName = $s.Text.Trim()
		if ($NewCategoryName -eq "") {
			$MsgBtns = [MessageBoxButtons]::OK
			[MessageBox]::Show($this.InvalidNameMsg, $this.InvalidNameTitle, $MsgBtns)
			return
		}
		foreach ($Category in $this.AgendaData) {
			if ( ($Category.GetName()) -eq $NewCategoryName) {
				$MsgBtns = [MessageBoxButtons]::OK
				[MessageBox]::Show($this.DuplicateNameMsg, $this.DuplicateNameTitle, $MsgBtns)
				return
			}
		}
		if ($this.IsNew -eq $True) {
			$NewCategory = New-Object Category
			$this.AgendaData.Add($NewCategory)
			$this.IsNew = $False
		}
		$s.Dispose()
		$CurrentIndex = $this.MainTabControl.SelectedIndex
		$this.AgendaData[$CurrentIndex].SetName($NewCategoryName)
		$this.MainTabControl.SelectedTab.Text = $NewCategoryName + $this.SelectedTabWhitespace
		$this.RelocateDeleteCategoryButton()
		$this.DeleteCategoryButton.Visible = $True
		$this.IsSaved = $False
	}

	[Void] RenameTextBox_Leave([Object] $s, [EventArgs] $e) {
		if ($s.Disposing -eq $True) {
			return
		}
		if ($this.IsNew -eq $True) {
			$CurrentTab = $this.MainTabControl.SelectedTab
			$DefaultName = $CurrentTab.Text.Trim()
			$NewCategoryName = -Split $DefaultName
			$MaxUnnamedCount = 0
			foreach ($Category in $this.AgendaData) {
				$CategoryName = -Split ( $Category.GetName() )
				if ($CategoryName.Count -lt 2 -or $CategoryName.Count -gt 3) {
					Continue
				}
				if ($CategoryName[0] -eq $NewCategoryName[0] -and $CategoryName[1] -eq $NewCategoryName[1]) {
					if ($CategoryName.Count -eq 3) {
						$Num = $CategoryName[2] -as [Int]
						if ($Null -eq $Num) {
							Continue
						}
						if ($Num -ge $MaxUnnamedCount) {
							$MaxUnnamedCount = $Num + 1
						}
					} else {
						if ($MaxUnnamedCount -eq 0) {
							$MaxUnnamedCount = 1
						}
					}
				}
			}
			if ($MaxUnnamedCount-gt 0) {
				$DefaultName = "$($DefaultName) $($MaxUnnamedCount)"
				$CurrentTab.Text = $DefaultName + $this.SelectedTabWhitespace
			}
			$this.RelocateDeleteCategoryButton()
			$NewCategory = New-Object Category
			$NewCategory.SetName($DefaultName)
			$this.AgendaData.Add($NewCategory)
			$this.IsSaved = $False
			$this.IsNew = $False
		}
		$s.Dispose()
		$this.DeleteCategoryButton.Visible = $True
	}
}


# Set-PSDebug -Trace 2
Open-Agenda