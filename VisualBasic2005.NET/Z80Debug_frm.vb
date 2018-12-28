Option Strict Off
Option Explicit On
Friend Class frmDlgFileOpen
	Inherits System.Windows.Forms.Form
	
	
	Dim curListIndex As Short
	Dim curDrv, flName As String
	Dim flExitMode As Boolean
	
	ReadOnly Property ExitMode() As Boolean
		Get
			ExitMode = flExitMode
		End Get
	End Property
	
	ReadOnly Property FileName() As String
		Get
			FileName = flName
		End Get
	End Property
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		Hide()
	End Sub
	
	Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click
		flExitMode = True
		Hide()
	End Sub
	
	'UPGRADE_WARNING: Event cboFileType.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cboFileType.Change was upgraded to cboFileType.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cboFileType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFileType.TextChanged
		cboFileType.SelectedIndex = curListIndex
	End Sub
	
	'UPGRADE_WARNING: Event cboFileType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboFileType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFileType.SelectedIndexChanged
		Dim beginPos, endPos As Integer
		curListIndex = cboFileType.SelectedIndex
		beginPos = InStr(cboFileType.Text, "(") + 1
		endPos = InStr(cboFileType.Text, ")")
		lstFileList.Pattern = Mid(cboFileType.Text, beginPos, endPos - beginPos)
	End Sub
	
	Private Sub frmDlgFileOpen_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Text = My.Application.Info.Title & " - Open Source File to Assembling..."
		cboFileType.Items.Add("Z80 Binary Files Type (*." & cSrcFileExt & ")")
		cboFileType.Items.Add("All Files (*.*)")
		cboFileType.SelectedIndex = 0
		cboFileType_SelectedIndexChanged(cboFileType, New System.EventArgs())
		btnOK.Enabled = False
		curDrv = lstDrvList.Drive
	End Sub
	
	Private Sub lstDirList_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstDirList.Change
		lstFileList.Path = lstDirList.Path
		ChDir(lstDirList.Path)
		txtFilePath.Text = ""
	End Sub
	
	Private Sub lstDrvList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstDrvList.SelectedIndexChanged
		On Error Resume Next
		lstDirList.Path = lstDrvList.Drive
		If Err.Number > 0 Then
			MsgBox(Err.Description, MsgBoxStyle.Critical)
			Err.Clear()
			lstDrvList.Drive = curDrv
		Else
			curDrv = lstDrvList.Drive
		End If
		ChDrive(lstDrvList.Drive)
	End Sub
	
	Private Sub lstFileList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstFileList.SelectedIndexChanged
		txtFilePath.Text = lstFileList.Path & "\" & lstFileList.FileName
		btnOK.Enabled = True
	End Sub
	
	Private Sub lstFileList_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstFileList.DoubleClick
		btnOK_Click(btnOK, New System.EventArgs())
	End Sub
	
	Private Sub lstFileList_PathChange(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstFileList.PathChange
		btnOK.Enabled = False
	End Sub
	
	'UPGRADE_WARNING: Event txtFilePath.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtFilePath_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFilePath.TextChanged
		flName = lstFileList.FileName
	End Sub
	
	Private Sub txtFilePath_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFilePath.Enter
		txtFilePath.Enabled = False
	End Sub
	
	Private Sub txtFilePath_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFilePath.Leave
		txtFilePath.Enabled = True
	End Sub
End Class