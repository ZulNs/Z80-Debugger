<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDlgFileOpen
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cboFileType As System.Windows.Forms.ComboBox
	Public WithEvents txtFilePath As System.Windows.Forms.TextBox
	Public WithEvents lstFileList As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
	Public WithEvents lstDrvList As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
	Public WithEvents lstDirList As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents btnOK As System.Windows.Forms.Button
	Public WithEvents lblFileType As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDlgFileOpen))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cboFileType = New System.Windows.Forms.ComboBox
		Me.txtFilePath = New System.Windows.Forms.TextBox
		Me.lstFileList = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox
		Me.lstDrvList = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
		Me.lstDirList = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
		Me.btnCancel = New System.Windows.Forms.Button
		Me.btnOK = New System.Windows.Forms.Button
		Me.lblFileType = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(486, 268)
		Me.Location = New System.Drawing.Point(184, 250)
		Me.Icon = CType(resources.GetObject("frmDlgFileOpen.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmDlgFileOpen"
		Me.cboFileType.Size = New System.Drawing.Size(186, 21)
		Me.cboFileType.Location = New System.Drawing.Point(85, 231)
		Me.cboFileType.TabIndex = 2
		Me.cboFileType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboFileType.BackColor = System.Drawing.SystemColors.Window
		Me.cboFileType.CausesValidation = True
		Me.cboFileType.Enabled = True
		Me.cboFileType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboFileType.IntegralHeight = True
		Me.cboFileType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboFileType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboFileType.Sorted = False
		Me.cboFileType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboFileType.TabStop = True
		Me.cboFileType.Visible = True
		Me.cboFileType.Name = "cboFileType"
		Me.txtFilePath.AutoSize = False
		Me.txtFilePath.Size = New System.Drawing.Size(462, 20)
		Me.txtFilePath.Location = New System.Drawing.Point(12, 11)
		Me.txtFilePath.TabIndex = 6
		Me.txtFilePath.TabStop = False
		Me.txtFilePath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFilePath.AcceptsReturn = True
		Me.txtFilePath.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFilePath.BackColor = System.Drawing.SystemColors.Window
		Me.txtFilePath.CausesValidation = True
		Me.txtFilePath.Enabled = True
		Me.txtFilePath.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFilePath.HideSelection = True
		Me.txtFilePath.ReadOnly = False
		Me.txtFilePath.Maxlength = 0
		Me.txtFilePath.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFilePath.MultiLine = False
		Me.txtFilePath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFilePath.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFilePath.Visible = True
		Me.txtFilePath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFilePath.Name = "txtFilePath"
		Me.lstFileList.Size = New System.Drawing.Size(254, 175)
		Me.lstFileList.Location = New System.Drawing.Point(221, 41)
		Me.lstFileList.TabIndex = 5
		Me.lstFileList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstFileList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstFileList.Archive = True
		Me.lstFileList.BackColor = System.Drawing.SystemColors.Window
		Me.lstFileList.CausesValidation = True
		Me.lstFileList.Enabled = True
		Me.lstFileList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstFileList.Hidden = False
		Me.lstFileList.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstFileList.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstFileList.Normal = True
		Me.lstFileList.Pattern = "*.*"
		Me.lstFileList.ReadOnly = True
		Me.lstFileList.System = False
		Me.lstFileList.TabStop = True
		Me.lstFileList.TopIndex = 0
		Me.lstFileList.Visible = True
		Me.lstFileList.Name = "lstFileList"
		Me.lstDrvList.Size = New System.Drawing.Size(197, 21)
		Me.lstDrvList.Location = New System.Drawing.Point(12, 42)
		Me.lstDrvList.TabIndex = 3
		Me.lstDrvList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstDrvList.BackColor = System.Drawing.SystemColors.Window
		Me.lstDrvList.CausesValidation = True
		Me.lstDrvList.Enabled = True
		Me.lstDrvList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstDrvList.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstDrvList.TabStop = True
		Me.lstDrvList.Visible = True
		Me.lstDrvList.Name = "lstDrvList"
		Me.lstDirList.Size = New System.Drawing.Size(198, 141)
		Me.lstDirList.Location = New System.Drawing.Point(12, 74)
		Me.lstDirList.TabIndex = 4
		Me.lstDirList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstDirList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstDirList.BackColor = System.Drawing.SystemColors.Window
		Me.lstDirList.CausesValidation = True
		Me.lstDirList.Enabled = True
		Me.lstDirList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstDirList.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstDirList.TabStop = True
		Me.lstDirList.Visible = True
		Me.lstDirList.Name = "lstDirList"
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCancel.Text = "&Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(81, 25)
		Me.btnCancel.Location = New System.Drawing.Point(393, 230)
		Me.btnCancel.TabIndex = 1
		Me.btnCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
		Me.btnCancel.CausesValidation = True
		Me.btnCancel.Enabled = True
		Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCancel.TabStop = True
		Me.btnCancel.Name = "btnCancel"
		Me.btnOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnOK.Text = "&OK"
		Me.btnOK.Size = New System.Drawing.Size(81, 25)
		Me.btnOK.Location = New System.Drawing.Point(291, 230)
		Me.btnOK.TabIndex = 0
		Me.btnOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnOK.BackColor = System.Drawing.SystemColors.Control
		Me.btnOK.CausesValidation = True
		Me.btnOK.Enabled = True
		Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnOK.TabStop = True
		Me.btnOK.Name = "btnOK"
		Me.lblFileType.Text = "Files of type :"
		Me.lblFileType.Size = New System.Drawing.Size(66, 16)
		Me.lblFileType.Location = New System.Drawing.Point(12, 235)
		Me.lblFileType.TabIndex = 7
		Me.lblFileType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFileType.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFileType.BackColor = System.Drawing.SystemColors.Control
		Me.lblFileType.Enabled = True
		Me.lblFileType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFileType.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFileType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFileType.UseMnemonic = True
		Me.lblFileType.Visible = True
		Me.lblFileType.AutoSize = False
		Me.lblFileType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFileType.Name = "lblFileType"
		Me.Controls.Add(cboFileType)
		Me.Controls.Add(txtFilePath)
		Me.Controls.Add(lstFileList)
		Me.Controls.Add(lstDrvList)
		Me.Controls.Add(lstDirList)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(btnOK)
		Me.Controls.Add(lblFileType)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class