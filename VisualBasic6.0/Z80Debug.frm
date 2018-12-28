VERSION 5.00
Begin VB.Form frmDlgFileOpen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4020
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7290
   Icon            =   "Z80Debug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboFileType 
      Height          =   315
      ItemData        =   "Z80Debug.frx":0442
      Left            =   1275
      List            =   "Z80Debug.frx":0444
      TabIndex        =   2
      Top             =   3465
      Width           =   2790
   End
   Begin VB.TextBox txtFilePath 
      Height          =   300
      Left            =   180
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   165
      Width           =   6930
   End
   Begin VB.FileListBox lstFileList 
      Height          =   2625
      Left            =   3315
      TabIndex        =   5
      Top             =   615
      Width           =   3810
   End
   Begin VB.DriveListBox lstDrvList 
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   630
      Width           =   2955
   End
   Begin VB.DirListBox lstDirList 
      Height          =   2115
      Left            =   180
      TabIndex        =   4
      Top             =   1110
      Width           =   2970
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5895
      TabIndex        =   1
      Top             =   3450
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4365
      TabIndex        =   0
      Top             =   3450
      Width           =   1215
   End
   Begin VB.Label lblFileType 
      Caption         =   "Files of type :"
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   3525
      Width           =   990
   End
End
Attribute VB_Name = "frmDlgFileOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim curListIndex As Integer, curDrv As String, flExitMode As Boolean, flName As String

Property Get ExitMode() As Boolean
    ExitMode = flExitMode
End Property

Property Get FileName() As String
    FileName = flName
End Property

Private Sub btnCancel_Click()
    Hide
End Sub

Private Sub btnOK_Click()
    flExitMode = True
    Hide
End Sub

Private Sub cboFileType_Change()
    cboFileType.ListIndex = curListIndex
End Sub

Private Sub cboFileType_Click()
    Dim beginPos As Long, endPos As Long
    curListIndex = cboFileType.ListIndex
    beginPos = InStr(cboFileType.Text, "(") + 1
    endPos = InStr(cboFileType.Text, ")")
    lstFileList.Pattern = Mid(cboFileType.Text, beginPos, endPos - beginPos)
End Sub

Private Sub Form_Load()
    Caption = App.Title & " - Open Source File to Assembling..."
    cboFileType.AddItem "Z80 Binary Files Type (*." & cSrcFileExt & ")"
    cboFileType.AddItem "All Files (*.*)"
    cboFileType.ListIndex = 0
    cboFileType_Click
    btnOK.Enabled = False
    curDrv = lstDrvList.Drive
End Sub

Private Sub lstDirList_Change()
    lstFileList.Path = lstDirList.Path
    ChDir lstDirList.Path
    txtFilePath.Text = ""
End Sub

Private Sub lstDrvList_Change()
    On Error Resume Next
    lstDirList.Path = lstDrvList.Drive
    If Err.Number > 0 Then
        MsgBox Err.Description, vbCritical
        Err.Clear
        lstDrvList.Drive = curDrv
    Else
        curDrv = lstDrvList.Drive
    End If
    ChDrive lstDrvList.Drive
End Sub

Private Sub lstFileList_Click()
    txtFilePath.Text = lstFileList.Path & "\" & lstFileList.FileName
    btnOK.Enabled = True
End Sub

Private Sub lstFileList_DblClick()
    btnOK_Click
End Sub

Private Sub lstFileList_PathChange()
    btnOK.Enabled = False
End Sub

Private Sub txtFilePath_Change()
    flName = lstFileList.FileName
End Sub

Private Sub txtFilePath_GotFocus()
    txtFilePath.Enabled = False
End Sub

Private Sub txtFilePath_LostFocus()
    txtFilePath.Enabled = True
End Sub
