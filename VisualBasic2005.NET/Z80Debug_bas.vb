Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module modMain
	
	'=========================================================================================='
	'                                                                                          '
	'              Z80 Debugger for ED-Laboratory's Microprocessor Trainer MPT-1              '
	'                                                                                          '
	'                Copyright (C) Iskandar Z. Nasibu, Gorontalo, February 2005                '
	'                                                                                          '
	'=========================================================================================='
	
	
	
	Public Const cSrcFileExt As String = "BIN"
	Const cOutFileExt As String = "Z80"
	Const cSrcFileStartStr As String = "<Z80_Executable_Codes>"
	Const cSrcFileEndStr As String = "<ZulNs#05-11-1970#Viva_New_Technology_Protocol#Gorontalo#Feb-2005>"
	
	Dim codes() As Byte
	Dim pC As Integer
	Dim srcFile, outFile As String
	
	'UPGRADE_WARNING: Application will terminate when Sub Main() finishes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="E08DDC71-66BA-424F-A612-80AF11498FF8"'
	Public Sub Main()
		Dim ln, ptr, i As Integer
		Dim op1, inst, op2 As String
		On Error GoTo ErrorHandler
		If Not ReadSrcFile Then
			If srcFile = "" Then
				Exit Sub
			Else
				GoTo ErrReadSrcFile
			End If
		End If
		DelFile(outFile)
		FileOpen(1, outFile, OpenMode.Output)
		PrintLine(1, GetHorLine(44))
		PrintLine(1, "ADDRESS MACHINE-CODE  #   OPCODE  OPERAND")
		PrintLine(1, GetHorLine(44))
		PrintLine(1)
		Do While ptr < UBound(codes)
			ln = codes(ptr) + 256 * codes(ptr + 1)
			If ln = 0 Then Exit Do
			If ptr Then PrintLine(1)
			pC = codes(ptr + 2) + 256 * codes(ptr + 3)
			ptr = ptr + 4
			ln = ln + ptr
			Do While ptr < ln
				Print(1, GetHex4(pC) & ":", TAB(9))
				i = ptr
				op1 = ""
				op2 = ""
				Select Case codes(ptr)
					Case 0 To &H3Fs
						GetInst_00_3F(ptr, inst, op1, op2)
					Case &H40s To &HBFs
						GetInst_40_BF(ptr, inst, op1, op2)
					Case &HCBs
						IncPC()
						ptr = ptr + 1
						GetExtInst_CB(ptr, inst, op1, op2)
					Case &HDDs, &HFDs
						IncPC()
						ptr = ptr + 1
						GetExtInst_DD_FD(ptr, inst, op1, op2)
					Case &HEDs
						IncPC()
						ptr = ptr + 1
						GetExtInst_ED(ptr, inst, op1, op2)
					Case Else
						GetInst_C0_FF(ptr, inst, op1, op2)
				End Select
				For i = i To ptr - 1
					Print(1, GetHexByte(codes(i)) & " ")
				Next 
				Print(1, TAB(23), "#   " & inst)
				If op1 <> "" Then
					Print(1, TAB(35), op1)
					If op2 <> "" Then Print(1, "," & op2)
				End If
				PrintLine(1, "")
			Loop 
		Loop 
		FileClose(1)
		MsgBox("Debugging process successful.", MsgBoxStyle.Information)
		Shell("Notepad.exe " & outFile, AppWinStyle.NormalFocus)
		Exit Sub
ErrorHandler: 
		FileClose(1)
		DelFile(outFile)
ErrReadSrcFile: 
		MsgBox(srcFile & " is not a valid Z80 executable file." & vbCr & "Debugging process aborted.", MsgBoxStyle.Critical)
	End Sub
	
	Private Function GetHorLine(ByRef ChrNum As Integer) As String
		For ChrNum = 1 To ChrNum
			GetHorLine = GetHorLine & "="
		Next 
	End Function
	
	Private Function GetHex4(ByRef num As Integer) As String
		GetHex4 = Hex(num)
		Do While Len(GetHex4) < 4
			GetHex4 = "0" & GetHex4
		Loop 
	End Function
	
	Private Function ReadSrcFile() As Boolean
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim flPtr, codesPtr As Integer
		Dim str_Renamed As String
		Dim readOK As Boolean
		If Not GetSrcFileName Then Exit Function
		FileOpen(1, srcFile, OpenMode.Binary, OpenAccess.Read)
		str_Renamed = cSrcFileStartStr
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(1, str_Renamed)
		If str_Renamed <> cSrcFileStartStr Then GoTo EndReadSrcFile
		codesPtr = -1
		Do While Not EOF(1)
			flPtr = Seek(1)
			str_Renamed = cSrcFileEndStr
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(1, str_Renamed)
			If str_Renamed = cSrcFileEndStr Then
				readOK = True
				Exit Do
			End If
			Seek(1, flPtr)
			codesPtr = codesPtr + 1
			ReDim Preserve codes(codesPtr)
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(1, codes(codesPtr))
		Loop 
		If readOK Then
			ReadSrcFile = True
		Else
			ReDim Preserve codes(0)
		End If
EndReadSrcFile: 
		FileClose(1)
	End Function
	
	Private Function GetSrcFileName() As Boolean
		Dim CmdTail, Path As String
		Dim fso As Object
		Dim UserRespons As MsgBoxResult
		CmdTail = VB.Command()
		If CmdTail = "" Then
			If Not GetSrcFileNameFromDlg Then Exit Function
		Else
			fso = CreateObject("Scripting.FileSystemObject")
			'UPGRADE_WARNING: Couldn't resolve default property of object fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetExtensionName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not fso.FileExists(CmdTail) Then If fso.GetExtensionName(CmdTail) = "" Then CmdTail = CmdTail & "." & cSrcFileExt
			'UPGRADE_WARNING: Couldn't resolve default property of object fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If fso.FileExists(CmdTail) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetParentFolderName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Path = fso.GetParentFolderName(CmdTail)
				If Path = "" Then
					srcFile = CmdTail
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetDriveName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ChDrive(fso.GetDriveName(Path))
					ChDir(Path)
					'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetFileName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					srcFile = fso.GetFileName(CmdTail)
				End If
				'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				fso = Nothing
			Else
				'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				fso = Nothing
				UserRespons = MsgBox("Can't found '" & CmdTail & "' file or '" & CmdTail & "' is not a legal file name." & vbCr & "Try to find it or another file by your self?", MsgBoxStyle.Question + MsgBoxStyle.OKCancel)
				If UserRespons = MsgBoxResult.OK Then
					If Not GetSrcFileNameFromDlg Then Exit Function
				Else
					Exit Function
				End If
			End If
		End If
		If UCase(GetFileExt(srcFile)) = cOutFileExt Then outFile = srcFile Else outFile = GetFileName(srcFile)
		outFile = outFile & "." & cOutFileExt
		GetSrcFileName = True
	End Function
	
	Private Function GetSrcFileNameFromDlg() As Boolean
		Dim dlg As New frmDlgFileOpen
		Dim blnExit As Boolean
		dlg.ShowDialog()
		blnExit = dlg.ExitMode
		If blnExit Then srcFile = dlg.FileName
		dlg.Close()
		'UPGRADE_NOTE: Object dlg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		dlg = Nothing
		If blnExit Then GetSrcFileNameFromDlg = True Else MsgBox("No file selected. Debugging process aborted.", MsgBoxStyle.Information)
	End Function
	
	Private Function IsFileExist(ByRef FileName As String) As Boolean
		Dim fso As Object
		fso = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IsFileExist = fso.FileExists(FileName)
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
	End Function
	
	Private Function DelFile(ByRef FileName As String) As Object
		Dim fso As Object
		If IsFileExist(FileName) Then
			fso = CreateObject("Scripting.FileSystemObject")
			'UPGRADE_WARNING: Couldn't resolve default property of object fso.DeleteFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fso.DeleteFile(FileName, True)
			'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			fso = Nothing
		End If
	End Function
	
	Private Function GetFileName(ByRef FullName As String) As String
		Dim fso As Object
		fso = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetBaseName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetFileName = fso.GetBaseName(FullName)
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
	End Function
	
	Private Function GetFileExt(ByRef FullName As String) As String
		Dim fso As Object
		fso = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetExtensionName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetFileExt = fso.GetExtensionName(FullName)
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
	End Function
	
	'===========================================================================================
	
	Sub Test()
		Dim ptr As Integer
		Dim o1, i, o2 As String
		ReDim codes(3)
		codes(0) = &H28s
		codes(1) = &H11s
		codes(2) = &H11s
		codes(3) = &H3Es
		'Do
		ptr = 0 : pC = 0 : i = "" : o1 = "" : o2 = ""
		'codes(0) = Val("&h" & InputBox("Code:"))
		GetInst_00_3F(ptr, i, o1, o2)
		If o1 <> "" Then
			i = i & " " & o1
			If o2 <> "" Then i = i & "," & o2
		End If
		MsgBox(i & vbCr & "PC:" & Str(pC) & vbCr & "Ptr:" & Str(ptr))
		'Loop
	End Sub
	
	Private Function Swap(ByRef var1 As Object, ByRef var2 As Object) As Object
		Dim tmp As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object tmp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tmp = var1
		'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		var1 = var2
		'UPGRADE_WARNING: Couldn't resolve default property of object tmp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		var2 = tmp
	End Function
	
	Private Function IncPC(Optional ByRef incVal As Integer = 1) As Object
		pC = pC + incVal
		If pC > 65535 Then pC = pC - 65536
	End Function
	
	Private Function GetHexByte(ByRef num As Byte) As String
		GetHexByte = Hex(num)
		If num < 16 Then GetHexByte = "0" & GetHexByte
	End Function
	
	Private Function GetHexWord(ByRef byteL As Byte, ByRef byteH As Byte) As String
		GetHexWord = GetHexByte(byteH) & GetHexByte(byteL)
	End Function
	
	Private Function MakeRef(ByRef op As String) As Object
		op = "(" & op & ")"
	End Function
	
	Private Function GetReg8(ByRef code As Byte) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetReg8 = Choose(CShort(code And 7) + 1, "B", "C", "D", "E", "H", "L", "(HL)", "A")
	End Function
	
	Private Function GetReg16(ByRef code As Byte) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetReg16 = Choose(CShort(code And 3) + 1, "BC", "DE", "HL", "SP")
	End Function
	
	Private Function GetDisplacement(ByRef code As Byte, ByRef dis As Byte) As String
		If code <> &HDDs And code <> &HFDs Then Exit Function
		If code = &HDDs Then GetDisplacement = "IX" Else GetDisplacement = "IY"
		If dis < 128 Then GetDisplacement = GetDisplacement & "+" & GetHexByte(dis) Else GetDisplacement = GetDisplacement & "-" & GetHexByte(256 - dis)
		MakeRef(GetDisplacement)
	End Function
	
	Private Function GetAbsAddr(ByRef dis As Byte, ByVal curPC As Integer) As String
		curPC = curPC + dis
		If dis > 127 Then curPC = curPC - 256
		If curPC > 65535 Then curPC = curPC - 65536
		GetAbsAddr = GetHexWord(curPC Mod 256, (curPC - curPC Mod 256) / 256)
	End Function
	
	Private Function GetInst_00_3F(ByRef ptr As Integer, ByRef inst As String, ByRef op1 As String, ByRef op2 As String) As Object
		IncPC()
		Select Case codes(ptr) And 7
			Case 0
				Select Case codes(ptr)
					Case 0
						inst = "NOP"
					Case 8
						inst = "EX"
						op1 = "AF"
						op2 = "AF'"
					Case Else
						IncPC()
						op1 = GetAbsAddr(codes(ptr + 1), pC)
						inst = "JR"
						Select Case CShort(codes(ptr) And &H38s) / 8
							Case 2
								inst = "DJNZ"
							Case 4 To 7
								'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								op2 = Choose(CShort(codes(ptr) And &H18s) / 8 + 1, "NZ", "Z", "NC", "C")
								Swap(op1, op2)
						End Select
						ptr = ptr + 1
				End Select
				GoTo EndGetInstBelow40h
			Case 1, 2, 3
				Select Case codes(ptr)
					Case &H32s, &H3As
						op1 = "A"
					Case Else
						op1 = GetReg16(CShort(codes(ptr) And &H30s) / 16)
				End Select
			Case 4, 5, 6
				op1 = GetReg8(CShort(codes(ptr) And &H38s) / 8)
			Case 7
				'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				inst = Choose(CShort(codes(ptr) And &H38s) / 8 + 1, "RLCA", "RRCA", "RLA", "RRA", "DAA", "CPL", "SCF", "CCF")
				GoTo EndGetInstBelow40h
		End Select
		Select Case codes(ptr) And 15
			Case 1, 2, 6, 10, 14
				inst = "LD"
				Select Case codes(ptr)
					Case 1, &H11s, &H21s, &H31s
						IncPC(2)
						ptr = ptr + 2
						op2 = GetHexWord(codes(ptr - 1), codes(ptr))
					Case 2, &H12s
						MakeRef(op1)
						op2 = "A"
					Case &H22s, &H32s
						IncPC(2)
						ptr = ptr + 2
						op2 = GetHexWord(codes(ptr - 1), codes(ptr))
						Swap(op1, op2)
						MakeRef(op1)
					Case &HAs, &H1As
						op2 = "A"
						Swap(op1, op2)
						MakeRef(op2)
					Case &H2As, &H3As
						IncPC(2)
						ptr = ptr + 2
						op2 = GetHexWord(codes(ptr - 1), codes(ptr))
						MakeRef(op2)
					Case Else
						IncPC()
						ptr = ptr + 1
						op2 = GetHexByte(codes(ptr))
				End Select
			Case Else
				Select Case codes(ptr) And 15
					Case 3, 4, 12
						inst = "INC"
					Case 5, 11, 13
						inst = "DEC"
					Case 9
						inst = "ADD"
						op2 = "HL"
						Swap(op1, op2)
				End Select
		End Select
EndGetInstBelow40h: 
		ptr = ptr + 1
	End Function
	
	Private Function GetInst_40_BF(ByRef ptr As Integer, ByRef inst As String, ByRef op1 As String, ByRef op2 As String) As Object
		IncPC()
		op1 = GetReg8(codes(ptr))
		Select Case codes(ptr)
			Case &H40s To &H75s, &H77s To &H7Fs
				inst = "LD"
				op2 = GetReg8(CShort(codes(ptr) And &H38s) / 8)
				Swap(op1, op2)
			Case &H76s
				inst = "HALT"
				op1 = ""
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				inst = Choose(CShort(codes(ptr) And &H38s) / 8 + 1, "ADD", "ADC", "SUB", "SBC", "AND", "XOR", "OR", "CP")
				Select Case codes(ptr)
					Case &H80s To &H8Fs, &H98s To &H9Fs
						op2 = "A"
						Swap(op1, op2)
				End Select
		End Select
		ptr = ptr + 1
	End Function
	
	Private Function GetInst_C0_FF(ByRef ptr As Integer, ByRef inst As String, ByRef op1 As String, ByRef op2 As String) As Object
		IncPC()
		Select Case codes(ptr) And 15
			Case 0, 2, 4, 6, 7, 8, 10, 12, 14, 15
				Select Case codes(ptr) And 7
					Case 0, 2, 4
						'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						op1 = Choose(CShort(codes(ptr) And &H38s) / 8 + 1, "NZ", "Z", "NC", "C", "PO", "PE", "P", "M")
						Select Case codes(ptr) And 7
							Case 0
								inst = "RET"
							Case 2, 4
								If (codes(ptr) And 7) = 2 Then inst = "JP" Else inst = "CALL"
								IncPC(2)
								ptr = ptr + 2
								op2 = GetHexWord(codes(ptr - 1), codes(ptr))
						End Select
					Case 6
						IncPC()
						'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						inst = Choose(CShort(codes(ptr) And &H38s) / 8 + 1, "ADD", "ADC", "SUB", "SBC", "AND", "XOR", "OR", "CP")
						op1 = GetHexByte(codes(ptr + 1))
						Select Case codes(ptr)
							Case &HC6s, &HCEs, &HDEs
								op2 = "A"
								Swap(op1, op2)
						End Select
						ptr = ptr + 1
					Case 7
						inst = "RST"
						op1 = GetHexByte(codes(ptr) And &H38s)
				End Select
			Case 1, 5
				If codes(ptr) And 4 Then inst = "PUSH" Else inst = "POP"
				'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				op1 = Choose(CShort(codes(ptr) And &H30s) / 16 + 1, "BC", "DE", "HL", "AF")
			Case 3, 11
				Select Case codes(ptr)
					Case &HC3s
						IncPC(2)
						inst = "JP"
						op1 = GetHexWord(codes(ptr + 1), codes(ptr + 2))
						ptr = ptr + 2
					Case &HD3s, &HDBs
						IncPC()
						op1 = GetHexByte(codes(ptr + 1))
						MakeRef(op1)
						If codes(ptr) = &HD3s Then
							inst = "OUT"
							op2 = "A"
						Else
							inst = "IN"
							op2 = "A"
							Swap(op1, op2)
						End If
						ptr = ptr + 1
					Case &HE3s, &HEBs, &HF3s, &HFBs
						Select Case codes(ptr)
							Case &HE3s, &HEBs
								inst = "EX"
								op2 = "HL"
								If codes(ptr) = &HE3s Then op1 = "(SP)" Else op1 = "DE"
							Case &HF3s
								inst = "DI"
							Case &HFBs
								inst = "EI"
						End Select
				End Select
			Case 9
				'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				inst = Choose(CShort(codes(ptr) And &H30s) / 16 + 1, "RET", "EXX", "JP", "LD")
				Select Case codes(ptr)
					Case &HE9s
						op1 = "(HL)"
					Case &HF9s
						op1 = "SP"
						op2 = "HL"
				End Select
			Case 13
				IncPC(2)
				inst = "CALL"
				op1 = GetHexWord(codes(ptr + 1), codes(ptr + 2))
				ptr = ptr + 2
		End Select
		ptr = ptr + 1
	End Function
	
	Private Function GetExtInst_CB(ByRef ptr As Integer, ByRef inst As String, ByRef op1 As String, ByRef op2 As String) As Object
		IncPC()
		op1 = GetReg8(codes(ptr))
		If codes(ptr) And &HC0s Then
			op2 = Mid(Str(CShort(codes(ptr) And &H38s) / 8), 2)
			Swap(op1, op2)
			Select Case (codes(ptr) And &HC0s)
				Case &H40s
					inst = "BIT"
				Case &H80s
					inst = "RES"
				Case &HC0s
					inst = "SET"
			End Select
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			inst = Choose(CShort(codes(ptr) And &H38s) / 8 + 1, "RLC", "RRC", "RL", "RR", "SLA", "SRA", "DEFB", "SRL")
			If inst = "DEFB" Then op1 = "CB," & Hex(codes(ptr))
		End If
		ptr = ptr + 1
	End Function
	
	Private Function GetExtInst_DD_FD(ByRef ptr As Integer, ByRef inst As String, ByRef op1 As String, ByRef op2 As String) As Object
		Dim prevPtr, prevPC As Integer
		Dim strDis, strReg As String
		Dim byDis As Byte
		prevPtr = ptr
		prevPC = pC
		Select Case codes(ptr)
			Case Is < &H40s
				GetInst_00_3F(ptr, inst, op1, op2)
			Case &H40s To &HBFs
				GetInst_40_BF(ptr, inst, op1, op2)
			Case &HCBs
				IncPC(2)
				ptr = ptr + 2
				GetExtInst_CB(ptr, inst, op1, op2)
				If op1 <> "(HL)" And op2 <> "(HL)" Then
					inst = "DEFB"
					op1 = Hex(codes(ptr - 4)) & ",CB," & GetHexByte(codes(ptr - 2)) & "," & GetHexByte(codes(ptr - 1))
					op2 = ""
					Exit Function
				End If
			Case &HEDs, &HDDs, &HFDs
				IncPC()
				inst = "DEFB"
				op1 = Hex(codes(ptr - 1)) & "," & Hex(codes(ptr))
				ptr = ptr + 1
				Exit Function
			Case Else
				GetInst_C0_FF(ptr, inst, op1, op2)
		End Select
		If op1 <> "HL" And op1 <> "(HL)" And op2 <> "HL" And op2 <> "(HL)" Or codes(prevPtr) = &HEBs Then
			pC = prevPC
			IncPC()
			ptr = prevPtr + 1
			inst = "DEFB"
			op1 = Hex(codes(ptr - 2)) & "," & GetHexByte(codes(ptr - 1))
			op2 = ""
			Exit Function
		End If
		If codes(prevPtr - 1) And &H20s Then strReg = "IY" Else strReg = "IX"
		If (op1 = "(HL)" Or op2 = "(HL)") And inst <> "JP" Then
			Select Case codes(prevPtr)
				Case &H36s
					byDis = codes(ptr - 1)
					op2 = GetHexByte(codes(ptr))
				Case &HCBs
					byDis = codes(ptr - 2)
				Case Else
					byDis = codes(ptr)
			End Select
			If byDis > 127 Then
				strDis = "-"
				byDis = 256 - byDis
			Else
				strDis = "+"
			End If
			strDis = strDis & GetHexByte(byDis)
			op1 = Replace(op1, "HL", strReg & strDis)
			op2 = Replace(op2, "HL", strReg & strDis)
		End If
		op1 = Replace(op1, "HL", strReg)
		op2 = Replace(op2, "HL", strReg)
		If (Left(op1, 2) = "(I" Or Left(op2, 2) = "(I") And codes(prevPtr) <> &HCBs Then
			IncPC()
			ptr = ptr + 1
		End If
	End Function
	
	Private Function GetExtInst_ED(ByRef ptr As Integer, ByRef inst As String, ByRef op1 As String, ByRef op2 As String) As Object
		IncPC()
		Select Case codes(ptr)
			Case &H40s To &H4Bs, &H4Ds, &H50s To &H53s, &H56s To &H5Bs, &H5Es, &H60s, &H61s, &H62s, &H67s To &H6As, &H6Fs, &H72s, &H73s, &H78s To &H7Bs
				Select Case codes(ptr) And 7
					Case 0, 1
						op1 = GetReg8(CShort(codes(ptr) And &H38s) / 8)
						op2 = "(C)"
						If codes(ptr) And 1 Then
							inst = "OUT"
							Swap(op1, op2)
						Else
							inst = "IN"
						End If
					Case 2
						If codes(ptr) And 8 Then inst = "ADC" Else inst = "SBC"
						op1 = "HL"
						op2 = GetReg16(CShort(codes(ptr) And &H30s) / 16)
					Case 3
						IncPC(2)
						inst = "LD"
						op1 = GetHexWord(codes(ptr + 1), codes(ptr + 2))
						MakeRef(op1)
						op2 = GetReg16(CShort(codes(ptr) And &H30s) / 16)
						If codes(ptr) And 8 Then Swap(op1, op2)
						ptr = ptr + 2
					Case 4
						inst = "NEG"
					Case 5
						If codes(ptr) And 8 Then inst = "RETI" Else inst = "RETN"
					Case 6
						inst = "IM"
						'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						op1 = Choose(CShort(codes(ptr) And &H18s) / 8 + 1, "0", "", "1", "2")
					Case 7
						Select Case codes(ptr)
							Case &H47s, &H57s
								inst = "LD"
								op1 = "I"
								op2 = "A"
								If codes(ptr) And &H10s Then Swap(op1, op2)
							Case &H67s
								inst = "RRD"
							Case &H6Fs
								inst = "RLD"
						End Select
				End Select
			Case &HA0s To &HA3s, &HA8s To &HABs, &HB0s To &HB3s, &HB8s To &HBBs
				'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				inst = Choose(CShort(codes(ptr) And 3) + 1, "LD", "CP", "IN", "OT")
				If (codes(ptr) And &H13s) = 3 Then inst = "OUT"
				If codes(ptr) And 8 Then inst = inst & "D" Else inst = inst & "I"
				If codes(ptr) And &H10s Then inst = inst & "R"
			Case Else
				inst = "DEFB"
				op1 = "ED," & GetHexByte(codes(ptr))
		End Select
		ptr = ptr + 1
	End Function
End Module