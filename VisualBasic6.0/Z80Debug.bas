Attribute VB_Name = "modMain"

'=========================================================================================='
'                                                                                          '
'              Z80 Debugger for ED-Laboratory's Microprocessor Trainer MPT-1              '
'                                                                                          '
'                Copyright (C) Iskandar Z. Nasibu, Gorontalo, February 2005                '
'                                                                                          '
'=========================================================================================='


Option Explicit

Public Const cSrcFileExt As String = "BIN"
Const cOutFileExt As String = "Z80"
Const cSrcFileStartStr As String = "<Z80_Executable_Codes>"
Const cSrcFileEndStr As String = _
    "<ZulNs#05-11-1970#Viva_New_Technology_Protocol#Gorontalo#Feb-2005>"

Dim codes() As Byte, pC As Long
Dim srcFile As String, outFile As String

Sub Main()
    Dim ptr As Long, ln As Long, i As Long, inst As String, op1 As String, op2 As String
    On Error GoTo ErrorHandler
    If Not ReadSrcFile Then
        If srcFile = "" Then
            Exit Sub
        Else
            GoTo ErrReadSrcFile
        End If
    End If
    DelFile outFile
    Open outFile For Output As #1
    Print #1, GetHorLine(44)
    Print #1, "ADDRESS MACHINE-CODE  #   OPCODE  OPERAND"
    Print #1, GetHorLine(44)
    Print #1,
    Do While ptr < UBound(codes)
        ln = codes(ptr) + 256 * codes(ptr + 1)
        If ln = 0 Then Exit Do
        If ptr Then Print #1,
        pC = codes(ptr + 2) + 256 * codes(ptr + 3)
        ptr = ptr + 4
        ln = ln + ptr
        Do While ptr < ln
            Print #1, GetHex4(pC); ":"; Tab(9);
            i = ptr
            op1 = ""
            op2 = ""
            Select Case codes(ptr)
            Case 0 To &H3F
                GetInst_00_3F ptr, inst, op1, op2
            Case &H40 To &HBF
                GetInst_40_BF ptr, inst, op1, op2
            Case &HCB
                IncPC
                ptr = ptr + 1
                GetExtInst_CB ptr, inst, op1, op2
            Case &HDD, &HFD
                IncPC
                ptr = ptr + 1
                GetExtInst_DD_FD ptr, inst, op1, op2
            Case &HED
                IncPC
                ptr = ptr + 1
                GetExtInst_ED ptr, inst, op1, op2
            Case Else
                GetInst_C0_FF ptr, inst, op1, op2
            End Select
            For i = i To ptr - 1
                Print #1, GetHexByte(codes(i)); " ";
            Next
            Print #1, Tab(23); "#   "; inst;
            If op1 <> "" Then
                Print #1, Tab(35); op1;
                If op2 <> "" Then Print #1, ","; op2;
            End If
            Print #1, ""
        Loop
    Loop
    Close #1
    MsgBox "Debugging process successful.", vbInformation
    Shell "Notepad.exe " & outFile, vbNormalFocus
    Exit Sub
ErrorHandler:
    Close #1
    DelFile outFile
ErrReadSrcFile:
    MsgBox srcFile & " is not a valid Z80 executable file." & vbCr & _
        "Debugging process aborted.", vbCritical
End Sub

Private Function GetHorLine(ChrNum As Long) As String
    For ChrNum = 1 To ChrNum
        GetHorLine = GetHorLine & "="
    Next
End Function

Private Function GetHex4(num As Long) As String
    GetHex4 = Hex(num)
    Do While Len(GetHex4) < 4
        GetHex4 = "0" & GetHex4
    Loop
End Function

Private Function ReadSrcFile() As Boolean
    Dim flPtr As Long, str As String, codesPtr As Long, readOK As Boolean
    If Not GetSrcFileName Then Exit Function
    Open srcFile For Binary Access Read As #1
    str = cSrcFileStartStr
    Get #1, , str
    If str <> cSrcFileStartStr Then GoTo EndReadSrcFile
    codesPtr = -1
    Do While Not EOF(1)
        flPtr = Seek(1)
        str = cSrcFileEndStr
        Get #1, , str
        If str = cSrcFileEndStr Then
            readOK = True
            Exit Do
        End If
        Seek #1, flPtr
        codesPtr = codesPtr + 1
        ReDim Preserve codes(codesPtr)
        Get #1, , codes(codesPtr)
    Loop
    If readOK Then
        ReadSrcFile = True
    Else
        ReDim Preserve codes(0)
    End If
EndReadSrcFile:
    Close #1
End Function

Private Function GetSrcFileName() As Boolean
    Dim CmdTail As String, fso As Object, Path As String, UserRespons As VbMsgBoxResult
    CmdTail = Command()
    If CmdTail = "" Then
        If Not GetSrcFileNameFromDlg Then Exit Function
    Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FileExists(CmdTail) Then _
            If fso.GetExtensionName(CmdTail) = "" Then CmdTail = CmdTail & "." & cSrcFileExt
        If fso.FileExists(CmdTail) Then
            Path = fso.GetParentFolderName(CmdTail)
            If Path = "" Then
                srcFile = CmdTail
            Else
                ChDrive fso.GetDriveName(Path)
                ChDir Path
                srcFile = fso.GetFileName(CmdTail)
            End If
            Set fso = Nothing
        Else
            Set fso = Nothing
            UserRespons = MsgBox("Can't found '" & CmdTail & "' file or '" & _
                CmdTail & "' is not a legal file name." & vbCr & _
                "Try to find it or another file by your self?", vbQuestion + vbOKCancel)
            If UserRespons = vbOK Then
                If Not GetSrcFileNameFromDlg Then Exit Function
            Else
                Exit Function
            End If
        End If
    End If
    If UCase(GetFileExt(srcFile)) = cOutFileExt Then outFile = srcFile _
    Else outFile = GetFileName(srcFile)
    outFile = outFile & "." & cOutFileExt
    GetSrcFileName = True
End Function

Private Function GetSrcFileNameFromDlg() As Boolean
    Dim dlg As New frmDlgFileOpen, blnExit As Boolean
    dlg.Show vbModal
    blnExit = dlg.ExitMode
    If blnExit Then srcFile = dlg.FileName
    Unload dlg
    Set dlg = Nothing
    If blnExit Then GetSrcFileNameFromDlg = True _
    Else MsgBox "No file selected. Debugging process aborted.", vbInformation
End Function

Private Function IsFileExist(FileName As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IsFileExist = fso.FileExists(FileName)
    Set fso = Nothing
End Function

Private Function DelFile(FileName As String)
    Dim fso As Object
    If IsFileExist(FileName) Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile FileName, True
        Set fso = Nothing
    End If
End Function

Private Function GetFileName(FullName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetBaseName(FullName)
    Set fso = Nothing
End Function

Private Function GetFileExt(FullName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileExt = fso.GetExtensionName(FullName)
    Set fso = Nothing
End Function

'===========================================================================================

Sub Test()
    Dim ptr As Long, i As String, o1 As String, o2 As String
    ReDim codes(3)
    codes(0) = &H28
    codes(1) = &H11
    codes(2) = &H11
    codes(3) = &H3E
    'Do
        ptr = 0: pC = 0: i = "": o1 = "": o2 = ""
        'codes(0) = Val("&h" & InputBox("Code:"))
        GetInst_00_3F ptr, i, o1, o2
        If o1 <> "" Then
            i = i & " " & o1
            If o2 <> "" Then i = i & "," & o2
        End If
        MsgBox i & vbCr & "PC:" & str(pC) & vbCr & "Ptr:" & str(ptr)
    'Loop
End Sub

Private Function Swap(var1, var2)
    Dim tmp
    tmp = var1
    var1 = var2
    var2 = tmp
End Function

Private Function IncPC(Optional incVal As Long = 1)
    pC = pC + incVal
    If pC > 65535 Then pC = pC - 65536
End Function

Private Function GetHexByte(num As Byte) As String
    GetHexByte = Hex(num)
    If num < 16 Then GetHexByte = "0" & GetHexByte
End Function

Private Function GetHexWord(byteL As Byte, byteH As Byte) As String
    GetHexWord = GetHexByte(byteH) & GetHexByte(byteL)
End Function

Private Function MakeRef(op As String)
    op = "(" & op & ")"
End Function

Private Function GetReg8(code As Byte) As String
    GetReg8 = Choose((code And 7) + 1, "B", "C", "D", "E", "H", "L", "(HL)", "A")
End Function

Private Function GetReg16(code As Byte) As String
    GetReg16 = Choose((code And 3) + 1, "BC", "DE", "HL", "SP")
End Function

Private Function GetDisplacement(code As Byte, dis As Byte) As String
    If code <> &HDD And code <> &HFD Then Exit Function
    If code = &HDD Then GetDisplacement = "IX" Else GetDisplacement = "IY"
    If dis < 128 Then GetDisplacement = GetDisplacement & "+" & GetHexByte(dis) _
    Else GetDisplacement = GetDisplacement & "-" & GetHexByte(256 - dis)
    MakeRef GetDisplacement
End Function

Private Function GetAbsAddr(dis As Byte, ByVal curPC As Long) As String
    curPC = curPC + dis
    If dis > 127 Then curPC = curPC - 256
    If curPC > 65535 Then curPC = curPC - 65536
    GetAbsAddr = GetHexWord(curPC Mod 256, (curPC - curPC Mod 256) / 256)
End Function

Private Function GetInst_00_3F(ptr As Long, inst As String, op1 As String, op2 As String)
    IncPC
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
            IncPC
            op1 = GetAbsAddr(codes(ptr + 1), pC)
            inst = "JR"
            Select Case (codes(ptr) And &H38) / 8
            Case 2
                inst = "DJNZ"
            Case 4 To 7
                op2 = Choose((codes(ptr) And &H18) / 8 + 1, "NZ", "Z", "NC", "C")
                Swap op1, op2
            End Select
            ptr = ptr + 1
        End Select
        GoTo EndGetInstBelow40h
    Case 1, 2, 3
        Select Case codes(ptr)
        Case &H32, &H3A
            op1 = "A"
        Case Else
            op1 = GetReg16((codes(ptr) And &H30) / 16)
        End Select
    Case 4, 5, 6
        op1 = GetReg8((codes(ptr) And &H38) / 8)
    Case 7
        inst = Choose((codes(ptr) And &H38) / 8 + 1, _
            "RLCA", "RRCA", "RLA", "RRA", "DAA", "CPL", "SCF", "CCF")
        GoTo EndGetInstBelow40h
    End Select
    Select Case codes(ptr) And 15
    Case 1, 2, 6, 10, 14
        inst = "LD"
        Select Case codes(ptr)
        Case 1, &H11, &H21, &H31
            IncPC 2
            ptr = ptr + 2
            op2 = GetHexWord(codes(ptr - 1), codes(ptr))
        Case 2, &H12
            MakeRef op1
            op2 = "A"
        Case &H22, &H32
            IncPC 2
            ptr = ptr + 2
            op2 = GetHexWord(codes(ptr - 1), codes(ptr))
            Swap op1, op2
            MakeRef op1
        Case &HA, &H1A
            op2 = "A"
            Swap op1, op2
            MakeRef op2
        Case &H2A, &H3A
            IncPC 2
            ptr = ptr + 2
            op2 = GetHexWord(codes(ptr - 1), codes(ptr))
            MakeRef op2
        Case Else
            IncPC
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
            Swap op1, op2
        End Select
    End Select
EndGetInstBelow40h:
    ptr = ptr + 1
End Function

Private Function GetInst_40_BF(ptr As Long, inst As String, op1 As String, op2 As String)
    IncPC
    op1 = GetReg8(codes(ptr))
    Select Case codes(ptr)
    Case &H40 To &H75, &H77 To &H7F
        inst = "LD"
        op2 = GetReg8((codes(ptr) And &H38) / 8)
        Swap op1, op2
    Case &H76
        inst = "HALT"
        op1 = ""
    Case Else
        inst = Choose((codes(ptr) And &H38) / 8 + 1, _
            "ADD", "ADC", "SUB", "SBC", "AND", "XOR", "OR", "CP")
        Select Case codes(ptr)
        Case &H80 To &H8F, &H98 To &H9F
            op2 = "A"
            Swap op1, op2
        End Select
    End Select
    ptr = ptr + 1
End Function

Private Function GetInst_C0_FF(ptr As Long, inst As String, op1 As String, op2 As String)
    IncPC
    Select Case codes(ptr) And 15
    Case 0, 2, 4, 6, 7, 8, 10, 12, 14, 15
        Select Case codes(ptr) And 7
        Case 0, 2, 4
            op1 = Choose((codes(ptr) And &H38) / 8 + 1, _
                "NZ", "Z", "NC", "C", "PO", "PE", "P", "M")
            Select Case codes(ptr) And 7
            Case 0
                inst = "RET"
            Case 2, 4
                If (codes(ptr) And 7) = 2 Then inst = "JP" Else inst = "CALL"
                IncPC 2
                ptr = ptr + 2
                op2 = GetHexWord(codes(ptr - 1), codes(ptr))
            End Select
        Case 6
            IncPC
            inst = Choose((codes(ptr) And &H38) / 8 + 1, _
                "ADD", "ADC", "SUB", "SBC", "AND", "XOR", "OR", "CP")
            op1 = GetHexByte(codes(ptr + 1))
            Select Case codes(ptr)
            Case &HC6, &HCE, &HDE
                op2 = "A"
                Swap op1, op2
            End Select
            ptr = ptr + 1
        Case 7
            inst = "RST"
            op1 = GetHexByte(codes(ptr) And &H38)
        End Select
    Case 1, 5
        If codes(ptr) And 4 Then inst = "PUSH" Else inst = "POP"
        op1 = Choose((codes(ptr) And &H30) / 16 + 1, "BC", "DE", "HL", "AF")
    Case 3, 11
        Select Case codes(ptr)
        Case &HC3
            IncPC 2
            inst = "JP"
            op1 = GetHexWord(codes(ptr + 1), codes(ptr + 2))
            ptr = ptr + 2
        Case &HD3, &HDB
            IncPC
            op1 = GetHexByte(codes(ptr + 1))
            MakeRef op1
            If codes(ptr) = &HD3 Then
                inst = "OUT"
                op2 = "A"
            Else
                inst = "IN"
                op2 = "A"
                Swap op1, op2
            End If
            ptr = ptr + 1
        Case &HE3, &HEB, &HF3, &HFB
            Select Case codes(ptr)
            Case &HE3, &HEB
                inst = "EX"
                op2 = "HL"
                If codes(ptr) = &HE3 Then op1 = "(SP)" Else op1 = "DE"
            Case &HF3
                inst = "DI"
            Case &HFB
                inst = "EI"
            End Select
        End Select
    Case 9
        inst = Choose((codes(ptr) And &H30) / 16 + 1, "RET", "EXX", "JP", "LD")
        Select Case codes(ptr)
        Case &HE9
            op1 = "(HL)"
        Case &HF9
            op1 = "SP"
            op2 = "HL"
        End Select
    Case 13
        IncPC 2
        inst = "CALL"
        op1 = GetHexWord(codes(ptr + 1), codes(ptr + 2))
        ptr = ptr + 2
    End Select
    ptr = ptr + 1
End Function

Private Function GetExtInst_CB(ptr As Long, inst As String, op1 As String, op2 As String)
    IncPC
    op1 = GetReg8(codes(ptr))
    If codes(ptr) And &HC0 Then
        op2 = Mid(str((codes(ptr) And &H38) / 8), 2)
        Swap op1, op2
        Select Case (codes(ptr) And &HC0)
        Case &H40
            inst = "BIT"
        Case &H80
            inst = "RES"
        Case &HC0
            inst = "SET"
        End Select
    Else
        inst = Choose((codes(ptr) And &H38) / 8 + 1, _
            "RLC", "RRC", "RL", "RR", "SLA", "SRA", "DEFB", "SRL")
        If inst = "DEFB" Then op1 = "CB," & Hex(codes(ptr))
    End If
    ptr = ptr + 1
End Function

Private Function GetExtInst_DD_FD(ptr As Long, inst As String, op1 As String, op2 As String)
    Dim prevPtr As Long, prevPC As Long, strDis As String, byDis As Byte, strReg As String
    prevPtr = ptr
    prevPC = pC
    Select Case codes(ptr)
    Case Is < &H40
        GetInst_00_3F ptr, inst, op1, op2
    Case &H40 To &HBF
        GetInst_40_BF ptr, inst, op1, op2
    Case &HCB
        IncPC 2
        ptr = ptr + 2
        GetExtInst_CB ptr, inst, op1, op2
        If op1 <> "(HL)" And op2 <> "(HL)" Then
            inst = "DEFB"
            op1 = Hex(codes(ptr - 4)) & ",CB," & GetHexByte(codes(ptr - 2)) & "," & _
                GetHexByte(codes(ptr - 1))
            op2 = ""
            Exit Function
        End If
    Case &HED, &HDD, &HFD
        IncPC
        inst = "DEFB"
        op1 = Hex(codes(ptr - 1)) & "," & Hex(codes(ptr))
        ptr = ptr + 1
        Exit Function
    Case Else
        GetInst_C0_FF ptr, inst, op1, op2
    End Select
    If op1 <> "HL" And op1 <> "(HL)" And op2 <> "HL" And op2 <> "(HL)" _
    Or codes(prevPtr) = &HEB Then
        pC = prevPC
        IncPC
        ptr = prevPtr + 1
        inst = "DEFB"
        op1 = Hex(codes(ptr - 2)) & "," & GetHexByte(codes(ptr - 1))
        op2 = ""
        Exit Function
    End If
    If codes(prevPtr - 1) And &H20 Then strReg = "IY" Else strReg = "IX"
    If (op1 = "(HL)" Or op2 = "(HL)") And inst <> "JP" Then
        Select Case codes(prevPtr)
        Case &H36
            byDis = codes(ptr - 1)
            op2 = GetHexByte(codes(ptr))
        Case &HCB
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
    If (Left(op1, 2) = "(I" Or Left(op2, 2) = "(I") And codes(prevPtr) <> &HCB Then
        IncPC
        ptr = ptr + 1
    End If
End Function

Private Function GetExtInst_ED(ptr As Long, inst As String, op1 As String, op2 As String)
    IncPC
    Select Case codes(ptr)
    Case &H40 To &H4B, &H4D, &H50 To &H53, &H56 To &H5B, &H5E, _
    &H60, &H61, &H62, &H67 To &H6A, &H6F, &H72, &H73, &H78 To &H7B
        Select Case codes(ptr) And 7
        Case 0, 1
            op1 = GetReg8((codes(ptr) And &H38) / 8)
            op2 = "(C)"
            If codes(ptr) And 1 Then
                inst = "OUT"
                Swap op1, op2
            Else
                inst = "IN"
            End If
        Case 2
            If codes(ptr) And 8 Then inst = "ADC" Else inst = "SBC"
            op1 = "HL"
            op2 = GetReg16((codes(ptr) And &H30) / 16)
        Case 3
            IncPC 2
            inst = "LD"
            op1 = GetHexWord(codes(ptr + 1), codes(ptr + 2))
            MakeRef op1
            op2 = GetReg16((codes(ptr) And &H30) / 16)
            If codes(ptr) And 8 Then Swap op1, op2
            ptr = ptr + 2
        Case 4
            inst = "NEG"
        Case 5
            If codes(ptr) And 8 Then inst = "RETI" Else inst = "RETN"
        Case 6
            inst = "IM"
            op1 = Choose((codes(ptr) And &H18) / 8 + 1, "0", "", "1", "2")
        Case 7
            Select Case codes(ptr)
            Case &H47, &H57
                inst = "LD"
                op1 = "I"
                op2 = "A"
                If codes(ptr) And &H10 Then Swap op1, op2
            Case &H67
                inst = "RRD"
            Case &H6F
                inst = "RLD"
            End Select
        End Select
    Case &HA0 To &HA3, &HA8 To &HAB, &HB0 To &HB3, &HB8 To &HBB
        inst = Choose((codes(ptr) And 3) + 1, "LD", "CP", "IN", "OT")
        If (codes(ptr) And &H13) = 3 Then inst = "OUT"
        If codes(ptr) And 8 Then inst = inst & "D" Else inst = inst & "I"
        If codes(ptr) And &H10 Then inst = inst & "R"
    Case Else
        inst = "DEFB"
        op1 = "ED," & GetHexByte(codes(ptr))
    End Select
    ptr = ptr + 1
End Function
