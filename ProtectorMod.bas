Attribute VB_Name = "ProtectorMod"
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
'EXE Builder (Freeware)
'(C) 2004 Sriharish
' Email: Sriharish@msn.com?Subject=EXEProtector
' version 0.6
' Builder: Open
' Note: Read documentation for license and more
' Bug Report and suggestions: sriharish@msn.com?Subject=ExeProtector
'-----------------------------------------------------
'The Protected EXE Shell do not contain any viruses or trojans, worms etc
'------------------------------------------------------
'This Code will not transfer any information from your PC through the internet
' WARNING:DO NOT MODIFY THIS CODE
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
'Loader version : 0.6, internal: 0.06
'Loader Author: ooo My Self (sriharish)
'Antidebug: 0.3 Softice Detection:
'Special Thanks: Wilson Chan (China), John Taylor (OH), CN, anonymous emailer :<>
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
'Builder Module version : 0.6
'============================================
'Version Updates for loader and Builder
'============================================
'0.02: Added Reset Trial on new versions
'0.03: Added Encrypted Entry Point Info
'     : Fixed a Loader bug in Win 98
'     : Fixed a Loader crash in Win NT ( Although its not properly supported, so don't ask me any questions )
'0.04: (Private)Fixed Registration bug at last (Thanks to all pscode howlers)
'0.05: New "Cracking Tool" added to anti-debug list
'    : Supports .NET(not completely),VB,VC++,MASM exe files
'    : Improved memory handling techniques
'    : Improved Trial By Count
'    : Increased Loader Speed
'    : Fixed Memory Mapping Bug in loader
'    : Reduced Kegenerator size
'0.06: Added Commandline arguments to loader
'    : Removed Fake memory for .net exe's
'    : Improved CRC check
'    : New hardware finger print technique
'    : New EASY and Secure Registration
'    : Reduced Loader Size
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
Option Explicit
Type Enve
    appname As String
    Filename As String
    TKey As String
    LStart As String
    HPrint As String
    Dloader As String
    UKey As String
    email As String
    website As String
    Filemod As String
    InTrial As String
    Onecpy As String
    RTrial As String
    USDK As String
    ep As String
    filsize As String
    CurVer As String
    Slot As String
    Blacklist(0 To 2001) As String
    ByteStrip(1 To 1024) As String
End Type
Public Licinfo As Enve
Dim rcrypt As clsRijndael

Public Sub sortblacklist()

  Dim i As Integer

    Set rcrypt = New clsRijndael
    On Error GoTo error
    If builder.List1.ListCount = 0 Then
        Exit Sub
    End If
    builder.List1.ListIndex = 0
    For i = 0 To builder.List1.ListCount - 1
        If builder.List1.Text <> "" Then
            Licinfo.Blacklist(i) = rcrypt.EncryptString(builder.List1.Text, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
        End If
        If builder.List1.ListIndex <> builder.List1.ListCount - 1 Then
            builder.List1.ListIndex = builder.List1.ListIndex + 1
        End If
    Next i

Exit Sub

error:
    MsgBox Err.Description

End Sub

Public Sub BuildLicense()

    Set rcrypt = New clsRijndael
  Dim filenum As Integer
    filenum = FreeFile
    On Error GoTo error
    Licinfo.appname = rcrypt.EncryptString(builder.appname.Text, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    Licinfo.CurVer = rcrypt.EncryptString(builder.Text1.Text, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    Licinfo.Filename = rcrypt.EncryptString(builder.savedlg.filetitle, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    If Len(Hex(Licinfo.ep)) <= 4 Then
    Licinfo.Slot = "0"
    Else
    Licinfo.Slot = "1"
    End If
    
    If builder.Checkloader.Value = 1 Then
        Licinfo.LStart = rcrypt.EncryptString("1", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.LStart = rcrypt.EncryptString("0", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    If builder.CheckDLoader.Value = 1 Then
        Licinfo.Dloader = rcrypt.EncryptString("1", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.Dloader = rcrypt.EncryptString("0", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    If builder.Checkfilemod.Value = 1 Then
        Licinfo.Filemod = rcrypt.EncryptString("1", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.Filemod = rcrypt.EncryptString("0", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    If builder.Checkhardware.Value = 1 Then
        Licinfo.HPrint = rcrypt.EncryptString("1", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.HPrint = rcrypt.EncryptString("0", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    If builder.CheckIncreaseTrial.Value = 1 Then
        Licinfo.InTrial = rcrypt.EncryptString("1", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.InTrial = rcrypt.EncryptString("0", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    If builder.CheckOnecopy.Value = 1 Then
        Licinfo.Onecpy = rcrypt.EncryptString("1", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.Onecpy = rcrypt.EncryptString("0", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    If builder.CheckReset.Value = 1 Then
        Licinfo.RTrial = rcrypt.EncryptString("1", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.RTrial = rcrypt.EncryptString("0", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    If builder.website.Text = "" Then
        Licinfo.website = rcrypt.EncryptString("#", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.website = rcrypt.EncryptString(builder.website.Text, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    If builder.email.Text = "" Then
        Licinfo.email = rcrypt.EncryptString("#", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.email = rcrypt.EncryptString(builder.email.Text, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    If builder.CheckSDK.Value = 1 Then
        Licinfo.USDK = rcrypt.EncryptString("1", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
      Else
        Licinfo.USDK = rcrypt.EncryptString("0", Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    End If
    Licinfo.Slot = rcrypt.EncryptString(Licinfo.Slot, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    Licinfo.UKey = rcrypt.EncryptString(Licinfo.UKey, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    Licinfo.TKey = rcrypt.EncryptString(Licinfo.TKey, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    Licinfo.ep = rcrypt.EncryptString(Val(Licinfo.ep) + 1, Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    Licinfo.filsize = rcrypt.EncryptString(FileLen(builder.Filename), Chr(46) & Chr(46) & Chr(46) & Chr(80) & Chr(79) & Chr(82) & Chr(84) & Chr(85) & Chr(83) & Chr(46) & Chr(46) & Chr(46), False)
    On Error Resume Next
    Kill Left(builder.Filename.Text, Len(builder.Filename.Text) - Len(builder.savedlg.filetitle)) & "Portus.lic"
    Open Left(builder.Filename.Text, Len(builder.Filename.Text) - Len(builder.savedlg.filetitle)) & "Portus.lic" For Binary As filenum
    Put #filenum, , Licinfo
    Close #filenum

Exit Sub

error:
    MsgBox Err.Description
End
End Sub

Public Sub stripfile()

  Dim filenum As Integer
  Dim stripchunk As Byte
  Dim stringval As String
  Dim i, K As Integer

    Set rcrypt = New clsRijndael
    On Error GoTo fileerror
    filenum = FreeFile
    K = 1
    Open builder.Filename.Text For Binary As filenum
    For i = Val(Licinfo.ep + 1) To Val(Licinfo.ep + 1023)
        Get #filenum, i, stripchunk
        Licinfo.ByteStrip(K) = rcrypt.EncryptString(Hex(CDec(stripchunk)), "...#...S")
       
        'Put #filenum, i, 0
        stripchunk = Empty
        stringval = Empty
        K = K + 1
       
    Next
    Close #filenum
    Open builder.Filename.Text For Binary As filenum
    For i = Val(Licinfo.ep + 1) To Val(Licinfo.ep + 1022)
        Put #filenum, i, 0
        stripchunk = Empty
        stringval = Empty
        Next
    Close #filenum
Exit Sub

fileerror:
    MsgBox "Invalid file." & vbCrLf & Err.Description, vbCritical

End Sub
