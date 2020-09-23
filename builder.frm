VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form builder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exe Protector by Sriharish - Freeware!, Best Software Protection in PSCODE.COM"
   ClientHeight    =   6495
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11205
   Icon            =   "builder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11205
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   -120
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -240
      Top             =   -360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   6600
      ScaleHeight     =   6435
      ScaleWidth      =   15
      TabIndex        =   40
      Top             =   0
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   38
      Text            =   "1.2"
      Top             =   5500
      Width           =   735
   End
   Begin ExeProtector.EzCryptoApi Crypto 
      Left            =   -720
      Top             =   -840
      _ExtentX        =   1640
      _ExtentY        =   1905
      HashAlgorithm   =   3
      Password        =   ""
      EncryptionAlgorithm=   4
      Speed           =   10
   End
   Begin MSComDlg.CommonDialog savedlg 
      Left            =   -240
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox CheckTest 
      Caption         =   "Automatically Launch after protection"
      Height          =   255
      Left            =   6720
      TabIndex        =   29
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton Doprotect 
      Caption         =   "Ok, Protect EXE"
      Height          =   495
      Left            =   7560
      TabIndex        =   37
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CheckBox CheckReport 
      Caption         =   "Create Report after protection"
      Height          =   255
      Left            =   6720
      TabIndex        =   27
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton copy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   9960
      TabIndex        =   26
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton edit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   8880
      TabIndex        =   25
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7800
      TabIndex        =   24
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      Top             =   3840
      Width           =   975
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H00800080&
      Height          =   3180
      ItemData        =   "builder.frx":1CCA
      Left            =   6720
      List            =   "builder.frx":1CCC
      TabIndex        =   36
      Top             =   480
      Width           =   4335
   End
   Begin VB.CheckBox CheckSDK 
      Caption         =   "Allow access to use SDK DLL (Coming Soon)"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   6000
      Width           =   3735
   End
   Begin VB.CheckBox CheckBackup 
      Caption         =   "Create Exe Backup"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CheckBox CheckReset 
      Caption         =   "Reset Trial on new versions"
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CheckBox CheckOnecopy 
      Caption         =   "Allow only one copy"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CheckBox CheckIncreaseTrial 
      Caption         =   "Increase Trial Status on request"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CheckBox Checkfilemod 
      Caption         =   "Detect file modification"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox website 
      Height          =   285
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   16
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox email 
      Height          =   285
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   15
      Top             =   4080
      Width           =   4215
   End
   Begin VB.TextBox UnlockKey 
      Height          =   285
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   14
      Top             =   3600
      Width           =   4215
   End
   Begin VB.CommandButton browse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox CheckDLoader 
      Caption         =   "Download new loader if available"
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CheckBox Checkhardware 
      Caption         =   "Enable Hardware Finger Print"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Filename 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.CheckBox Checkloader 
      Caption         =   "Do not show loader at startup (required if you are protecting screensavers)"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   5895
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   19595265
      CurrentDate     =   38152
   End
   Begin VB.OptionButton Option3 
      Caption         =   "By Date"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "By Count"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "By Days"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Trialkey 
      Height          =   285
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   4
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox appname 
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   $"builder.frx":1CCE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   855
      Left            =   6720
      TabIndex        =   41
      Top             =   5640
      Width           =   4335
   End
   Begin VB.Label Label9 
      Caption         =   "Version"
      Height          =   255
      Left            =   4800
      TabIndex        =   39
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Stolen Codes:"
      Height          =   255
      Left            =   6720
      TabIndex        =   35
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Website or Buynow URL:"
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Support Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Unlock Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Exe File:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Trial Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Trial Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Program Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "&Open &Project"
      End
      Begin VB.Menu Save 
         Caption         =   "&Save &Project"
      End
      Begin VB.Menu buildkegen 
         Caption         =   "&Build Keygenerator"
      End
      Begin VB.Menu split 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "&Help"
      Begin VB.Menu document 
         Caption         =   "&Documentation"
         Shortcut        =   {F1}
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "builder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
'EXE Builder (Freeware)
'(C) 2004 Sriharish
' Email: Sriharish@msn.com?Subject=EXEProtector
' version 0.3
' Builder: Open
' Note: Read documentation for license and more
' Bug Report and suggestions: sriharish@msn.com?Subject=ExeProtector
'-----------------------------------------------------
'The Protected EXE Shell do not contain any virus or trojans, worms etc
'------------------------------------------------------
'This Code will not transfer any information from your PC through the internet
' WARNING:DO NOT MODIFY THIS CODE
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
'Loader version : 0.6, internal: 0.06
'Loader Author: ooo My Self (sriharish)
'Antidebug: 0.3 Softice Detection:1.2
'Special Thanks: Wilson Chan (China), John Taylor (OH), CN, anonymous emailer :<>
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
Option Explicit
Dim checkpe As CPEEditor
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub about_Click()

    frmSplash.Show 0, Me

End Sub

Private Sub Add_Click()

    Addstolen.Show 1

End Sub

Private Sub browse_Click()

    With savedlg
        .Filename = ""
        .Filter = "Win32 Executable|*.exe|Screensaver File |*.scr|"
        .ShowOpen
        If .Filename = "" Then Exit Sub
        Filename.Text = .Filename
    End With

End Sub

Private Sub buildkegen_Click()

    Dim reschunk() As Byte
    Dim filenumber As Integer
    With CommonDialog2
        .Filter = "Exe File |*.exe|"
        .Filename = ""
        .ShowSave
    filenumber = FreeFile
    If .Filename = "" Then Exit Sub
    Open .Filename For Binary As filenumber
        reschunk = LoadResData(3, "CUSTOM")
    Put #filenumber, , reschunk()
    Close #filenumber
    End With
    
End Sub



Private Sub copy_Click()

    Clipboard.Clear
    Clipboard.SetText List1.Text

End Sub

Private Sub Delete_Click()

    On Error Resume Next
        List1.RemoveItem List1.ListIndex

End Sub

Private Sub Doprotect_Click()
Dim shellsuccess As Long

    Set checkpe = New CPEEditor
    If Filename.Text = "" Then
        MsgBox "Filename is missing", vbCritical
        Exit Sub
    End If
    If appname.Text = "" Then
        MsgBox "Application Name is missing", vbCritical
        Exit Sub
    End If
    If Trialkey.Text = "" Or Len(Trialkey.Text) < 5 Then
        MsgBox "The Trial key should be minimum 5 Chars/Digits Long", vbCritical
        Exit Sub
    End If

    If UnlockKey.Text = "" Or Len(UnlockKey.Text) < 5 Then
        MsgBox "The Unlock key should be minumum 5 chars/digits long", vbCritical
        Exit Sub
    End If
    If Text1.Text = "" Or IsNumeric(Text1.Text) = False Or Len(Text1.Text) <> 4 Then
        MsgBox "Wrong version info. Correct format is X.XX", vbCritical
        Exit Sub
    End If
    checkpe.LoadFile Filename.Text
    Licinfo.ep = checkpe.OptionalHeader.AddressOfEntryPoint
    If Licinfo.ep = "" Or Licinfo.ep = 0 Then
        MsgBox "Invalid file to protect.", vbCritical
        Exit Sub
    End If
    If List1.ListCount > 2000 Then
        MsgBox "You cannot add more than 2000 black listed code. I hope you don't have many ;-)", vbCritical
        Exit Sub
    End If
    waitfrm.Show 0, Me
    waitfrm.Refresh
    If CheckBackup.Value = 1 Then
        FileCopy savedlg.Filename, savedlg.Filename & ".bak"
    End If
    Doprotect.Enabled = False
    preparekey1
    preparekey2
    sortblacklist
    stripfile
    BuildLicense
    If CheckReport.Value = 1 Then
    createreport
    End If
    Unload waitfrm
    Doprotect.Enabled = True
   
   Name Filename.Text As Filename.Text & ".locked"
    placeloader
    MsgBox "Your software is sucessfully protected. If your protected exe doesn't work then try to run the protected exe with portusem switch." & _
     vbCrLf & "For example: myexe.exe portusem or You can create a shortcut with target as myexe.exe portusem." & vbCrLf & "Even then if it fails to run then i'm sorry. " & vbCrLf & _
     "Adding commandline will not reduce loader's security.", vbInformation
    If CheckTest.Value = 1 Then
    shellsuccess = ShellExecute(0&, "Open", Filename.Text, 0&, 0&, 0&)
    End If
End Sub

Private Sub edit_Click()

    If List1.Text = "" Then
        Exit Sub
    End If
    editfrm.Text1.Text = List1.Text
    editfrm.Show 1

End Sub

Private Sub exit_Click()

    End

End Sub

Private Sub Form_Load()

    DTPicker1.Value = Date
    MsgBox "Currently the loader supports following executables: " & vbCrLf & vbCrLf & _
           "Visual Basic v6.0" & vbCrLf & _
            "Visual Basic Screensavers" & vbCrLf & _
           "Visual C++ (MFC)" & vbCrLf & _
           "Macromedia Flash" & vbCrLf & _
           "MASM 32 ( Partially Supported- no gurantee,sometimes insecure.)" & vbCrLf & _
           ".NET ( Partially Supported, no gurantee )" & vbCrLf & vbCrLf & _
           "Delphi Exe's are NOT supported,and many Exe's will not work if you don't read documentation.", vbInformation
         
End Sub

Private Sub Option3_Click()

    Text2.Text = ""
    Text3.Text = ""
    Text2.Enabled = False
    Text3.Enabled = False
    DTPicker1.Enabled = True

End Sub

Private Sub Option1_Click()

    Text2.Enabled = True
    Text3.Text = ""
    Text3.Enabled = False
    DTPicker1.Enabled = False

End Sub

Private Sub Option2_Click()

    Text2.Enabled = False
    Text2.Text = ""
    Text3.Enabled = True
    DTPicker1.Enabled = False

End Sub

Private Sub Save_Click()

  ' I know i could have used INI type but you know
  ' i don't think this part is very important
  ' than the protection this code offers
  
  Dim filenum As Integer

    filenum = FreeFile
    If savedlg.Filename = "" Then
        MsgBox "You must select a file to protect", vbCritical
        Exit Sub
    End If
    On Error GoTo error
    With CommonDialog1
        .Filename = ""
        .Filter = "Exe Protector Project |*.eprj|"
        .DialogTitle = "Save Project"
        .ShowSave
        If Dir(.Filename) <> "" Then
            If MsgBox("This file already exits. Are you sure you want to overwrite the existing file?", vbYesNo + vbExclamation, "Warning!") = vbNo Then
                Exit Sub
            End If
          Else
            Open .Filename For Output As filenum
            Print #filenum, "Exe Protector v0.3"
            Print #filenum, Filename.Text
            Print #filenum, appname.Text
            Print #filenum, Trialkey.Text
            If Option1.Value = True Then
                Print #filenum, "1"
                Print #filenum, Text2.Text
            End If
            If Option2.Value = True Then
                Print #filenum, "2"
                Print #filenum, Text3.Text
            End If
            If Option3.Value = True Then
                Print #filenum, "3"
                Print #filenum, Format(DTPicker1.Value, "MM-DD-YYYY")
            End If
            Print #filenum, Trim$(Checkloader.Value)
            Print #filenum, Trim$(Checkhardware.Value)
            Print #filenum, Trim$(CheckDLoader.Value)
            Print #filenum, UnlockKey.Text
            Print #filenum, email.Text
            Print #filenum, website.Text
            Print #filenum, Trim$(Checkfilemod.Value)
            Print #filenum, Trim$(CheckIncreaseTrial.Value)
            Print #filenum, Trim$(CheckOnecopy.Value)
            Print #filenum, Trim$(CheckReset.Value)
            Print #filenum, Text1.Text
            Print #filenum, Trim$(CheckBackup.Value)
            Print #filenum, Trim$(CheckSDK.Value)
            Print #filenum, Trim$(CheckReport.Value)
            Print #filenum, Trim$(CheckTest.Value)
            Print #filenum, savedlg.filetitle
            Close filenum
        End If
    End With

Exit Sub

error:
    MsgBox Err.Description

End Sub

Private Sub open_Click()

  Dim filenum As Integer
  Dim filetitle As String
  Dim tempstring As String

    filenum = FreeFile
    On Error GoTo error
    With CommonDialog1
        .Filename = ""
        .Filter = "Exe Protector Project |*.eprj|"
        .DialogTitle = "Open Project"
        .ShowOpen
        If .Filename = "" Then Exit Sub
        Open .Filename For Input As filenum
        Line Input #filenum, tempstring
        If tempstring <> "Exe Protector v0.3" Then
            MsgBox "Invalid or incompatile project file", vbCritical
            Exit Sub
        End If
    End With
    Line Input #filenum, tempstring
    Filename.Text = Trim$(tempstring)
    Line Input #filenum, tempstring
    appname.Text = tempstring
    Line Input #filenum, tempstring
    Trialkey.Text = Trim$(tempstring)
    Line Input #filenum, tempstring
    If tempstring = "1" Then
        Option1.Value = True
        Line Input #filenum, tempstring
        Text2.Text = tempstring
      Else
        If tempstring = "2" Then
            Option2.Value = True
            Line Input #filenum, tempstring
            Text3.Text = tempstring
          Else
            If tempstring = "3" Then
                Option3.Value = True
                Line Input #filenum, tempstring
                DTPicker1.Value = tempstring
            End If
        End If
    End If
    Line Input #filenum, tempstring
    Checkloader.Value = tempstring
    Line Input #filenum, tempstring
    Checkhardware.Value = tempstring
    Line Input #filenum, tempstring
    CheckDLoader.Value = tempstring
    Line Input #filenum, tempstring
    UnlockKey.Text = tempstring
    Line Input #filenum, tempstring
    email.Text = tempstring
    Line Input #filenum, tempstring
    website.Text = tempstring
    Line Input #filenum, tempstring
    Checkfilemod.Value = tempstring
    Line Input #filenum, tempstring
    CheckIncreaseTrial.Value = tempstring
    Line Input #filenum, tempstring
    CheckOnecopy.Value = tempstring
    Line Input #filenum, tempstring
    CheckReset.Value = tempstring
    Line Input #filenum, tempstring
    Text1.Text = tempstring
    Line Input #filenum, tempstring
    CheckBackup.Value = tempstring
    Line Input #filenum, tempstring
    CheckSDK.Value = tempstring
    Line Input #filenum, tempstring
    CheckReport.Value = tempstring
    Line Input #filenum, tempstring
    CheckTest.Value = tempstring
    Close #filenum
    With savedlg
        .Filename = ""
        .DialogTitle = "Open Exe File to Protect"
        .Filter = "Win32 Executable|*.exe|Screensaver File |*.scr|"
        .ShowOpen
        If .Filename = "" Then Exit Sub
        Filename.Text = .Filename
    End With

Exit Sub

Exit Sub

    savedlg.Filename = Text1.Text
error:
    MsgBox Err.Description, vbCritical

End Sub
Private Sub placeloader()
Dim reschunk() As Byte
Dim filenumber As Integer
filenumber = FreeFile
    Open Filename.Text For Binary As filenumber
        reschunk = LoadResData(102, "CUSTOM")
    Put #filenumber, , reschunk()
    Close #filenumber
End Sub

':) Ulli's VB Code Formatter V2.14.7 (6/24/2004 4:11:52 PM) 21 + 310 = 331 Lines
