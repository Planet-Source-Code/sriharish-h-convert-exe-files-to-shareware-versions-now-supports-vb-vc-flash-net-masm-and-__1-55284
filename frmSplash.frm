VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3570
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3570
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Continue"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Vote"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Email Author"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1080
      Top             =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
Dim shellsuccess As Long

shellsuccess = ShellExecute(fH, "Open", "mailto:sriharish@msn.com?Subject=Exeprotector", 0&, 0&, 0&)
End Sub

Private Sub Command3_Click()
builder.Show
Unload Me
End Sub

Private Sub Command2_Click()
Dim shellsuccess As Long

shellsuccess = ShellExecute(fH, "Open", "http://b.domaindlx.com/discbreaker/Votepage.asp", 0&, 0&, 0&)
End Sub

Private Sub Form_Load()
Clipboard.SetData Me.Picture, vbCFBitmap
End Sub
