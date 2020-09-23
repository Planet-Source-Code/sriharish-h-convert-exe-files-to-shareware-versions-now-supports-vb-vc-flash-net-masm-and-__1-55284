VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Addstolen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Stolen Codes"
   ClientHeight    =   1020
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   6105
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   550
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   75
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Load from Text file:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Stolen Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Addstolen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
If Text1.Text = "" Then
Exit Sub
Else
builder.List1.AddItem Text1.Text
Text1.Text = ""
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim filenumber As Integer
Dim chunk As String
With CommonDialog1
.Filter = "Text files |*.txt|"
.Filename = ""
.ShowOpen
filenumber = FreeFile
If .Filename = "" Then Exit Sub
Open .Filename For Input As filenumber
Do Until EOF(filenumber)
Line Input #filenumber, chunk
builder.List1.AddItem chunk
Loop
Close #filenumber
End With
End Sub

