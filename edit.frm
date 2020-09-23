VERSION 5.00
Begin VB.Form editfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update"
   ClientHeight    =   465
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   75
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   75
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "editfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()

    builder.List1.List(builder.List1.ListIndex) = Text1.Text

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

