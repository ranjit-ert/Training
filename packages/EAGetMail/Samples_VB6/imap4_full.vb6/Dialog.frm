VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1305
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox textFolder 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Folder Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    textFolder.Text = ""
    Me.Hide
End Sub

Private Sub OKButton_Click()
    If textFolder.Text = "" Then
        MsgBox "Please input folder name!"
        Exit Sub
    End If
    Me.Hide
End Sub
