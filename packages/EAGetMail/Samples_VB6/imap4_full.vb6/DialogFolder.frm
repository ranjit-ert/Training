VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Dialog1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Folder"
   ClientHeight    =   6915
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView lstFolder 
      Height          =   6015
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10610
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    lstFolder.SelectedItem = Nothing
    Me.Hide
End Sub

Private Sub OKButton_Click()
 Dim oNode As Node
 Set oNode = lstFolder.SelectedItem
 If oNode Is Nothing Then
    MsgBox "Please select a folder"
    Exit Sub
 End If
 
 If oNode.Parent Is Nothing Then
    MsgBox "Please select a folder"
    Exit Sub
 End If
 
 Me.Hide
End Sub
