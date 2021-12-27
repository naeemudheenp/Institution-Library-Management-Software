VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EXIT"
   ClientHeight    =   1155
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5190
   Icon            =   "exitdialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "exitdialog.frx":038A
   ScaleHeight     =   1155
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "                                ARE YOU SURE YOU WANT TO EXIT ?"
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Dialog.Hide
Form1.Show


End Sub

Private Sub OKButton_Click()





End


End Sub
