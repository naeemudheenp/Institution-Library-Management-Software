VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "HELO"
   ClientHeight    =   3990
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":038A
   ScaleHeight     =   3990
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
     frmSplash.Hide
  frmLogin.Show
  
End Sub

Private Sub Frame1_Click()
  frmSplash.Hide
  frmLogin.Show
  
End Sub


