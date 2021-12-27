VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   0  'None
   Caption         =   "SLEEP"
   ClientHeight    =   9195
   ClientLeft      =   2250
   ClientTop       =   1935
   ClientWidth     =   6990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTip.frx":0000
   ScaleHeight     =   9195
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   4200
   End
   Begin VB.PictureBox Picture1 
      Height          =   9255
      Left            =   -120
      Picture         =   "frmTip.frx":17B5A
      ScaleHeight     =   9195
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   -120
      Width           =   7095
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer






Private Sub Form_KeyPress(KeyAscii As Integer)
frmTip.Hide
Form1.Show

End Sub

Private Sub Form_Load()
   counter = 1
    Picture1.Picture = LoadPicture("E:\software\images\SCREENSAVER IMAGES\image" & counter & ".jpg")
    Timer1.Interval = 5000 '5 seconds
    Timer1.Enabled = True
    
End Sub

Public Sub DisplayCurrentTip()
    
    
End Sub

Private Sub Picture1_Click()
frmTip.Hide
Form1.Show

End Sub

Private Sub Timer1_Timer()
counter = counter + 1
    If counter > 10 Then counter = 1
    Picture1.Picture = LoadPicture("E:\software\images\SCREENSAVER IMAGES\image" & counter & ".jpg")
End Sub
