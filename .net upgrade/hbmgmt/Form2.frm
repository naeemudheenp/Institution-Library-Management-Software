VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   LinkTopic       =   "Form2"
   ScaleHeight     =   8280
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   11
      Top             =   2280
      Width           =   350
   End
   Begin VB.CommandButton Command8 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   4080
      TabIndex        =   5
      Top             =   6960
      Width           =   4575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SETTINGS"
      Height          =   1215
      Left            =   8400
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SALES REPORT"
      Height          =   1215
      Left            =   8400
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SALES WINDOW"
      Height          =   1215
      Left            =   5880
      TabIndex        =   1
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STOCK MANAGMENT"
      Height          =   1215
      Left            =   8400
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "ABOUT"
      Height          =   345
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form2.frx":32C8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1850
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   8280
      Left            =   0
      Picture         =   "Form2.frx":BC2A
      ScaleHeight     =   8280
      ScaleWidth      =   2340
      TabIndex        =   6
      Top             =   0
      Width           =   2335
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   120
         Picture         =   "Form2.frx":1458C
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   8
         Top             =   1200
         Width           =   350
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK MANAGMENT"
      BeginProperty Font 
         Name            =   "@Adobe Heiti Std R"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK MANAGMENT"
      BeginProperty Font 
         Name            =   "@Adobe Heiti Std R"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD STOCK"
      BeginProperty Font 
         Name            =   "Adobe Heiti Std R"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD STOCK"
      BeginProperty Font 
         Name            =   "Adobe Heiti Std R"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD STOCK"
      BeginProperty Font 
         Name            =   "Adobe Heiti Std R"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD STOCK"
      BeginProperty Font 
         Name            =   "Adobe Heiti Std R"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show
End Sub

Private Sub Command2_Click()
Form4.Show
End Sub

Private Sub Command3_Click()
Form1.Show
End Sub

Private Sub Command4_Click()
Form6.Show
End Sub
b Command5_Click()
MsgBox
Private SuSub("HBMGMT\n version 1.0\nDEVELOPED BY NAEEMUDHEEN ASLAM ")
End

Private Sub Command6_Click()
Form9.Show

End Sub

Private Sub Label1_Click()
Form3.Show

End Sub

Private Sub Label7_Click()

End Sub
