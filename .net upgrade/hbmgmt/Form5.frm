VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form5"
   ScaleHeight     =   3420
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "MAIN MENU"
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SHOW REPORT BEFORE SPECIFIED DATE"
      Height          =   735
      Left            =   4800
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SHOW REPORT AFTER SPECIFIED DATE"
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SHOW REPORT ON SPECIFIED DATE"
      Height          =   735
      Left            =   4800
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Text            =   "SELECT DATE"
      Top             =   480
      Width           =   4695
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   4575
      _Version        =   524288
      _ExtentX        =   8070
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2018
      Month           =   5
      Day             =   27
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "                                                                                                            SALES REPORT"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Calendar1_Click()

End Sub

Private Sub Command1_Click()
Form6.Show

End Sub

Private Sub Command2_Click()
Form8.Show
End Sub

Private Sub Command3_Click()
Form7.Show
End Sub

Private Sub Command4_Click()
Me.Hide
Form2.Show

End Sub

Public Sub Form_Load()

End Sub
