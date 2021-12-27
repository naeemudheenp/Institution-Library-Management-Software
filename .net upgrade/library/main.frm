VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "MAIN SCREEN"
   ClientHeight    =   9495
   ClientLeft      =   195
   ClientTop       =   540
   ClientWidth     =   20250
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "main.frx":038A
   ScaleHeight     =   9495
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command12 
      Caption         =   "SLEEP"
      Height          =   735
      Left            =   0
      TabIndex        =   14
      Top             =   9480
      Width           =   5295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "CLOSE GRID"
      Height          =   495
      Left            =   15000
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ISSUE LIST"
      Height          =   2655
      Left            =   3240
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "REPORT"
      Height          =   2055
      Left            =   3240
      TabIndex        =   9
      Top             =   7440
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   9600
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT* FROM BOOKLIST;"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   " SEARCH"
      Height          =   495
      Left            =   13080
      TabIndex        =   12
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "main.frx":1D84B
      Height          =   10455
      Left            =   5280
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   18441
      _Version        =   393216
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "bookid"
         Caption         =   "bookid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "bookname"
         Caption         =   "bookname"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "author"
         Caption         =   "author"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "publication"
         Caption         =   "publication"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "qty"
         Caption         =   "qty"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "booksleft"
         Caption         =   "booksleft"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "shelf"
         Caption         =   "shelf"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "rack"
         Caption         =   "rack"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "price"
         Caption         =   "price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NOT ISSUED BOOKS"
      Height          =   2415
      Left            =   3240
      TabIndex        =   7
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SETTINGS"
      Height          =   2415
      Left            =   -120
      TabIndex        =   6
      Top             =   5040
      Width           =   3375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ISSUE"
      Height          =   1935
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ABOUT"
      Height          =   2055
      Left            =   -120
      TabIndex        =   8
      Top             =   7440
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000007&
      Caption         =   "EXIT"
      Height          =   735
      Left            =   -240
      TabIndex        =   10
      Top             =   10200
      Width           =   5535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BOOK LIST "
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "ADD BOOK"
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "                    CHOOSE THE SERVICE"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Combo1_Change()

End Sub

Private Sub cmbsrc_Form1()







End Sub

Private Sub Command1_Click()
Form2.Show

End Sub

Private Sub Command10_Click()
Form4.Show

End Sub

Private Sub Command11_Click()
Form8.Show




End Sub

Private Sub Command12_Click()
Form1.Hide
frmTip.Show

End Sub

Private Sub Command2_Click()
Form3.Show

End Sub

Private Sub Command3_Click()
Form1.Hide
Dialog.Show

End Sub

Private Sub Command4_Click()
MsgBox ("APP VERSION:1.0 CREATED BY NAEEMUDHEEN ASLAM ")

End Sub

Private Sub Command5_Click()
    Form9.Show
    
End Sub

Private Sub Command6_Click()
Form5.Show

End Sub

Private Sub Command7_Click()
Form11.Show




End Sub

Private Sub Command9_Click()
DataGrid1.Visible = False
Command9.Visible = False





End Sub



Private Sub Form_Load()
If DataGrid1.Visible = True Then
Command9.Visible = True
End If





End Sub

Private Sub Text1_Change()
Adodc1.Refresh
Adodc1.RecordSource = "Select * from booklist where Bookname Like '" & Text1.Text & "%' OR Bookid Like '" & Text1.Text & "%' OR author Like '" & Text1.Text & "%'OR publication Like '" & Text1.Text & "%' "
Adodc1.Refresh
DataGrid1.Refresh
If Adodc1.Recordset.EOF = True Then
GoTo fr
Else
DataGrid1.Visible = True
End If
Adodc1.Refresh
DataGrid1.Refresh
fr:
If DataGrid1.Visible = True Then
Command9.Visible = True
End If
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""

End Sub

