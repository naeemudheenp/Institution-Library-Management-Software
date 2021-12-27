VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   BackColor       =   &H8000000B&
   Caption         =   "REPORT"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9930
   Icon            =   "REPORT.frx":0000
   LinkTopic       =   "Form8"
   Picture         =   "REPORT.frx":038A
   ScaleHeight     =   8220
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "REPORT.frx":361B9
      Height          =   1095
      Left            =   3120
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1931
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   3
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
         DataField       =   "studentname"
         Caption         =   "studentname"
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
         DataField       =   "time"
         Caption         =   "time"
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
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "REPORT.frx":361CE
      Height          =   735
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   2040
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "issue"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   2040
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
      RecordSource    =   "select * from booklist;"
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
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6720
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5280
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4440
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   4215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer

Private Sub Command1_Click()
Command1.Visible = False






Form8.PrintForm

Command1.Visible = True




End Sub

Private Sub Command3_Click()

Command1.Visible = False

Form8.PrintForm
Command1.Visible = True

End Sub

Private Sub Command4_Click()
Form3.Command4_Click






End Sub

Private Sub Command5_Click()
Form5.Command1.Visible = False
Form5.Command2.Visible = False



Form5.PrintForm
Form5.Command1.Visible = True
Form5.Command2.Visible = True





End Sub

Private Sub Command6_Click()
dd.Visible = True


dd.Height = 10000000
End Sub

Private Sub Command2_Click()
Form8.Hide




Form1.Show

End Sub

Private Sub Form_Load()
Dim sum As Integer
Dim i As Integer
Dim j As Integer
Dim lft As Integer
Dim t As Integer
Dim price As Integer
Dim p As Integer
Dim iss As Integer


Adodc1.Refresh
iss = DataGrid2.ApproxCount

If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then



sum = 0
lft = 0
price = 0
GoTo l
Else
For i = 0 To DataGrid1.ApproxCount - 1
t = Adodc1.Recordset!booksleft
j = Adodc1.Recordset!qty
p = Adodc1.Recordset!price
lft = lft + t
sum = sum + j
price = price + p
If Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True Then
GoTo l
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
End If
End If
Adodc1.Recordset.MoveNext
Next

l:
Text1.Text = sum
Text2.Text = lft
Text3.Text = price
Text4.Text = iss
Dim dt As Date
dt = DateTime.Now



Text5.Text = dt





End If

End Sub

