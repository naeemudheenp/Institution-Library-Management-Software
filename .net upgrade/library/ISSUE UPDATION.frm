VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   Caption         =   "ISSUE UPDATION"
   ClientHeight    =   10395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10860
   Icon            =   "ISSUE UPDATION.frx":0000
   LinkTopic       =   "Form5"
   Picture         =   "ISSUE UPDATION.frx":038A
   ScaleHeight     =   10395
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   4920
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   6240
      ScaleHeight     =   4095
      ScaleWidth      =   3975
      TabIndex        =   13
      Top             =   3960
      Width           =   3975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "CLOSE GRID"
      Height          =   495
      Left            =   9720
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "ISSUE UPDATION.frx":A5D2
      Height          =   1935
      Left            =   5640
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3413
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   855
      Left            =   600
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      RecordSource    =   "SELECT* FROM issue;"
      Caption         =   "Adodc3"
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
      DataSource      =   "adodc1"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FIND"
      Height          =   495
      Left            =   8520
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "REFRESH"
      Height          =   975
      Left            =   4080
      TabIndex        =   2
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PRINT"
      Height          =   975
      Left            =   2760
      TabIndex        =   3
      Top             =   9480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   855
      Left            =   1560
      Top             =   3480
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
      RecordSource    =   "select * from booklist;"
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
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   855
      Left            =   2280
      Top             =   4200
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ISSUE UPDATION.frx":A5E7
      Height          =   9015
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   15901
      _Version        =   393216
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   27
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
   Begin VB.CommandButton Command2 
      Caption         =   "MAIN MENU"
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REMOVE ISSUE"
      Height          =   975
      Left            =   1440
      TabIndex        =   4
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "                                        ISSUE LIST"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer

Dim c As Integer
Private Sub Command1_Click()

Dim a As String
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("DATABASE EMPTY")

Else
Text1.Text = Adodc1.Recordset!bookid
Text2.Text = Text1.Text
If Adodc2.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
GoTo ftr
Else




Adodc2.Recordset.MoveFirst


Adodc2.Recordset.FIND "bookid = '" & Text2.Text & "'", 0, adSearchForward
If Adodc2.Recordset.EOF = True Then
GoTo ftr
Else



n = Adodc2.Recordset!booksleft
n = n + 1
Adodc2.Recordset!booksleft = n
Adodc2.Recordset.Update
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("DATABASE EMPTY")
GoTo lst
End If
ftr:
Adodc1.Recordset.Delete
MsgBox ("ISSUE DELETED")
lst:
End If
End If
Adodc1.Refresh
Adodc2.Refresh
End If
End Sub

Private Sub Command2_Click()
Form5.Hide




Form1.Show


End Sub

Private Sub Command3_Click()
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("UNABLE TO PRINT DATABASE EMPTY")
GoTo SD
Else
Dim i As Integer
Dim str As String
Dim n As Integer
kt:


If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = False Then

GoTo fr:
Else
Printer.Print "ISSUE LIST"
Printer.Print "BOOKNAME            STUDENTNAME            DATE\TIME"
n = 0
st:
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = False Then
GoTo fr:
Else
str = ""
For i = 0 To DataGrid1.Columns.Count - 1

DataGrid1.Col = i
str = str + DataGrid1.Text + "         "


Next

Printer.Print str
n = n + 1

Adodc1.Recordset.MoveNext
If n = 50 Then
Printer.NewPage
GoTo kt
Else
GoTo st
End If
End If
End If
End If
fr:
Printer.EndDoc
SD:
End Sub

Private Sub Command4_Click()
Adodc1.Refresh
Adodc2.Refresh

End Sub

Private Sub Command9_Click()
DataGrid2.Visible = False
Command9.Visible = False
End Sub

Private Sub Form_Load()

If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("ISSUE LIST EMPTY")

End If


 counter = 1
    Picture1.Picture = LoadPicture("E:\software\images\image" & counter & ".jpg")
    Timer1.Interval = 5000 '5 seconds
    Timer1.Enabled = True
End Sub

Private Sub Text3_Change()
Adodc1.Refresh

Adodc3.Refresh
Adodc3.RecordSource = "Select * from issue where studentname Like '" & Text3.Text & "%' OR Bookid Like '" & Text3.Text & "%' "
DataGrid2.Refresh
If Adodc3.Recordset.EOF = True Then
GoTo fr
Else
DataGrid2.Visible = True

DataGrid2.Refresh
Dim fnd As String
fnd = Adodc3.Recordset!bookid

Adodc1.Recordset.FIND "bookid = '" + fnd + "'", 0, adSearchForward

DataGrid1.Refresh
End If
fr:
If DataGrid2.Visible = True Then
Command9.Visible = True
End If
End Sub

Private Sub Text3_Click()
Text3.Text = ""

End Sub
Private Sub Timer1_Timer()
    counter = counter + 1
    If counter > 6 Then counter = 1
    Picture1.Picture = LoadPicture("E:\software\images\image" & counter & ".jpg")
End Sub

