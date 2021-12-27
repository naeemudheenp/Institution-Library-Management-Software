VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "BOOKLIST"
   ClientHeight    =   9300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16545
   Icon            =   "BOOKLIST.frx":0000
   LinkTopic       =   "Form3"
   Picture         =   "BOOKLIST.frx":038A
   ScaleHeight     =   9300
   ScaleWidth      =   16545
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   735
      Left            =   9000
      Top             =   4440
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
      RecordSource    =   "SELECT* FROM BOOKLIST;"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "BOOKLIST.frx":1D84B
      Height          =   1215
      Left            =   10920
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2143
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
   Begin VB.CommandButton Command9 
      Caption         =   "CLOSE GRID"
      Height          =   615
      Left            =   15000
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "FIND"
      Height          =   615
      Left            =   14160
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   10920
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   7200
      Top             =   4320
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "BOOKLIST.frx":1D860
      Height          =   8295
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   14631
      _Version        =   393216
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   4320
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      RecordSource    =   "booklist"
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
   Begin VB.CommandButton Command5 
      Caption         =   "MAIN MENU"
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "REFRESH"
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   8880
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   $"BOOKLIST.frx":1D875
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer

Public Sub Command1_Click()
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("DATABASE EMPTY")
GoTo lst
End If

Adodc2.Refresh
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("DATABASE EMPTY")
GoTo lst
End If

Text1.Text = Adodc1.Recordset!bookid
Text2.Text = Text1.Text


Adodc1.Recordset.Delete
MsgBox ("BOOK DELETED")
Adodc1.Refresh
'DELETING FROM ISSUE LIST'
mrt:
If Adodc2.Recordset.EOF = False And Adodc2.Recordset.BOF = False Then

Adodc2.Recordset.FIND "bookid = '" & Text2.Text & "'", 0, adSearchForward

If Adodc2.Recordset.EOF = False Then
Adodc2.Recordset.Delete
End If
If Adodc2.Recordset.BOF = False And Adodc2.Recordset.EOF = False Then
Adodc2.Refresh

GoTo mrt
End If
End If





lst:
End Sub


Private Sub Command2_Click()
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("DATABASE EMPTY")
GoTo lst
End If






Adodc1.Recordset.Update
Adodc1.Recordset.MoveNext
MsgBox ("BOOK INFO UPDATED")
lst:
Adodc1.Refresh

End Sub

Private Sub Command3_Click()




Adodc1.Refresh




End Sub


Public Sub Command4_Click()
Adodc1.Refresh






If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("UNABLE TO PRINT DATABASE EMPTY")
GoTo SD
Else
Dim i As Integer
Dim str As String
Dim n As Integer
kt:
n = 0

If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then

GoTo fr:
Else

Printer.Print "BOOKLIST"
Printer.Print "BOOKID            BOOKNAME            AUTHOR            PUBLICATION            QTY            BOOKSLEFT            SHELF            RACK            PRICE"
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
fr:
Printer.EndDoc
End If
SD:

End Sub





Private Sub Command5_Click()

Form3.Hide

If c = 1 Then

c = 0
Else

c = 1

End If

Form1.Show
End Sub

Private Sub Command9_Click()
DataGrid2.Visible = False
Command9.Visible = False
End Sub

Private Sub Form_Load()
Adodc1.Refresh
If Adodc1.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
MsgBox ("BOOKLIST EMPTY.PLEASE ADD A BOOK")
Form3.Hide

End If





End Sub

Private Sub Text3_Change()
Adodc1.Refresh


Adodc3.Refresh
Adodc3.RecordSource = "Select * from booklist where Bookname Like '" & Text3.Text & "%' OR Bookid Like '" & Text3.Text & "%' OR author Like '" & Text3.Text & "%'OR publication Like '" & Text3.Text & "%' "
Adodc3.Refresh
DataGrid2.Refresh
If Adodc3.Recordset.EOF = True Then
GoTo fr
Else
DataGrid2.Visible = True
End If
Adodc3.Refresh
DataGrid2.Refresh
Dim fnd As String
fnd = Adodc3.Recordset!bookid

Adodc1.Recordset.FIND "bookid = '" + fnd + "'", 0, adSearchForward

DataGrid1.Refresh





fr:
If DataGrid2.Visible = True Then
Command9.Visible = True
End If
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

