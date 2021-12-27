VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form11 
   Caption         =   "NOT ISSUED BOOKS"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16245
   Icon            =   "9596a.frx":0000
   LinkTopic       =   "Form11"
   Picture         =   "9596a.frx":038A
   ScaleHeight     =   8355
   ScaleWidth      =   16245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "CLOSE GRID"
      Height          =   495
      Left            =   15000
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   7200
      Top             =   3840
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "9596a.frx":1D84B
      Height          =   855
      Left            =   9960
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1508
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
   Begin VB.CommandButton Command3 
      Caption         =   "FIND"
      Height          =   495
      Left            =   14160
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9960
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   -120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   9360
      Top             =   7800
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
      RecordSource    =   $"9596a.frx":1D860
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
      Bindings        =   "9596a.frx":1D96B
      Height          =   7935
      Left            =   0
      TabIndex        =   13
      Top             =   480
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   13996
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
      BeginProperty Column05 
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
   Begin VB.CommandButton Command2 
      Caption         =   "MAIN MENU"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      DataField       =   "bookid"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "bookid"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc3"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc3"
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc3"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc3"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox fuck 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7320
      ScaleHeight     =   675
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   11280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "                          NOT ISSUED BOOKS"
      Height          =   495
      Left            =   6720
      TabIndex        =   12
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function FIND() As String
Adodc1.Refresh

If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("DATABASE EMPTY")
GoTo fr
Else





Adodc1.Recordset.MoveFirst
Adodc1.Recordset.FIND "bookid = '" & Text1.Text & "'", 0, adSearchForward
  If Adodc1.Recordset.EOF Then
  Adodc1.Recordset.MoveFirst
  Adodc1.Recordset.FIND "bookname = '" & Text1.Text & "'", 0, adSearchForward
  Else
           MsgBox ("RECORD FOUND CRITERIA : BOOKID")
           DataGrid1.Visible = True
           GoTo fr
           End If
    If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.FIND "author = '" & Text1.Text & "'", 0, adSearchForward
    Else
           MsgBox ("RECORD FOUND CRITERIA : BOOKNAME")
           DataGrid1.Visible = True
           GoTo fr
           End If
      If Adodc1.Recordset.EOF Then
      Adodc1.Recordset.MoveFirst
      Adodc1.Recordset.FIND "publication = '" & Text1.Text & "'", 0, adSearchForward
         Else
           MsgBox ("RECORD FOUND CRITERIA : AUTHOR")
           DataGrid1.Visible = True
           GoTo fr
           End If
         If Adodc1.Recordset.EOF Then
         MsgBox ("RECORD NOT FOUND")
           Else
           MsgBox ("RECORD FOUND CRITERIA : PUBLICATION")
           DataGrid1.Visible = True
          GoTo fr
          
           End If
 End If






fr:

End Function

Private Sub Command1_Click()
Adodc1.Refresh


End Sub

Private Sub Command2_Click()
Form11.Hide
Form1.Show

End Sub


Private Sub Command4_Click()
Dim c As Integer
c = 0
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("UNABLE TO PRINT DATABASE EMPTY")
GoTo SD
Else
Dim i As Integer
Dim str As String
Dim n As Integer
kt:


If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
GoTo fr:
Else
Printer.Print "BOOKLIST"
Printer.Print "BOOKID            BOOKNAME            AUTHOR            PUBLICATION            QTY            BOOKSLEFT            SHELF            RACK            PRICE"
n = 0
st:
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
GoTo fr:
Else
str = ""
For i = 0 To DataGrid1.Columns.Count - 1

DataGrid1.Col = i
str = str + DataGrid1.Text + "         "
c = c + 1
If c = DataGrid1.ApproxCount Then
GoTo fr
End If
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

Private Sub Command9_Click()
DataGrid2.Visible = False
Command9.Visible = False
End Sub

Private Sub Text4_Change()
Adodc1.Refresh

Adodc2.Refresh
Adodc2.RecordSource = "Select * from booklist where Bookname Like '" & Text4.Text & "%' OR Bookid Like '" & Text4.Text & "%' OR author Like '" & Text4.Text & "%'OR publication Like '" & Text4.Text & "%' "
Adodc2.Refresh
DataGrid2.Refresh
If Adodc2.Recordset.EOF = True Then
GoTo fr
Else
DataGrid2.Visible = True
End If
Adodc2.Refresh
DataGrid2.Refresh
Dim fnd As String
fnd = Adodc2.Recordset!bookid

Adodc1.Recordset.FIND "bookid = '" + fnd + "'", 0, adSearchForward

DataGrid1.Refresh
fr:
If DataGrid2.Visible = True Then
Command9.Visible = True
End If
End Sub

Private Sub Text4_Click()
Text1.Text = ""
End Sub



