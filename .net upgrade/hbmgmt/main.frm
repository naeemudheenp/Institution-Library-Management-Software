VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "MAIN MENU"
      Height          =   1095
      Left            =   9000
      TabIndex        =   20
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HIDE"
      Height          =   255
      Left            =   10200
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   735
      Left            =   9000
      TabIndex        =   19
      Text            =   "0"
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox Text10 
      Height          =   735
      Left            =   9000
      TabIndex        =   16
      Text            =   "0"
      Top             =   5640
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      DataSource      =   "Adodc3"
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Text            =   "Text9"
      Top             =   8880
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      DataSource      =   "Adodc3"
      Height          =   285
      Left            =   3600
      TabIndex        =   12
      Text            =   "Text8"
      Top             =   9120
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2880
      Top             =   10080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from sales;"
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
   Begin VB.CommandButton Command2 
      Caption         =   "PURCHASE"
      Height          =   615
      Left            =   9000
      TabIndex        =   4
      Top             =   8040
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      Height          =   735
      Left            =   9000
      TabIndex        =   11
      Text            =   "0"
      Top             =   7320
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc3"
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc3"
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   360
      Top             =   9120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from products;"
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
      Bindings        =   "main.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         DataField       =   "id"
         Caption         =   "UNIQUE ID"
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
         DataField       =   "productname"
         Caption         =   "PRODUCT NAME"
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
         DataField       =   "price"
         Caption         =   "PRICE"
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
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1395.213
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2400
      Top             =   9000
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ontime"
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
      Bindings        =   "main.frx":0015
      Height          =   6375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
      Caption         =   "SALES"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "UNIQUE ID"
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
         DataField       =   "productname"
         Caption         =   "PRODUCT NAME"
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
         DataField       =   "qty"
         Caption         =   "QUANTITY"
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
         DataField       =   "price"
         Caption         =   "PRICE"
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
         DataField       =   "total"
         Caption         =   "TOTAL AMOUNT"
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
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   4
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            DividerStyle    =   4
            ColumnWidth     =   1800
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   8880
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENTER"
      Height          =   375
      Left            =   10200
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   8535
   End
   Begin VB.Label Label4 
      Caption         =   "TOTAL QUANTITY"
      Height          =   735
      Left            =   9000
      TabIndex        =   18
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "TOTAL AMOUNT"
      Height          =   735
      Left            =   9000
      TabIndex        =   17
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "FINAL AMOUNT"
      Height          =   735
      Left            =   9000
      TabIndex        =   15
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "                                                                                                                     SALES WINDOW"
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public total As Integer
Private Sub Command1_Click()
If Text1.Text = "" And Text2.Text = "" Then
MsgBox ("NO ITEMS LISTED.")
GoTo dd

Else
Dim val1 As Integer
Dim val2 As Integer
Dim qt As Integer
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("NO ITEMS LISTED.")
GoTo dd
End If

Text3.Text = Adodc2.Recordset!id
Text4.Text = Adodc2.Recordset!ProductName
Text6.Text = Adodc2.Recordset!price
Adodc1.Recordset.AddNew
Adodc1.Recordset!id = Text3.Text
Adodc1.Recordset!ProductName = Text4.Text
Adodc1.Recordset!qty = Text2.Text

Adodc1.Recordset!price = Text6.Text
qt = Text2.Text + qt
Text11.Text = qt
val1 = Text2.Text
val2 = Text6.Text
val2 = val1 * val2
total = val2 + total
Text10.Text = total
Adodc1.Recordset!total = val2
Adodc1.Recordset.Update
Adodc1.Refresh
DataGrid1.Refresh
Text7.Text = total
Text1.Text = ""
Text2.Text = ""
DataGrid2.Visible = False
End If




dd:
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim str As String
If Adodc1.Recordset.EOF And Adodc1.Recordset.BOF = True Then
MsgBox ("NO PURCHASE MADE")
GoTo LTT
Else

Printer.Print "WELCOME TO HOTEL"
Printer.Print DateTime.Now

Printer.Print "BILL"

Printer.Print "---------"

fr:
If Adodc1.Recordset.EOF = True Then
GoTo lt
End If
str = ""
For i = 0 To DataGrid1.Columns.Count - 1
DataGrid1.Col = i
str = str + DataGrid1.Text + "         "
Next

Printer.Print str
If Adodc1.Recordset.EOF = True Then
GoTo lt
Else
Adodc1.Recordset.MoveNext
GoTo fr
End If
lt:
Dim tt As String
Dim v As String
v = Text7.Text
tt = "TOTAL:" + v
Printer.Print tt

Printer.Print "HOPE YOU ENJOYED THE FOOD . COME AGAIN "
Printer.Print "MACBETH CORP"
Printer.EndDoc
End If
'senting data
snt:
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("NO PURCHASE MADE")
GoTo LTT
End If
Adodc1.Recordset.MoveFirst
Adodc1.Refresh
If Adodc1.Recordset.EOF = True Then
GoTo ed
Else
Adodc1.Recordset.Update
Text3.Text = Adodc1.Recordset!id
Text4.Text = Adodc1.Recordset!ProductName
Text6.Text = Adodc1.Recordset!price
Text8.Text = Adodc1.Recordset!qty
Text9.Text = Adodc1.Recordset!total
Adodc1.Recordset.Delete
Adodc3.Recordset.AddNew


Adodc3.Recordset!id = Text3.Text
Adodc3.Recordset!ProductName = Text4.Text
Adodc3.Recordset!qty = Text8.Text
Adodc3.Recordset!price = Text6.Text
Adodc3.Recordset!total = Text9.Text
Adodc3.Recordset!Time = DateTime.Date
Adodc3.Recordset.Update

GoTo snt:
End If
ed:
Adodc1.Refresh

DataGrid1.Refresh

MsgBox ("PURCHASE COMPLETED")
Text7.Text = ""
Text11.Text = ""
Text10.Text = ""
LTT:


End Sub


Private Sub Command3_Click()
If DataGrid2.Visible = True Then
DataGrid2.Visible = False
Else
MsgBox ("SEARCH GRID HIDDEN !")
End If

End Sub

Private Sub Command4_Click()
Form2.Show
End Sub

Private Sub DataGrid2_Click()
If Adodc2.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
Else
Text1.Text = Adodc2.Recordset!ProductName
End If

End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
GoTo fr
End If
Adodc2.Refresh
Adodc2.RecordSource = "Select * FROM products where productname Like '" & Text1.Text & "%' "
If Adodc2.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then

GoTo fr
End If
Adodc2.Refresh
DataGrid2.Refresh
If Adodc2.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
GoTo fr
MsgBox ("ITEM NOT FOUND! TRY SOME OTHER KEYWORDS")
Text1.Text = ""

Else
DataGrid2.Visible = True
End If
Adodc2.Refresh
DataGrid2.Refresh
fr:
End Sub

Private Sub Text1_Click()
Text1.Text = ""
Text2.Text = ""


End Sub
