VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14055
   LinkTopic       =   "Form6"
   ScaleHeight     =   7695
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "MAINMENU"
      Height          =   735
      Left            =   11640
      TabIndex        =   9
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   11640
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   11640
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "STOCK.frx":0000
      Height          =   6855
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "STOCK REPORT"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "PRODUCT UNIQUE ID"
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
         Caption         =   "TOTAL"
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
         DataField       =   "time"
         Caption         =   "DATE"
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
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1995.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "            FINAL REPORT"
      Height          =   495
      Left            =   11760
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "          TOTAL QUANTITY"
      Height          =   615
      Left            =   11760
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "           TOTAL AMOUNT "
      Height          =   615
      Left            =   11760
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "SELECTED DATE:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "                                                                     REPORT BEFORE SPECIFIED DATE"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13935
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide

Form2.Show

End Sub

Private Sub Form_Load()

Text1.Text = Form5.Calendar1.Value
Adodc1.RecordSource = "Select * FROM sales where time > '" & Text1.Text & "' "
Adodc1.Refresh
DataGrid1.Refresh
Dim psum As Integer
Dim pqty As Integer
Dim val As String
Dim qval As String
psum = 0
pqty = 0
For i = 0 To DataGrid1.ApproxCount - 1
val = Adodc1.Recordset!price
qval = Adodc1.Recordset!qty
psum = val + psum
pqty = qval + pqty
If Adodc1.Recordset.EOF = True Then
GoTo fr
End If
Adodc1.Recordset.MoveNext
Next
fr:
Text2.Text = psum
Text3.Text = pqty
End Sub
