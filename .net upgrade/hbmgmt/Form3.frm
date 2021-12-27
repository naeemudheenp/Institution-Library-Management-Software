VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000A&
   Caption         =   "Form3"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5805
   LinkTopic       =   "Form3"
   ScaleHeight     =   2760
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   975
      Left            =   1920
      TabIndex        =   9
      Top             =   3120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1720
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD "
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox tprice 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox tpname 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox tid 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "PRICE"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "PRODUCT NAME"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "UNIQUE ID"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "                                                           STOCK"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If tid = "" Then
MsgBox ("UNIQUE ID FIELD IS IMPORTANT. IT SHOULD BE FILLED.")
GoTo stp
End If
Adodc1.RecordSource = "Select * FROM products where id = '" & tid.Text & "' ;"
DataGrid1.Refresh
Adodc1.Refresh
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.AddNew
Adodc1.Recordset!id = tid
Adodc1.Recordset!ProductName = tname
Adodc1.Recordset!price = tprice
Adodc1.Recordset.Update
tid.Text = ""
tname.Text = ""
tprice = ""
MsgBox ("ITEMS ADDED")
Else
MsgBox ("ITEM WITH THIS ID EXITS")
tid.Text = ""
tname.Text = ""
tprice = ""
End If
stp:
End Sub

Private Sub Command2_Click()
tid.Text = ""
tname.Text = ""
tprice = ""
End Sub

Private Sub Command3_Click()
Me.Hide
Form2.Show
End Sub
