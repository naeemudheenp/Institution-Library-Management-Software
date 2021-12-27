VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H80000010&
   Caption         =   "STOCK MANAGEMENT"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12315
   LinkTopic       =   "Form4"
   ScaleHeight     =   8610
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "MAIN MENU"
      Height          =   495
      Left            =   10440
      TabIndex        =   9
      Top             =   2400
      Width           =   1830
   End
   Begin VB.CommandButton Command6 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   10440
      TabIndex        =   8
      Top             =   1440
      Width           =   1830
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      Caption         =   "DELETE"
      Height          =   495
      Left            =   10440
      TabIndex        =   7
      Top             =   1920
      Width           =   1830
   End
   Begin VB.CommandButton Command3 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   10440
      TabIndex        =   5
      Top             =   960
      Width           =   1830
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      Height          =   495
      Left            =   9360
      TabIndex        =   4
      Top             =   480
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEARCH"
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   8295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":0000
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   12938
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "STOCK MANAGER"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "id"
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
         DataField       =   "price"
         Caption         =   "PRODUCT PRICE"
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
            DividerStyle    =   3
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   8640
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
      RecordSource    =   "SELECT * FROM products;"
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
   Begin VB.Label Label2 
      Caption         =   "       STOCK OPTIONS"
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   $"Form4.frx":0015
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Adodc1.RecordSource = "Select * FROM products ;"
Adodc1.Refresh
DataGrid1.Refresh

End Sub

Private Sub Command3_Click()
Adodc1.Refresh
DataGrid1.Refresh
MsgBox ("STOCK RELOADED")
End Sub

Private Sub Command4_Click()
Me.Hide
Form2.Show

End Sub

Private Sub Command5_Click()
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("STOCK EMPTY")
Else
Adodc1.Recordset.Delete
MsgBox ("STOCK DELETED")
End If
End Sub

Private Sub Command6_Click()
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("STOCK EMPTY")
Else
Adodc1.Recordset.Update
MsgBox ("STOCK UPDATED")
End If
End Sub

Private Sub Text1_Change()
Adodc1.Refresh
Adodc1.RecordSource = "Select * FROM products where productname Like '" & Text1.Text & "%' "
If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
MsgBox ("ITEM NOT FOUND! TRY SOME OTHER KEYWORDS")
Text1.Text = ""
GoTo fr
Else
DataGrid1.Refresh

End If
fr:
End Sub
