VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "ADD BOOK"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14040
   Icon            =   "ADDBOOK.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "ADDBOOK.frx":038A
   ScaleHeight     =   5265
   ScaleWidth      =   14040
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   12840
      Top             =   3840
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ADDBOOK.frx":90FE
      Height          =   735
      Left            =   4560
      TabIndex        =   19
      Top             =   120
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   12000
      Top             =   6240
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
   Begin VB.TextBox Text8 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2520
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MAIN MENU"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "ADD"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2520
      TabIndex        =   6
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RACK"
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PUBLICATION"
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SHELF NUMBER"
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BOOKID"
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    QUANTITY"
      Height          =   375
      Left            =   -120
      TabIndex        =   14
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "                        ADD BOOKS"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   PRICE"
      Height          =   375
      Left            =   -120
      TabIndex        =   12
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   AUTHOR"
      Height          =   375
      Left            =   -120
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "                        BOOK NAME "
      Height          =   375
      Left            =   -1080
      TabIndex        =   7
      Top             =   1560
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer

Dim c As Integer



Public Sub Command1_Click()

Dim n As String

Dim value As String
Dim value1 As String
Dim value2 As String
Dim value3 As String
Dim value4 As String
Dim value5 As String
Dim value6 As String
Dim value7 As String
Dim limit As Integer
Dim h As String
value = Text5.Text
value1 = Text7.Text
value2 = Text6.Text
value3 = Text8.Text
value4 = Text4.Text
value5 = Text3.Text
value6 = Text2.Text
value7 = Text1.Text
Dim gl As Integer
Dim j As Integer
limit = DataGrid1.ApproxCount
j = 0
n = Text4.Text






If value1 = "" Or value2 = "" Or value3 = "" Or value4 = "" Or value5 = "" Or value6 = "" Or value7 = "" Or value = "" Then
MsgBox ("You Cannot Leave Textbox empty")
Else
For i = 0 To limit - 1
h = Adodc1.Recordset!bookid

If h = value Then
MsgBox ("Book With This Id Has Been Entered Already")
GoTo lst
Else
Adodc1.Recordset.MoveNext
End If
Next
Adodc1.Recordset.AddNew
Adodc1.Recordset!bookname = Text1.Text
Adodc1.Recordset!author = Text2.Text
Adodc1.Recordset!price = Text3.Text
Adodc1.Recordset!qty = Text4.Text
Adodc1.Recordset!bookid = Text5.Text
Adodc1.Recordset!publication = Text7.Text
Adodc1.Recordset!shelf = Text6.Text
Adodc1.Recordset!rack = Text8.Text
Adodc1.Recordset!booksleft = Text4.Text



Adodc1.Recordset.Update
MsgBox ("BOOK ADDED")
Adodc1.Refresh


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End If
lst:
Adodc1.Refresh

End Sub


Private Sub Command2_Click()
Form2.Hide




Form1.Show

End Sub


Private Sub Form_Load()
Adodc1.Refresh

 
End Sub

Private Sub Text1_LostFocus()
Dim too As String
too = Text1.Text
Text1.Text = ""
Text1.Text = too

End Sub

Private Sub Text2_LostFocus()
Dim too2 As String

too2 = Text2.Text
Text2.Text = ""
Text2.Text = too2
End Sub

Private Sub Text7_LostFocus()
Dim too1 As String

too1 = Text7.Text
Text7.Text = ""
Text7.Text = too1
End Sub

