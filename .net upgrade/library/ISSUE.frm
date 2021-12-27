VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   Caption         =   "ISSUE"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8040
   Icon            =   "ISSUE.frx":0000
   LinkTopic       =   "Form4"
   Picture         =   "ISSUE.frx":038A
   ScaleHeight     =   6270
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   240
      Top             =   3720
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   2160
      Top             =   4440
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
   Begin VB.CommandButton Command2 
      Caption         =   "MAIN MENU"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ISSUE"
      Height          =   495
      Left            =   3240
      Picture         =   "ISSUE.frx":A5D2
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT NAME"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BOOK CODE"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "                         ISSUE"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer

Private Sub Command1_Click()
Adodc1.Refresh
Adodc2.Refresh

Text4.Text = Text2.Text


Dim a As String
Dim k As Integer
Dim tb2 As Integer
Dim lm As Integer

Dim i As Integer

a = Text1.Text
Dim h As String
Dim n As Integer
Dim o As Integer
Dim j As Integer

If a = "" Then
MsgBox ("You Cannot Leave The Box Empty")
Else

Adodc2.Recordset.FIND "bookid = '" & Text4.Text & "'", 0, adSearchForward
If Adodc2.Recordset.EOF Then
MsgBox ("BOOK NOT FOUND WITH THIS ID")
Else
Dim f As String
f = Adodc2.Recordset!booksleft
f = f - 1
Adodc2.Recordset!booksleft = f
Adodc2.Recordset.Update


Adodc1.Recordset.AddNew
Adodc1.Recordset!studentname = Text1.Text
Adodc1.Recordset!bookid = Text2.Text
Adodc1.Recordset!Time = Text3.Text
Adodc1.Recordset.Update
MsgBox ("BOOK ISSUED")
Text1.Text = ""
Text2.Text = ""
Adodc2.Recordset.MoveFirst



End If
End If

End Sub

Private Sub Command2_Click()
Form4.Hide





Form1.Show
End Sub



Private Sub Form_Load()
Adodc1.Refresh

Dim dt As Date
Dim tm As Date

dt = DateTime.Date

dt = dt
Text3.Text = dt
End Sub

Private Sub Text1_LostFocus()
Dim too As String
too = Text1.Text
Text1.Text = ""
Text1.Text = too

End Sub
