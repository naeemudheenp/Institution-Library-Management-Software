VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   LinkTopic       =   "Form9"
   ScaleHeight     =   1950
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "MAIN MENU"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "select * from security;"
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
   Begin VB.CommandButton Command1 
      Caption         =   "CHANGE"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "ENTER NEW PASSWORD"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "ENTER NEW USERNAME"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "                                                         CHANGE PASSWORD"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim p As String
Dim u As String
u = Text1.Text
p = Text2.Text

Adodc1.Recordset.Update

Adodc1.Recordset!UserName = u
Adodc1.Recordset!Password = p
MsgBox ("    PASSWORD CHANGED")
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command2_Click()
Me.Hide
Form2.Show

End Sub
