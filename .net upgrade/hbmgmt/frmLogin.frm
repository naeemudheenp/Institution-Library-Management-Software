VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   3075
   ClientLeft      =   2790
   ClientTop       =   3090
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1816.812
   ScaleMode       =   0  'User
   ScaleWidth      =   2915.427
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   170
      Left            =   2760
      Picture         =   "frmLogin.frx":0000
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   4
      Top             =   120
      Width           =   165
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   7  'Diagonal Cross
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      Picture         =   "frmLogin.frx":2B93
      ScaleHeight     =   375
      ScaleWidth      =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   1087
   End
   Begin VB.TextBox txtusername 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000A&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "                 username"
      Top             =   840
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3360
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.TextBox txtPassword 
      DataSource      =   "Adodc1"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "password"
      Top             =   1485
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "                SIGN IN"
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    End
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    Dim vus As String
    Dim vps As String
    vus = Adodc1.Recordset!UserName
    vps = Adodc1.Recordset!Password
    If txtPassword = vps And txtusername = vus Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        MsgBox ("LOGGED IN")
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!"
        
    End If
End Sub

Private Sub Picture1_Click()
   'check for correct password
    Dim vus As String
    Dim vps As String
    vus = Adodc1.Recordset!UserName
    vps = Adodc1.Recordset!Password
    If txtPassword = vps And txtusername = vus Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        MsgBox ("LOGGED IN")
        LoginSucceeded = True
        Me.Hide
        Form2.Show
    Else
        MsgBox "Invalid Password, try again!"
        
    End If
End Sub

Private Sub Picture2_Click()
End
End Sub

Private Sub txtusername_GotFocus()
txtusername.Text = ""
txtPassword.Text = ""
End Sub
