VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form9 
   BackColor       =   &H80000007&
   Caption         =   "SETTINGS"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9705
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":038A
   ScaleHeight     =   3585
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   6960
      Top             =   1200
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
      RecordSource    =   "datagrid"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADD ACCOUNT IMAGE"
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      Top             =   2880
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   6840
      Top             =   6240
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
      RecordSource    =   "login"
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
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "1"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Tag             =   "0"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEW PIN"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NEW PIN"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NEWUSERNAME"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OLD PIN"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OLD USERNAME"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Private Sub Command1_Click()
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim g As String
Dim f As String
a = Text1.Text
b = Text2.Text
g = Text3.Text
f = Text4.Text

c = Adodc1.Recordset!Password
d = Adodc1.Recordset!UserName

If a = c And d = b Then

Adodc1.Recordset!Password = f
Adodc1.Recordset!UserName = g
Adodc1.Recordset.Update
MsgBox ("PIN CHANGED ")

Else
MsgBox ("OLD PASSWORD INCORRECT")
End If
Form9.Hide




End Sub

Private Sub Command2_Click()
Form9.Hide




Form1.Show

End Sub

Private Sub Command3_Click()
Dim ad As String
CommonDialog1.ShowOpen
ad = CommonDialog1.FileName
Text5.Text = ad

Adodc2.Recordset!ad = Text5.Text
Adodc2.Recordset.Update
MsgBox ("IMAGE UPDATED")


End Sub

