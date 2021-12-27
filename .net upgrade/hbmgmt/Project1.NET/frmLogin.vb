Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmLogin
	Inherits System.Windows.Forms.Form
	
	Public LoginSucceeded As Boolean
	
	Private Sub cmdCancel_Click()
		'set the global var to false
		'to denote a failed login
		LoginSucceeded = False
		Me.Hide()
		End
	End Sub
	
	Private Sub cmdOK_Click()
		'check for correct password
		Dim vus As String
		Dim vps As String
		vus = Adodc1.Recordset.Fields("UserName").Value
		vps = Adodc1.Recordset.Fields("Password").Value
		If txtPassword.Text = vps And txtusername.Text = vus Then
			'place code to here to pass the
			'success to the calling sub
			'setting a global var is the easiest
			MsgBox("LOGGED IN")
			LoginSucceeded = True
			Me.Hide()
		Else
			MsgBox("Invalid Password, try again!")
			
		End If
	End Sub
	
	Private Sub Picture1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Picture1.Click
		'check for correct password
		Dim vus As String
		Dim vps As String
		vus = Adodc1.Recordset.Fields("UserName").Value
		vps = Adodc1.Recordset.Fields("Password").Value
		If txtPassword.Text = vps And txtusername.Text = vus Then
			'place code to here to pass the
			'success to the calling sub
			'setting a global var is the easiest
			MsgBox("LOGGED IN")
			LoginSucceeded = True
			Me.Hide()
			Form2.Show()
		Else
			MsgBox("Invalid Password, try again!")
			
		End If
	End Sub
	
	Private Sub Picture2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Picture2.Click
		End
	End Sub
	
	Private Sub txtusername_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtusername.Enter
		txtusername.Text = ""
		txtPassword.Text = ""
	End Sub
End Class