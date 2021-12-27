Option Strict Off
Option Explicit On
Friend Class frmLogin
	Inherits System.Windows.Forms.Form
	
	Private Sub frmlogin_load()
		
	End Sub
	Private Sub cmdOK_Click()
		
	End Sub
	
	
	
	Private Sub Button1_Click()
		
		
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim a As String
		Dim b As String
		Dim X As String
		X = "login"
		a = u1.Text
		b = u2.Text
		Dim c As String
		Dim d As String
		
		
		
		c = Adodc1.Recordset.Fields("Password").Value
		d = Adodc1.Recordset.Fields("UserName").Value
		
		If a = d And b = c Then
			
			
			Me.Hide()
			
			Form1.Show()
			
			
			
		Else
			MsgBox(" PASSWORD INCORRECT")
		End If
	End Sub
	
	Private Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Text1.Text = Adodc2.Recordset.Fields("ad").Value
		Picture1.Image = System.Drawing.Image.FromFile(Text1.Text)
		
		
		
	End Sub
End Class