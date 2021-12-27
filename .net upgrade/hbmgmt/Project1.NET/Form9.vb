Option Strict Off
Option Explicit On
Friend Class Form9
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim p As String
		Dim u As String
		u = Text1.Text
		p = Text2.Text
		
		Adodc1.Recordset.Update()
		
		Adodc1.Recordset.Fields("UserName").Value = u
		Adodc1.Recordset.Fields("Password").Value = p
		MsgBox("    PASSWORD CHANGED")
		Text1.Text = ""
		Text2.Text = ""
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Me.Hide()
		Form2.Show()
		
	End Sub
End Class