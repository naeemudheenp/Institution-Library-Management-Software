Option Strict Off
Option Explicit On
Friend Class Form9
	Inherits System.Windows.Forms.Form
	Dim c As Short
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
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
		
		c = Adodc1.Recordset.Fields("Password").Value
		d = Adodc1.Recordset.Fields("UserName").Value
		
		If a = c And d = b Then
			
			Adodc1.Recordset.Fields("Password").Value = f
			Adodc1.Recordset.Fields("UserName").Value = g
			Adodc1.Recordset.Update()
			MsgBox("PIN CHANGED ")
			
		Else
			MsgBox("OLD PASSWORD INCORRECT")
		End If
		Me.Hide()
		
		
		
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Me.Hide()
		
		
		
		
		Form1.Show()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Dim ad As String
		CommonDialog1Open.ShowDialog()
		ad = CommonDialog1Open.FileName
		Text5.Text = ad
		
		Adodc2.Recordset.Fields("ad").Value = Text5.Text
		Adodc2.Recordset.Update()
		MsgBox("IMAGE UPDATED")
		
		
	End Sub
End Class