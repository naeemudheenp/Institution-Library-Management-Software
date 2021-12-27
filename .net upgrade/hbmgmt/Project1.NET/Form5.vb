Option Strict Off
Option Explicit On
Friend Class Form5
	Inherits System.Windows.Forms.Form
	Public Sub Calendar1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Calendar1.ClickEvent
		
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Form6.Show()
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Form8.Show()
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Form7.Show()
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		Me.Hide()
		Form2.Show()
		
	End Sub
	
	Public Sub Form5_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
	End Sub
End Class