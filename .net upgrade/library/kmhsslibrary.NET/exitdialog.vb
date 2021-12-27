Option Strict Off
Option Explicit On
Friend Class Dialog
	Inherits System.Windows.Forms.Form
	
	
	Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
		Me.Hide()
		Form1.Show()
		
		
	End Sub
	
	Private Sub OKButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OKButton.Click
		
		
		
		
		
		End
		
		
	End Sub
End Class