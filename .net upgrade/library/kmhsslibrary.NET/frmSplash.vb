Option Strict Off
Option Explicit On
Friend Class frmSplash
	Inherits System.Windows.Forms.Form
	Dim counter As Short
	
	
	Private Sub frmSplash_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Me.Hide()
		frmLogin.Show()
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub Frame1_Click()
		Me.Hide()
		frmLogin.Show()
		
	End Sub
End Class