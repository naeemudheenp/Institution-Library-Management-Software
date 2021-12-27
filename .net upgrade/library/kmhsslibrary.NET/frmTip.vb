Option Strict Off
Option Explicit On
Friend Class frmTip
	Inherits System.Windows.Forms.Form
	Dim counter As Short
	
	
	
	
	
	
	Private Sub frmTip_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Me.Hide()
		Form1.Show()
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmTip_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		counter = 1
		Picture1.Image = System.Drawing.Image.FromFile("E:\software\images\SCREENSAVER IMAGES\image" & counter & ".jpg")
		Timer1.Interval = 5000 '5 seconds
		Timer1.Enabled = True
		
	End Sub
	
	Public Sub DisplayCurrentTip()
		
		
	End Sub
	
	Private Sub Picture1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Picture1.Click
		Me.Hide()
		Form1.Show()
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		counter = counter + 1
		If counter > 10 Then counter = 1
		Picture1.Image = System.Drawing.Image.FromFile("E:\software\images\SCREENSAVER IMAGES\image" & counter & ".jpg")
	End Sub
End Class