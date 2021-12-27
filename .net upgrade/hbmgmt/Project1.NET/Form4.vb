Option Strict Off
Option Explicit On
Friend Class Form4
	Inherits System.Windows.Forms.Form
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Adodc1.RecordSource = "Select * FROM products ;"
		Adodc1.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid1.CtlRefresh()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Adodc1.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid1.CtlRefresh()
		MsgBox("STOCK RELOADED")
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		Me.Hide()
		Form2.Show()
		
	End Sub
	
	Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("STOCK EMPTY")
		Else
			Adodc1.Recordset.Delete()
			MsgBox("STOCK DELETED")
		End If
	End Sub
	
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("STOCK EMPTY")
		Else
			Adodc1.Recordset.Update()
			MsgBox("STOCK UPDATED")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Text1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.TextChanged
		Adodc1.Refresh()
		Adodc1.RecordSource = "Select * FROM products where productname Like '" & Text1.Text & "%' "
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("ITEM NOT FOUND! TRY SOME OTHER KEYWORDS")
			Text1.Text = ""
			GoTo fr
		Else
			'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			DataGrid1.CtlRefresh()
			
		End If
fr: 
	End Sub
End Class