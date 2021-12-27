Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	
	
	
	Private Sub Combo1_Change()
		
	End Sub
	
	Private Sub cmbsrc_Form1()
		
		
		
		
		
		
		
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Form2.Show()
		
	End Sub
	
	Private Sub Command10_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command10.Click
		Form4.Show()
		
	End Sub
	
	Private Sub Command11_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command11.Click
		Form8.Show()
		
		
		
		
	End Sub
	
	Private Sub Command12_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command12.Click
		Me.Hide()
		frmTip.Show()
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Form3.Show()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Me.Hide()
		Dialog.Show()
		
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		MsgBox("APP VERSION:1.0 CREATED BY NAEEMUDHEEN ASLAM ")
		
	End Sub
	
	Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
		Form9.Show()
		
	End Sub
	
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		Form5.Show()
		
	End Sub
	
	Private Sub Command7_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command7.Click
		Form11.Show()
		
		
		
		
	End Sub
	
	Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command9.Click
		DataGrid1.Visible = False
		Command9.Visible = False
		
		
		
		
		
	End Sub
	
	
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		If DataGrid1.Visible = True Then
			Command9.Visible = True
		End If
		
		
		
		
		
	End Sub
	
	'UPGRADE_WARNING: Event Text1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.TextChanged
		Adodc1.Refresh()
		Adodc1.RecordSource = "Select * from booklist where Bookname Like '" & Text1.Text & "%' OR Bookid Like '" & Text1.Text & "%' OR author Like '" & Text1.Text & "%'OR publication Like '" & Text1.Text & "%' "
		Adodc1.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid1.CtlRefresh()
		If Adodc1.Recordset.EOF = True Then
			GoTo fr
		Else
			DataGrid1.Visible = True
		End If
		Adodc1.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid1.CtlRefresh()
fr: 
		If DataGrid1.Visible = True Then
			Command9.Visible = True
		End If
	End Sub
	
	Private Sub Text1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Enter
		Text1.Text = ""
		
	End Sub
End Class