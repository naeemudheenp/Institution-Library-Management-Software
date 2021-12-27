Option Strict Off
Option Explicit On
Friend Class Form3
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Dim tname As TextBox
		If tid.Text = "" Then
			MsgBox("UNIQUE ID FIELD IS IMPORTANT. IT SHOULD BE FILLED.")
			GoTo stp
		End If
		Adodc1.RecordSource = "Select * FROM products where id = '" & tid.Text & "' ;"
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid1.CtlRefresh()
		Adodc1.Refresh()
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			Adodc1.Recordset.AddNew()
			Adodc1.Recordset.Fields("id").Value = tid.Text
			'UPGRADE_WARNING: Couldn't resolve default property of object tname. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Adodc1.Recordset.Fields("ProductName").Value = tname.text
			Adodc1.Recordset.Fields("price").Value = tprice.Text
			Adodc1.Recordset.Update()
			tid.Text = ""
			'UPGRADE_WARNING: Couldn't resolve default property of object tname.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			tname.Text = ""
			tprice.Text = ""
			MsgBox("ITEMS ADDED")
		Else
			MsgBox("ITEM WITH THIS ID EXITS")
			tid.Text = ""
			'UPGRADE_WARNING: Couldn't resolve default property of object tname.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			tname.Text = ""
			tprice.Text = ""
		End If
stp: 
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Dim tname As Object
		tid.Text = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object tname.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tname.Text = ""
		tprice.Text = ""
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Me.Hide()
		Form2.Show()
	End Sub
End Class