Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class Form1
	Inherits System.Windows.Forms.Form
	Public total As Short
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim val1 As Short
		Dim val2 As Short
		Dim qt As Short
		If Text1.Text = "" And Text2.Text = "" Then
			MsgBox("NO ITEMS LISTED.")
			GoTo dd
			
		Else
			If Text1.Text = "" Or Text2.Text = "" Then
				MsgBox("NO ITEMS LISTED.")
				GoTo dd
			End If
			
			Text3.Text = Adodc2.Recordset.Fields("id").Value
			Text4.Text = Adodc2.Recordset.Fields("ProductName").Value
			Text6.Text = Adodc2.Recordset.Fields("price").Value
			Adodc1.Recordset.AddNew()
			Adodc1.Recordset.Fields("id").Value = Text3.Text
			Adodc1.Recordset.Fields("ProductName").Value = Text4.Text
			Adodc1.Recordset.Fields("qty").Value = Text2.Text
			
			Adodc1.Recordset.Fields("price").Value = Text6.Text
			qt = CDbl(Text2.Text) + qt
			Text11.Text = CStr(qt)
			val1 = CShort(Text2.Text)
			val2 = CShort(Text6.Text)
			val2 = val1 * val2
			total = val2 + total
			Text10.Text = CStr(total)
			Adodc1.Recordset.Fields("total").Value = val2
			Adodc1.Recordset.Update()
			Adodc1.Refresh()
			'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			DataGrid1.CtlRefresh()
			Text7.Text = CStr(total)
			Text1.Text = ""
			Text2.Text = ""
			DataGrid2.Visible = False
		End If
		
		
		
		
dd: 
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Dim Printer As New Printer
		Dim i As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim tt As String
		Dim v As String
		If Adodc1.Recordset.EOF And Adodc1.Recordset.BOF = True Then
			MsgBox("NO PURCHASE MADE")
			GoTo LTT
		Else
			
			Printer.Print("WELCOME TO HOTEL")
			Printer.Print(CStr(Now))
			
			Printer.Print("BILL")
			
			Printer.Print("---------")
			
fr: 
			If Adodc1.Recordset.EOF = True Then
				GoTo lt
			End If
			str_Renamed = ""
			For i = 0 To DataGrid1.Columns.Count - 1
				DataGrid1.Col = i
				str_Renamed = str_Renamed & DataGrid1.Text & "         "
			Next 
			
			Printer.Print(str_Renamed)
			If Adodc1.Recordset.EOF = True Then
				GoTo lt
			Else
				Adodc1.Recordset.MoveNext()
				GoTo fr
			End If
lt: 
			v = Text7.Text
			tt = "TOTAL:" & v
			Printer.Print(tt)
			
			Printer.Print("HOPE YOU ENJOYED THE FOOD . COME AGAIN ")
			Printer.Print("MACBETH CORP")
			Printer.EndDoc()
		End If
		'senting data
snt: 
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("NO PURCHASE MADE")
			GoTo LTT
		End If
		Adodc1.Recordset.MoveFirst()
		Adodc1.Refresh()
		If Adodc1.Recordset.EOF = True Then
			GoTo ed
		Else
			Adodc1.Recordset.Update()
			Text3.Text = Adodc1.Recordset.Fields("id").Value
			Text4.Text = Adodc1.Recordset.Fields("ProductName").Value
			Text6.Text = Adodc1.Recordset.Fields("price").Value
			Text8.Text = Adodc1.Recordset.Fields("qty").Value
			Text9.Text = Adodc1.Recordset.Fields("total").Value
			Adodc1.Recordset.Delete()
			Adodc3.Recordset.AddNew()
			
			
			Adodc3.Recordset.Fields("id").Value = Text3.Text
			Adodc3.Recordset.Fields("ProductName").Value = Text4.Text
			Adodc3.Recordset.Fields("qty").Value = Text8.Text
			Adodc3.Recordset.Fields("price").Value = Text6.Text
			Adodc3.Recordset.Fields("total").Value = Text9.Text
			Adodc3.Recordset.Fields("Time").Value = Today
			Adodc3.Recordset.Update()
			
			GoTo snt
		End If
ed: 
		Adodc1.Refresh()
		
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid1.CtlRefresh()
		
		MsgBox("PURCHASE COMPLETED")
		Text7.Text = ""
		Text11.Text = ""
		Text10.Text = ""
LTT: 
		
		
	End Sub
	
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		If DataGrid2.Visible = True Then
			DataGrid2.Visible = False
		Else
			MsgBox("SEARCH GRID HIDDEN !")
		End If
		
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		Form2.Show()
	End Sub
	
	Private Sub DataGrid2_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DataGrid2.ClickEvent
		If Adodc2.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
		Else
			Text1.Text = Adodc2.Recordset.Fields("ProductName").Value
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event Text1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.TextChanged
		If Text1.Text = "" Then
			GoTo fr
		End If
		Adodc2.Refresh()
		Adodc2.RecordSource = "Select * FROM products where productname Like '" & Text1.Text & "%' "
		If Adodc2.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
			
			GoTo fr
		End If
		Adodc2.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid2.CtlRefresh()
		If Adodc2.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
			GoTo fr
			MsgBox("ITEM NOT FOUND! TRY SOME OTHER KEYWORDS")
			Text1.Text = ""
			
		Else
			DataGrid2.Visible = True
		End If
		Adodc2.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid2.CtlRefresh()
fr: 
	End Sub
	
	Private Sub Text1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Click
		Text1.Text = ""
		Text2.Text = ""
		
		
	End Sub
End Class