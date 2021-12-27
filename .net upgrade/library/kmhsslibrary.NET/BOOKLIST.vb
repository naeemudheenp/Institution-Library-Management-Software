Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class Form3
	Inherits System.Windows.Forms.Form
	Dim c As Short
	
	Public Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("DATABASE EMPTY")
			GoTo lst
		End If
		
		Adodc2.Refresh()
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("DATABASE EMPTY")
			GoTo lst
		End If
		
		Text1.Text = Adodc1.Recordset.Fields("bookid").Value
		Text2.Text = Text1.Text
		
		
		Adodc1.Recordset.Delete()
		MsgBox("BOOK DELETED")
		Adodc1.Refresh()
		'DELETING FROM ISSUE LIST'
mrt: 
		If Adodc2.Recordset.EOF = False And Adodc2.Recordset.BOF = False Then
			
			Adodc2.Recordset.FIND("bookid = '" & Text2.Text & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
			
			If Adodc2.Recordset.EOF = False Then
				Adodc2.Recordset.Delete()
			End If
			If Adodc2.Recordset.BOF = False And Adodc2.Recordset.EOF = False Then
				Adodc2.Refresh()
				
				GoTo mrt
			End If
		End If
		
		
		
		
		
lst: 
	End Sub
	
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("DATABASE EMPTY")
			GoTo lst
		End If
		
		
		
		
		
		
		Adodc1.Recordset.Update()
		Adodc1.Recordset.MoveNext()
		MsgBox("BOOK INFO UPDATED")
lst: 
		Adodc1.Refresh()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		
		
		
		
		Adodc1.Refresh()
		
		
		
		
	End Sub
	
	
	Public Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		Dim Printer As New Printer
		Adodc1.Refresh()
		
		
		
		
		
		
		Dim i As Short
		Dim n As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("UNABLE TO PRINT DATABASE EMPTY")
			GoTo SD
		Else
kt: 
			n = 0
			
			If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
				
				GoTo fr
			Else
				
				Printer.Print("BOOKLIST")
				Printer.Print("BOOKID            BOOKNAME            AUTHOR            PUBLICATION            QTY            BOOKSLEFT            SHELF            RACK            PRICE")
				n = 0
st: 
				If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = False Then
					GoTo fr
					
				Else
					str_Renamed = ""
					For i = 0 To DataGrid1.Columns.Count - 1
						
						DataGrid1.Col = i
						str_Renamed = str_Renamed & DataGrid1.Text & "         "
						
						
					Next 
					
					Printer.Print(str_Renamed)
					n = n + 1
					
					Adodc1.Recordset.MoveNext()
					If n = 50 Then
						Printer.NewPage()
						GoTo kt
					Else
						GoTo st
					End If
				End If
			End If
fr: 
			Printer.EndDoc()
		End If
SD: 
		
	End Sub
	
	
	
	
	
	Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
		
		Me.Hide()
		
		If c = 1 Then
			
			c = 0
		Else
			
			c = 1
			
		End If
		
		Form1.Show()
	End Sub
	
	Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command9.Click
		DataGrid2.Visible = False
		Command9.Visible = False
	End Sub
	
	Private Sub Form3_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Adodc1.Refresh()
		If Adodc1.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
			MsgBox("BOOKLIST EMPTY.PLEASE ADD A BOOK")
			Me.Hide()
			
		End If
		
		
		
		
		
	End Sub
	
	'UPGRADE_WARNING: Event Text3.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text3_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text3.TextChanged
		Adodc1.Refresh()
		
		
		Adodc3.Refresh()
		Adodc3.RecordSource = "Select * from booklist where Bookname Like '" & Text3.Text & "%' OR Bookid Like '" & Text3.Text & "%' OR author Like '" & Text3.Text & "%'OR publication Like '" & Text3.Text & "%' "
		Adodc3.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid2.CtlRefresh()
		If Adodc3.Recordset.EOF = True Then
			GoTo fr
		Else
			DataGrid2.Visible = True
		End If
		Adodc3.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid2.CtlRefresh()
		Dim fnd As String
		fnd = Adodc3.Recordset.Fields("bookid").Value
		
		Adodc1.Recordset.FIND("bookid = '" & fnd & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
		
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid1.CtlRefresh()
		
		
		
		
		
fr: 
		If DataGrid2.Visible = True Then
			Command9.Visible = True
		End If
	End Sub
	
	Private Sub Text3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text3.Click
		Text3.Text = ""
	End Sub
End Class