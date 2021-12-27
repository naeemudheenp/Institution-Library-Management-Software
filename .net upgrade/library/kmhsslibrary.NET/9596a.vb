Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class Form11
	Inherits System.Windows.Forms.Form
	Function FIND() As String
		Adodc1.Refresh()
		
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("DATABASE EMPTY")
			GoTo fr
		Else
			
			
			
			
			
			Adodc1.Recordset.MoveFirst()
			Adodc1.Recordset.FIND("bookid = '" & Text1.Text & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
			If Adodc1.Recordset.EOF Then
				Adodc1.Recordset.MoveFirst()
				Adodc1.Recordset.FIND("bookname = '" & Text1.Text & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
			Else
				MsgBox("RECORD FOUND CRITERIA : BOOKID")
				DataGrid1.Visible = True
				GoTo fr
			End If
			If Adodc1.Recordset.EOF Then
				Adodc1.Recordset.MoveFirst()
				Adodc1.Recordset.FIND("author = '" & Text1.Text & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
			Else
				MsgBox("RECORD FOUND CRITERIA : BOOKNAME")
				DataGrid1.Visible = True
				GoTo fr
			End If
			If Adodc1.Recordset.EOF Then
				Adodc1.Recordset.MoveFirst()
				Adodc1.Recordset.FIND("publication = '" & Text1.Text & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
			Else
				MsgBox("RECORD FOUND CRITERIA : AUTHOR")
				DataGrid1.Visible = True
				GoTo fr
			End If
			If Adodc1.Recordset.EOF Then
				MsgBox("RECORD NOT FOUND")
			Else
				MsgBox("RECORD FOUND CRITERIA : PUBLICATION")
				DataGrid1.Visible = True
				GoTo fr
				
			End If
		End If
		
		
		
		
		
		
fr: 
		
	End Function
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Adodc1.Refresh()
		
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Me.Hide()
		Form1.Show()
		
	End Sub
	
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		Dim Printer As New Printer
		Dim c As Short
		c = 0
		Dim i As Short
		Dim n As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("UNABLE TO PRINT DATABASE EMPTY")
			GoTo SD
		Else
kt: 
			
			
			If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
				GoTo fr
			Else
				Printer.Print("BOOKLIST")
				Printer.Print("BOOKID            BOOKNAME            AUTHOR            PUBLICATION            QTY            BOOKSLEFT            SHELF            RACK            PRICE")
				n = 0
st: 
				If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
					GoTo fr
				Else
					str_Renamed = ""
					For i = 0 To DataGrid1.Columns.Count - 1
						
						DataGrid1.Col = i
						str_Renamed = str_Renamed & DataGrid1.Text & "         "
						c = c + 1
						If c = DataGrid1.ApproxCount Then
							GoTo fr
						End If
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
	
	Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command9.Click
		DataGrid2.Visible = False
		Command9.Visible = False
	End Sub
	
	'UPGRADE_WARNING: Event Text4.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text4_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text4.TextChanged
		Adodc1.Refresh()
		
		Adodc2.Refresh()
		Adodc2.RecordSource = "Select * from booklist where Bookname Like '" & Text4.Text & "%' OR Bookid Like '" & Text4.Text & "%' OR author Like '" & Text4.Text & "%'OR publication Like '" & Text4.Text & "%' "
		Adodc2.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid2.CtlRefresh()
		If Adodc2.Recordset.EOF = True Then
			GoTo fr
		Else
			DataGrid2.Visible = True
		End If
		Adodc2.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid2.CtlRefresh()
		Dim fnd As String
		fnd = Adodc2.Recordset.Fields("bookid").Value
		
		Adodc1.Recordset.FIND("bookid = '" & fnd & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
		
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid1.CtlRefresh()
fr: 
		If DataGrid2.Visible = True Then
			Command9.Visible = True
		End If
	End Sub
	
	Private Sub Text4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text4.Click
		Text1.Text = ""
	End Sub
End Class