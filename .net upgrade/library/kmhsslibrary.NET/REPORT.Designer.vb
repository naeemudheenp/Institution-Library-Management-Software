<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form8
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		VB6_AddADODataBinding()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			VB6_RemoveADODataBinding()
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public PrintForm1 As Microsoft.VisualBasic.PowerPacks.Printing.PrintForm
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Text5 As System.Windows.Forms.TextBox
	Public WithEvents DataGrid2 As AxMSDataGridLib.AxDataGrid
	Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Adodc2 As VB6.ADODC
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents Text4 As System.Windows.Forms.TextBox
	Public WithEvents Text3 As System.Windows.Forms.TextBox
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form8))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.PrintForm1 = New Microsoft.VisualBasic.PowerPacks.Printing.PrintForm(Me)
		Me.Command1 = New System.Windows.Forms.Button
		Me.Text5 = New System.Windows.Forms.TextBox
		Me.DataGrid2 = New AxMSDataGridLib.AxDataGrid
		Me.DataGrid1 = New AxMSDataGridLib.AxDataGrid
		Me.Adodc2 = New VB6.ADODC
		Me.Adodc1 = New VB6.ADODC
		Me.Text4 = New System.Windows.Forms.TextBox
		Me.Text3 = New System.Windows.Forms.TextBox
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.InactiveBorder
		Me.Text = "REPORT"
		Me.ClientSize = New System.Drawing.Size(662, 548)
		Me.Location = New System.Drawing.Point(8, 31)
		Me.Icon = CType(resources.GetObject("Form8.Icon"), System.Drawing.Icon)
		Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		Me.BackgroundImage = CType(resources.GetObject("Form8.BackgroundImage"), System.Drawing.Image)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Visible = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form8"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "PRINT"
		Me.Command1.Size = New System.Drawing.Size(121, 33)
		Me.Command1.Location = New System.Drawing.Point(464, 496)
		Me.Command1.TabIndex = 0
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Text5.AutoSize = False
		Me.Text5.Enabled = False
		Me.Text5.Size = New System.Drawing.Size(153, 33)
		Me.Text5.Location = New System.Drawing.Point(416, 8)
		Me.Text5.ReadOnly = True
		Me.Text5.TabIndex = 7
		Me.Text5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text5.AcceptsReturn = True
		Me.Text5.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text5.BackColor = System.Drawing.SystemColors.Window
		Me.Text5.CausesValidation = True
		Me.Text5.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text5.HideSelection = True
		Me.Text5.Maxlength = 0
		Me.Text5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text5.MultiLine = False
		Me.Text5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text5.TabStop = True
		Me.Text5.Visible = True
		Me.Text5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Text5.Name = "Text5"
		DataGrid2.OcxState = CType(resources.GetObject("DataGrid2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid2.Size = New System.Drawing.Size(265, 73)
		Me.DataGrid2.Location = New System.Drawing.Point(208, 152)
		Me.DataGrid2.TabIndex = 6
		Me.DataGrid2.Visible = False
		Me.DataGrid2.Name = "DataGrid2"
		DataGrid1.OcxState = CType(resources.GetObject("DataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid1.Size = New System.Drawing.Size(105, 49)
		Me.DataGrid1.Location = New System.Drawing.Point(72, 64)
		Me.DataGrid1.TabIndex = 5
		Me.DataGrid1.Visible = False
		Me.DataGrid1.Name = "DataGrid1"
		Me.Adodc2.Size = New System.Drawing.Size(105, 49)
		Me.Adodc2.Location = New System.Drawing.Point(136, 40)
		Me.Adodc2.Visible = 0
		Me.Adodc2.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.Adodc2.ConnectionTimeout = 15
		Me.Adodc2.CommandTimeout = 30
		Me.Adodc2.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		Me.Adodc2.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.Adodc2.CommandType = ADODB.CommandTypeEnum.adCmdTable
		Me.Adodc2.CacheSize = 50
		Me.Adodc2.MaxRecords = 0
		Me.Adodc2.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.Adodc2.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.Adodc2.BackColor = System.Drawing.SystemColors.Window
		Me.Adodc2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Adodc2.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.Adodc2.Enabled = True
		Me.Adodc2.UserName = ""
		Me.Adodc2.RecordSource = "issue"
		Me.Adodc2.Text = "Adodc2"
		Me.Adodc2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc2.Name = "Adodc2"
		Me.Adodc1.Size = New System.Drawing.Size(105, 49)
		Me.Adodc1.Location = New System.Drawing.Point(136, 56)
		Me.Adodc1.Visible = 0
		Me.Adodc1.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.Adodc1.ConnectionTimeout = 15
		Me.Adodc1.CommandTimeout = 30
		Me.Adodc1.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		Me.Adodc1.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.Adodc1.CommandType = ADODB.CommandTypeEnum.adCmdUnknown
		Me.Adodc1.CacheSize = 50
		Me.Adodc1.MaxRecords = 0
		Me.Adodc1.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.Adodc1.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.Adodc1.BackColor = System.Drawing.SystemColors.Window
		Me.Adodc1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Adodc1.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.Adodc1.Enabled = True
		Me.Adodc1.UserName = ""
		Me.Adodc1.RecordSource = "select * from booklist;"
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		Me.Text4.AutoSize = False
		Me.Text4.Enabled = False
		Me.Text4.Size = New System.Drawing.Size(273, 33)
		Me.Text4.Location = New System.Drawing.Point(384, 448)
		Me.Text4.ReadOnly = True
		Me.Text4.TabIndex = 4
		Me.Text4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text4.AcceptsReturn = True
		Me.Text4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text4.BackColor = System.Drawing.SystemColors.Window
		Me.Text4.CausesValidation = True
		Me.Text4.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text4.HideSelection = True
		Me.Text4.Maxlength = 0
		Me.Text4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text4.MultiLine = False
		Me.Text4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text4.TabStop = True
		Me.Text4.Visible = True
		Me.Text4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Text4.Name = "Text4"
		Me.Text3.AutoSize = False
		Me.Text3.Enabled = False
		Me.Text3.Size = New System.Drawing.Size(273, 33)
		Me.Text3.Location = New System.Drawing.Point(352, 352)
		Me.Text3.ReadOnly = True
		Me.Text3.TabIndex = 3
		Me.Text3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text3.AcceptsReturn = True
		Me.Text3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text3.BackColor = System.Drawing.SystemColors.Window
		Me.Text3.CausesValidation = True
		Me.Text3.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text3.HideSelection = True
		Me.Text3.Maxlength = 0
		Me.Text3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text3.MultiLine = False
		Me.Text3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text3.TabStop = True
		Me.Text3.Visible = True
		Me.Text3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Text3.Name = "Text3"
		Me.Text2.AutoSize = False
		Me.Text2.Enabled = False
		Me.Text2.Size = New System.Drawing.Size(281, 33)
		Me.Text2.Location = New System.Drawing.Point(352, 296)
		Me.Text2.ReadOnly = True
		Me.Text2.TabIndex = 2
		Me.Text2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text2.AcceptsReturn = True
		Me.Text2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text2.BackColor = System.Drawing.SystemColors.Window
		Me.Text2.CausesValidation = True
		Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text2.HideSelection = True
		Me.Text2.Maxlength = 0
		Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text2.MultiLine = False
		Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text2.TabStop = True
		Me.Text2.Visible = True
		Me.Text2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Text2.Name = "Text2"
		Me.Text1.AutoSize = False
		Me.Text1.Enabled = False
		Me.Text1.Size = New System.Drawing.Size(281, 33)
		Me.Text1.Location = New System.Drawing.Point(296, 216)
		Me.Text1.ReadOnly = True
		Me.Text1.TabIndex = 1
		Me.Text1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.BackColor = System.Drawing.SystemColors.Window
		Me.Text1.CausesValidation = True
		Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text1.HideSelection = True
		Me.Text1.Maxlength = 0
		Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text1.MultiLine = False
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.TabStop = True
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Text1.Name = "Text1"
		Me.Controls.Add(Command1)
		Me.Controls.Add(Text5)
		Me.Controls.Add(DataGrid2)
		Me.Controls.Add(DataGrid1)
		Me.Controls.Add(Adodc2)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(Text4)
		Me.Controls.Add(Text3)
		Me.Controls.Add(Text2)
		Me.Controls.Add(Text1)
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
#Region "Upgrade Support"
	Public Sub VB6_AddADODataBinding()
		DataGrid1.DataSource = CType(Adodc1, MSDATASRC.DataSource)
		DataGrid2.DataSource = CType(Adodc2, MSDATASRC.DataSource)
	End Sub
	Public Sub VB6_RemoveADODataBinding()
		DataGrid1.DataSource = Nothing
		DataGrid2.DataSource = Nothing
	End Sub
#End Region 
End Class