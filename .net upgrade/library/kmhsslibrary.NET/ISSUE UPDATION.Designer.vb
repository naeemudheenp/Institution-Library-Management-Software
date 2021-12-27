<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form5
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
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents Command9 As System.Windows.Forms.Button
	Public WithEvents DataGrid2 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Adodc3 As VB6.ADODC
	Public WithEvents Text4 As System.Windows.Forms.TextBox
	Public WithEvents Command5 As System.Windows.Forms.Button
	Public WithEvents Text3 As System.Windows.Forms.TextBox
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Adodc2 As VB6.ADODC
	Public WithEvents adodc1 As VB6.ADODC
	Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form5))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.Picture1 = New System.Windows.Forms.PictureBox
		Me.Command9 = New System.Windows.Forms.Button
		Me.DataGrid2 = New AxMSDataGridLib.AxDataGrid
		Me.Adodc3 = New VB6.ADODC
		Me.Text4 = New System.Windows.Forms.TextBox
		Me.Command5 = New System.Windows.Forms.Button
		Me.Text3 = New System.Windows.Forms.TextBox
		Me.Command4 = New System.Windows.Forms.Button
		Me.Command3 = New System.Windows.Forms.Button
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.Adodc2 = New VB6.ADODC
		Me.adodc1 = New VB6.ADODC
		Me.DataGrid1 = New AxMSDataGridLib.AxDataGrid
		Me.Command2 = New System.Windows.Forms.Button
		Me.Command1 = New System.Windows.Forms.Button
		Me.Label2 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "ISSUE UPDATION"
		Me.ClientSize = New System.Drawing.Size(724, 693)
		Me.Location = New System.Drawing.Point(8, 31)
		Me.Icon = CType(resources.GetObject("Form5.Icon"), System.Drawing.Icon)
		Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		Me.BackgroundImage = CType(resources.GetObject("Form5.BackgroundImage"), System.Drawing.Image)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
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
		Me.Name = "Form5"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.Picture1.BackColor = System.Drawing.SystemColors.Window
		Me.Picture1.Size = New System.Drawing.Size(265, 273)
		Me.Picture1.Location = New System.Drawing.Point(416, 264)
		Me.Picture1.TabIndex = 13
		Me.Picture1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture1.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture1.CausesValidation = True
		Me.Picture1.Enabled = True
		Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture1.TabStop = True
		Me.Picture1.Visible = True
		Me.Picture1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Picture1.Name = "Picture1"
		Me.Command9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command9.Text = "CLOSE GRID"
		Me.Command9.Size = New System.Drawing.Size(73, 33)
		Me.Command9.Location = New System.Drawing.Point(648, 32)
		Me.Command9.TabIndex = 12
		Me.Command9.Visible = False
		Me.Command9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command9.BackColor = System.Drawing.SystemColors.Control
		Me.Command9.CausesValidation = True
		Me.Command9.Enabled = True
		Me.Command9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command9.TabStop = True
		Me.Command9.Name = "Command9"
		DataGrid2.OcxState = CType(resources.GetObject("DataGrid2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid2.Size = New System.Drawing.Size(345, 129)
		Me.DataGrid2.Location = New System.Drawing.Point(376, 64)
		Me.DataGrid2.TabIndex = 11
		Me.DataGrid2.Visible = False
		Me.DataGrid2.Name = "DataGrid2"
		Me.Adodc3.Size = New System.Drawing.Size(97, 57)
		Me.Adodc3.Location = New System.Drawing.Point(40, 256)
		Me.Adodc3.Visible = 0
		Me.Adodc3.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.Adodc3.ConnectionTimeout = 15
		Me.Adodc3.CommandTimeout = 30
		Me.Adodc3.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		Me.Adodc3.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.Adodc3.CommandType = ADODB.CommandTypeEnum.adCmdUnknown
		Me.Adodc3.CacheSize = 50
		Me.Adodc3.MaxRecords = 0
		Me.Adodc3.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.Adodc3.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.Adodc3.BackColor = System.Drawing.SystemColors.Window
		Me.Adodc3.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Adodc3.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.Adodc3.Enabled = True
		Me.Adodc3.UserName = ""
		Me.Adodc3.RecordSource = "SELECT* FROM issue;"
		Me.Adodc3.Text = "Adodc3"
		Me.Adodc3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc3.Name = "Adodc3"
		Me.Text4.AutoSize = False
		Me.Text4.Size = New System.Drawing.Size(81, 33)
		Me.Text4.Location = New System.Drawing.Point(320, 264)
		Me.Text4.TabIndex = 10
		Me.Text4.Text = "Text4"
		Me.Text4.Visible = False
		Me.Text4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text4.AcceptsReturn = True
		Me.Text4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text4.BackColor = System.Drawing.SystemColors.Window
		Me.Text4.CausesValidation = True
		Me.Text4.Enabled = True
		Me.Text4.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text4.HideSelection = True
		Me.Text4.ReadOnly = False
		Me.Text4.Maxlength = 0
		Me.Text4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text4.MultiLine = False
		Me.Text4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text4.TabStop = True
		Me.Text4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text4.Name = "Text4"
		Me.Command5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command5.Text = "FIND"
		Me.Command5.Size = New System.Drawing.Size(81, 33)
		Me.Command5.Location = New System.Drawing.Point(568, 32)
		Me.Command5.TabIndex = 1
		Me.Command5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command5.BackColor = System.Drawing.SystemColors.Control
		Me.Command5.CausesValidation = True
		Me.Command5.Enabled = True
		Me.Command5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command5.TabStop = True
		Me.Command5.Name = "Command5"
		Me.Text3.AutoSize = False
		Me.Text3.Size = New System.Drawing.Size(193, 33)
		Me.Text3.Location = New System.Drawing.Point(376, 32)
		Me.Text3.TabIndex = 0
		Me.Text3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text3.AcceptsReturn = True
		Me.Text3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text3.BackColor = System.Drawing.SystemColors.Window
		Me.Text3.CausesValidation = True
		Me.Text3.Enabled = True
		Me.Text3.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text3.HideSelection = True
		Me.Text3.ReadOnly = False
		Me.Text3.Maxlength = 0
		Me.Text3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text3.MultiLine = False
		Me.Text3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text3.TabStop = True
		Me.Text3.Visible = True
		Me.Text3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text3.Name = "Text3"
		Me.Command4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command4.Text = "REFRESH"
		Me.Command4.Size = New System.Drawing.Size(105, 65)
		Me.Command4.Location = New System.Drawing.Point(272, 632)
		Me.Command4.TabIndex = 2
		Me.Command4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command4.BackColor = System.Drawing.SystemColors.Control
		Me.Command4.CausesValidation = True
		Me.Command4.Enabled = True
		Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command4.TabStop = True
		Me.Command4.Name = "Command4"
		Me.Command3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command3.Text = "PRINT"
		Me.Command3.Size = New System.Drawing.Size(89, 65)
		Me.Command3.Location = New System.Drawing.Point(184, 632)
		Me.Command3.TabIndex = 3
		Me.Command3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command3.BackColor = System.Drawing.SystemColors.Control
		Me.Command3.CausesValidation = True
		Me.Command3.Enabled = True
		Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command3.TabStop = True
		Me.Command3.Name = "Command3"
		Me.Text2.AutoSize = False
		Me.Text2.Size = New System.Drawing.Size(81, 33)
		Me.Text2.Location = New System.Drawing.Point(88, 368)
		Me.Text2.TabIndex = 9
		Me.Text2.Visible = False
		Me.Text2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text2.AcceptsReturn = True
		Me.Text2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text2.BackColor = System.Drawing.SystemColors.Window
		Me.Text2.CausesValidation = True
		Me.Text2.Enabled = True
		Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text2.HideSelection = True
		Me.Text2.ReadOnly = False
		Me.Text2.Maxlength = 0
		Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text2.MultiLine = False
		Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text2.TabStop = True
		Me.Text2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text2.Name = "Text2"
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(81, 33)
		Me.Text1.Location = New System.Drawing.Point(184, 264)
		Me.Text1.TabIndex = 8
		Me.Text1.Visible = False
		Me.Text1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.BackColor = System.Drawing.SystemColors.Window
		Me.Text1.CausesValidation = True
		Me.Text1.Enabled = True
		Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text1.HideSelection = True
		Me.Text1.ReadOnly = False
		Me.Text1.Maxlength = 0
		Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text1.MultiLine = False
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.TabStop = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		Me.Adodc2.Size = New System.Drawing.Size(105, 57)
		Me.Adodc2.Location = New System.Drawing.Point(104, 232)
		Me.Adodc2.Visible = 0
		Me.Adodc2.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.Adodc2.ConnectionTimeout = 15
		Me.Adodc2.CommandTimeout = 30
		Me.Adodc2.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		Me.Adodc2.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.Adodc2.CommandType = ADODB.CommandTypeEnum.adCmdUnknown
		Me.Adodc2.CacheSize = 50
		Me.Adodc2.MaxRecords = 0
		Me.Adodc2.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.Adodc2.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.Adodc2.BackColor = System.Drawing.SystemColors.Window
		Me.Adodc2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Adodc2.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.Adodc2.Enabled = True
		Me.Adodc2.UserName = ""
		Me.Adodc2.RecordSource = "select * from booklist;"
		Me.Adodc2.Text = "Adodc2"
		Me.Adodc2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc2.Name = "Adodc2"
		Me.adodc1.Size = New System.Drawing.Size(105, 57)
		Me.adodc1.Location = New System.Drawing.Point(152, 280)
		Me.adodc1.Visible = 0
		Me.adodc1.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.adodc1.ConnectionTimeout = 15
		Me.adodc1.CommandTimeout = 30
		Me.adodc1.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		Me.adodc1.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.adodc1.CommandType = ADODB.CommandTypeEnum.adCmdTable
		Me.adodc1.CacheSize = 50
		Me.adodc1.MaxRecords = 0
		Me.adodc1.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.adodc1.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.adodc1.BackColor = System.Drawing.SystemColors.Window
		Me.adodc1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.adodc1.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.adodc1.Enabled = True
		Me.adodc1.UserName = ""
		Me.adodc1.RecordSource = "issue"
		Me.adodc1.Text = "Adodc1"
		Me.adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.adodc1.Name = "adodc1"
		DataGrid1.OcxState = CType(resources.GetObject("DataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid1.Size = New System.Drawing.Size(377, 601)
		Me.DataGrid1.Location = New System.Drawing.Point(0, 32)
		Me.DataGrid1.TabIndex = 7
		Me.DataGrid1.Name = "DataGrid1"
		Me.Command2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command2.Text = "MAIN MENU"
		Me.Command2.Size = New System.Drawing.Size(97, 65)
		Me.Command2.Location = New System.Drawing.Point(0, 632)
		Me.Command2.TabIndex = 5
		Me.Command2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command2.BackColor = System.Drawing.SystemColors.Control
		Me.Command2.CausesValidation = True
		Me.Command2.Enabled = True
		Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command2.TabStop = True
		Me.Command2.Name = "Command2"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "REMOVE ISSUE"
		Me.Command1.Size = New System.Drawing.Size(89, 65)
		Me.Command1.Location = New System.Drawing.Point(96, 632)
		Me.Command1.TabIndex = 4
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Label2.Text = "                                        ISSUE LIST"
		Me.Label2.Size = New System.Drawing.Size(961, 33)
		Me.Label2.Location = New System.Drawing.Point(0, 0)
		Me.Label2.TabIndex = 6
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Controls.Add(Picture1)
		Me.Controls.Add(Command9)
		Me.Controls.Add(DataGrid2)
		Me.Controls.Add(Adodc3)
		Me.Controls.Add(Text4)
		Me.Controls.Add(Command5)
		Me.Controls.Add(Text3)
		Me.Controls.Add(Command4)
		Me.Controls.Add(Command3)
		Me.Controls.Add(Text2)
		Me.Controls.Add(Text1)
		Me.Controls.Add(Adodc2)
		Me.Controls.Add(adodc1)
		Me.Controls.Add(DataGrid1)
		Me.Controls.Add(Command2)
		Me.Controls.Add(Command1)
		Me.Controls.Add(Label2)
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
#Region "Upgrade Support"
	Public Sub VB6_AddADODataBinding()
		DataGrid2.DataSource = CType(Adodc3, MSDATASRC.DataSource)
	End Sub
	Public Sub VB6_RemoveADODataBinding()
		DataGrid2.DataSource = Nothing
	End Sub
#End Region 
End Class