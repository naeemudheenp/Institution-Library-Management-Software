<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form11
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
	Private ADOBind_Adodc1 As VB6.MBindingCollection
	Private ADOBind_Adodc2 As VB6.MBindingCollection
	Public WithEvents Command9 As System.Windows.Forms.Button
	Public WithEvents Adodc2 As VB6.ADODC
	Public WithEvents DataGrid2 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Text4 As System.Windows.Forms.TextBox
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Text7 As System.Windows.Forms.TextBox
	Public WithEvents Text6 As System.Windows.Forms.TextBox
	Public WithEvents Text5 As System.Windows.Forms.TextBox
	Public WithEvents Text3 As System.Windows.Forms.TextBox
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents fuck As System.Windows.Forms.PictureBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form11))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Command9 = New System.Windows.Forms.Button
		Me.Adodc2 = New VB6.ADODC
		Me.DataGrid2 = New AxMSDataGridLib.AxDataGrid
		Me.Command3 = New System.Windows.Forms.Button
		Me.Text4 = New System.Windows.Forms.TextBox
		Me.Command4 = New System.Windows.Forms.Button
		Me.Command1 = New System.Windows.Forms.Button
		Me.Adodc1 = New VB6.ADODC
		Me.DataGrid1 = New AxMSDataGridLib.AxDataGrid
		Me.Command2 = New System.Windows.Forms.Button
		Me.Text7 = New System.Windows.Forms.TextBox
		Me.Text6 = New System.Windows.Forms.TextBox
		Me.Text5 = New System.Windows.Forms.TextBox
		Me.Text3 = New System.Windows.Forms.TextBox
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.fuck = New System.Windows.Forms.PictureBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "NOT ISSUED BOOKS"
		Me.ClientSize = New System.Drawing.Size(1083, 557)
		Me.Location = New System.Drawing.Point(8, 31)
		Me.Icon = CType(resources.GetObject("Form11.Icon"), System.Drawing.Icon)
		Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		Me.BackgroundImage = CType(resources.GetObject("Form11.BackgroundImage"), System.Drawing.Image)
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
		Me.Name = "Form11"
		Me.Command9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command9.Text = "CLOSE GRID"
		Me.Command9.Size = New System.Drawing.Size(65, 33)
		Me.Command9.Location = New System.Drawing.Point(1000, 0)
		Me.Command9.TabIndex = 15
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
		Me.Adodc2.Size = New System.Drawing.Size(97, 49)
		Me.Adodc2.Location = New System.Drawing.Point(480, 256)
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
		Me.Adodc2.RecordSource = "SELECT* FROM BOOKLIST;"
		Me.Adodc2.Text = "Adodc2"
		Me.Adodc2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc2.Name = "Adodc2"
		DataGrid2.OcxState = CType(resources.GetObject("DataGrid2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid2.Size = New System.Drawing.Size(393, 57)
		Me.DataGrid2.Location = New System.Drawing.Point(664, 32)
		Me.DataGrid2.TabIndex = 14
		Me.DataGrid2.Visible = False
		Me.DataGrid2.Name = "DataGrid2"
		Me.Command3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command3.Text = "FIND"
		Me.Command3.Size = New System.Drawing.Size(57, 33)
		Me.Command3.Location = New System.Drawing.Point(944, 0)
		Me.Command3.TabIndex = 1
		Me.Command3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command3.BackColor = System.Drawing.SystemColors.Control
		Me.Command3.CausesValidation = True
		Me.Command3.Enabled = True
		Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command3.TabStop = True
		Me.Command3.Name = "Command3"
		Me.Text4.AutoSize = False
		Me.Text4.Size = New System.Drawing.Size(281, 33)
		Me.Text4.Location = New System.Drawing.Point(664, 0)
		Me.Text4.TabIndex = 0
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
		Me.Text4.Visible = True
		Me.Text4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text4.Name = "Text4"
		Me.Command4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command4.Text = "PRINT"
		Me.Command4.Size = New System.Drawing.Size(145, 41)
		Me.Command4.Location = New System.Drawing.Point(152, -8)
		Me.Command4.TabIndex = 3
		Me.Command4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command4.BackColor = System.Drawing.SystemColors.Control
		Me.Command4.CausesValidation = True
		Me.Command4.Enabled = True
		Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command4.TabStop = True
		Me.Command4.Name = "Command4"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "REFRESH"
		Me.Command1.Size = New System.Drawing.Size(153, 33)
		Me.Command1.Location = New System.Drawing.Point(296, 0)
		Me.Command1.TabIndex = 2
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Adodc1.Size = New System.Drawing.Size(97, 49)
		Me.Adodc1.Location = New System.Drawing.Point(624, 520)
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
		Me.Adodc1.RecordSource = "SELECT booklist.bookid, booklist.bookname, booklist.author, booklist.publication, booklist.booksleft, booklist.qty, booklist.shelf, booklist.rack, booklist.price" & Chr(13) & Chr(10) & "FROM booklist LEFT JOIN issue ON booklist.bookid = issue.bookid" & Chr(13) & Chr(10) & "WHERE (((issue.bookid) Is Null));" & Chr(13) & Chr(10)
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		DataGrid1.OcxState = CType(resources.GetObject("DataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid1.Size = New System.Drawing.Size(1089, 529)
		Me.DataGrid1.Location = New System.Drawing.Point(0, 32)
		Me.DataGrid1.TabIndex = 13
		Me.DataGrid1.Name = "DataGrid1"
		Me.Command2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command2.Text = "MAIN MENU"
		Me.Command2.Size = New System.Drawing.Size(153, 33)
		Me.Command2.Location = New System.Drawing.Point(0, 0)
		Me.Command2.TabIndex = 4
		Me.Command2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command2.BackColor = System.Drawing.SystemColors.Control
		Me.Command2.CausesValidation = True
		Me.Command2.Enabled = True
		Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command2.TabStop = True
		Me.Command2.Name = "Command2"
		Me.Text7.AutoSize = False
		Me.Text7.Size = New System.Drawing.Size(81, 33)
		Me.Text7.Location = New System.Drawing.Point(64, 528)
		Me.Text7.TabIndex = 11
		Me.Text7.Text = "Text7"
		Me.Text7.Visible = False
		Me.Text7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text7.AcceptsReturn = True
		Me.Text7.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text7.BackColor = System.Drawing.SystemColors.Window
		Me.Text7.CausesValidation = True
		Me.Text7.Enabled = True
		Me.Text7.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text7.HideSelection = True
		Me.Text7.ReadOnly = False
		Me.Text7.Maxlength = 0
		Me.Text7.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text7.MultiLine = False
		Me.Text7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text7.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text7.TabStop = True
		Me.Text7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text7.Name = "Text7"
		Me.Text6.AutoSize = False
		Me.Text6.Size = New System.Drawing.Size(81, 33)
		Me.Text6.Location = New System.Drawing.Point(152, 528)
		Me.Text6.TabIndex = 10
		Me.Text6.Text = "Text6"
		Me.Text6.Visible = False
		Me.Text6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text6.AcceptsReturn = True
		Me.Text6.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text6.BackColor = System.Drawing.SystemColors.Window
		Me.Text6.CausesValidation = True
		Me.Text6.Enabled = True
		Me.Text6.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text6.HideSelection = True
		Me.Text6.ReadOnly = False
		Me.Text6.Maxlength = 0
		Me.Text6.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text6.MultiLine = False
		Me.Text6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text6.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text6.TabStop = True
		Me.Text6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text6.Name = "Text6"
		Me.Text5.AutoSize = False
		Me.Text5.Size = New System.Drawing.Size(81, 33)
		Me.Text5.Location = New System.Drawing.Point(272, 560)
		Me.Text5.TabIndex = 9
		Me.Text5.Text = "Text5"
		Me.Text5.Visible = False
		Me.Text5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text5.AcceptsReturn = True
		Me.Text5.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text5.BackColor = System.Drawing.SystemColors.Window
		Me.Text5.CausesValidation = True
		Me.Text5.Enabled = True
		Me.Text5.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text5.HideSelection = True
		Me.Text5.ReadOnly = False
		Me.Text5.Maxlength = 0
		Me.Text5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text5.MultiLine = False
		Me.Text5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text5.TabStop = True
		Me.Text5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text5.Name = "Text5"
		Me.Text3.AutoSize = False
		Me.Text3.Size = New System.Drawing.Size(81, 33)
		Me.Text3.Location = New System.Drawing.Point(64, 560)
		Me.Text3.TabIndex = 8
		Me.Text3.Text = "Text3"
		Me.Text3.Visible = False
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
		Me.Text3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text3.Name = "Text3"
		Me.Text2.AutoSize = False
		Me.Text2.Size = New System.Drawing.Size(81, 33)
		Me.Text2.Location = New System.Drawing.Point(352, 568)
		Me.Text2.TabIndex = 7
		Me.Text2.Text = "Text2"
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
		Me.Text1.Location = New System.Drawing.Point(152, 560)
		Me.Text1.TabIndex = 6
		Me.Text1.Text = "Text1"
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
		Me.fuck.BackColor = System.Drawing.SystemColors.Window
		Me.fuck.ForeColor = System.Drawing.SystemColors.WindowText
		Me.fuck.Size = New System.Drawing.Size(105, 49)
		Me.fuck.Location = New System.Drawing.Point(488, 752)
		Me.fuck.TabIndex = 5
		Me.fuck.Visible = False
		Me.fuck.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fuck.Dock = System.Windows.Forms.DockStyle.None
		Me.fuck.CausesValidation = True
		Me.fuck.Enabled = True
		Me.fuck.Cursor = System.Windows.Forms.Cursors.Default
		Me.fuck.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fuck.TabStop = True
		Me.fuck.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.fuck.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.fuck.Name = "fuck"
		Me.Label1.Text = "                          NOT ISSUED BOOKS"
		Me.Label1.Size = New System.Drawing.Size(217, 33)
		Me.Label1.Location = New System.Drawing.Point(448, 0)
		Me.Label1.TabIndex = 12
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(Command9)
		Me.Controls.Add(Adodc2)
		Me.Controls.Add(DataGrid2)
		Me.Controls.Add(Command3)
		Me.Controls.Add(Text4)
		Me.Controls.Add(Command4)
		Me.Controls.Add(Command1)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(DataGrid1)
		Me.Controls.Add(Command2)
		Me.Controls.Add(Text7)
		Me.Controls.Add(Text6)
		Me.Controls.Add(Text5)
		Me.Controls.Add(Text3)
		Me.Controls.Add(Text2)
		Me.Controls.Add(Text1)
		Me.Controls.Add(fuck)
		Me.Controls.Add(Label1)
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
#Region "Upgrade Support"
	Public Sub VB6_AddADODataBinding()
		ADOBind_Adodc1 = New VB6.MBindingCollection()
		ADOBind_Adodc2 = New VB6.MBindingCollection()
		ADOBind_Adodc1.DataSource = CType(Adodc1, MSDATASRC.DataSource)
		ADOBind_Adodc1.Add(Text6, "Text", "bookid", Nothing, "Text6")
		ADOBind_Adodc2.DataSource = CType(Adodc2, MSDATASRC.DataSource)
		ADOBind_Adodc2.Add(Text7, "Text", "bookid", Nothing, "Text7")
		DataGrid1.DataSource = CType(Adodc1, MSDATASRC.DataSource)
		DataGrid2.DataSource = CType(Adodc2, MSDATASRC.DataSource)
		ADOBind_Adodc1.UpdateMode = VB6.UpdateMode.vbUpdateWhenPropertyChanges
		ADOBind_Adodc1.UpdateControls()
		ADOBind_Adodc2.UpdateMode = VB6.UpdateMode.vbUpdateWhenPropertyChanges
		ADOBind_Adodc2.UpdateControls()
	End Sub
	Public Sub VB6_RemoveADODataBinding()
		ADOBind_Adodc1.Clear()
		ADOBind_Adodc1.Dispose()
		ADOBind_Adodc1 = Nothing
		ADOBind_Adodc2.Clear()
		ADOBind_Adodc2.Dispose()
		ADOBind_Adodc2 = Nothing
		DataGrid1.DataSource = Nothing
		DataGrid2.DataSource = Nothing
	End Sub
#End Region 
End Class