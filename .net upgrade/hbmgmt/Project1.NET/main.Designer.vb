<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form1
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
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Text11 As System.Windows.Forms.TextBox
	Public WithEvents Text10 As System.Windows.Forms.TextBox
	Public WithEvents Text9 As System.Windows.Forms.TextBox
	Public WithEvents Text8 As System.Windows.Forms.TextBox
	Public WithEvents Adodc3 As VB6.ADODC
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Text7 As System.Windows.Forms.TextBox
	Public WithEvents Text6 As System.Windows.Forms.TextBox
	Public WithEvents Text5 As System.Windows.Forms.TextBox
	Public WithEvents Text4 As System.Windows.Forms.TextBox
	Public WithEvents Text3 As System.Windows.Forms.TextBox
	Public WithEvents Adodc2 As VB6.ADODC
	Public WithEvents DataGrid2 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Command4 = New System.Windows.Forms.Button
		Me.Command3 = New System.Windows.Forms.Button
		Me.Text11 = New System.Windows.Forms.TextBox
		Me.Text10 = New System.Windows.Forms.TextBox
		Me.Text9 = New System.Windows.Forms.TextBox
		Me.Text8 = New System.Windows.Forms.TextBox
		Me.Adodc3 = New VB6.ADODC
		Me.Command2 = New System.Windows.Forms.Button
		Me.Text7 = New System.Windows.Forms.TextBox
		Me.Text6 = New System.Windows.Forms.TextBox
		Me.Text5 = New System.Windows.Forms.TextBox
		Me.Text4 = New System.Windows.Forms.TextBox
		Me.Text3 = New System.Windows.Forms.TextBox
		Me.Adodc2 = New VB6.ADODC
		Me.DataGrid2 = New AxMSDataGridLib.AxDataGrid
		Me.Adodc1 = New VB6.ADODC
		Me.DataGrid1 = New AxMSDataGridLib.AxDataGrid
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.Command1 = New System.Windows.Forms.Button
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(780, 584)
		Me.Location = New System.Drawing.Point(3, 26)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form1"
		Me.Command4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command4.Text = "MAIN MENU"
		Me.Command4.Size = New System.Drawing.Size(169, 73)
		Me.Command4.Location = New System.Drawing.Point(600, 120)
		Me.Command4.TabIndex = 20
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
		Me.Command3.Text = "HIDE"
		Me.Command3.Size = New System.Drawing.Size(97, 17)
		Me.Command3.Location = New System.Drawing.Point(680, 72)
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
		Me.Text11.AutoSize = False
		Me.Text11.Size = New System.Drawing.Size(169, 49)
		Me.Text11.Location = New System.Drawing.Point(600, 264)
		Me.Text11.TabIndex = 19
		Me.Text11.Text = "0"
		Me.Text11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text11.AcceptsReturn = True
		Me.Text11.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text11.BackColor = System.Drawing.SystemColors.Window
		Me.Text11.CausesValidation = True
		Me.Text11.Enabled = True
		Me.Text11.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text11.HideSelection = True
		Me.Text11.ReadOnly = False
		Me.Text11.Maxlength = 0
		Me.Text11.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text11.MultiLine = False
		Me.Text11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text11.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text11.TabStop = True
		Me.Text11.Visible = True
		Me.Text11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text11.Name = "Text11"
		Me.Text10.AutoSize = False
		Me.Text10.Size = New System.Drawing.Size(169, 49)
		Me.Text10.Location = New System.Drawing.Point(600, 376)
		Me.Text10.TabIndex = 16
		Me.Text10.Text = "0"
		Me.Text10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text10.AcceptsReturn = True
		Me.Text10.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text10.BackColor = System.Drawing.SystemColors.Window
		Me.Text10.CausesValidation = True
		Me.Text10.Enabled = True
		Me.Text10.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text10.HideSelection = True
		Me.Text10.ReadOnly = False
		Me.Text10.Maxlength = 0
		Me.Text10.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text10.MultiLine = False
		Me.Text10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text10.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text10.TabStop = True
		Me.Text10.Visible = True
		Me.Text10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text10.Name = "Text10"
		Me.Text9.AutoSize = False
		Me.Text9.Size = New System.Drawing.Size(153, 33)
		Me.Text9.Location = New System.Drawing.Point(128, 592)
		Me.Text9.TabIndex = 13
		Me.Text9.Text = "Text9"
		Me.Text9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text9.AcceptsReturn = True
		Me.Text9.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text9.BackColor = System.Drawing.SystemColors.Window
		Me.Text9.CausesValidation = True
		Me.Text9.Enabled = True
		Me.Text9.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text9.HideSelection = True
		Me.Text9.ReadOnly = False
		Me.Text9.Maxlength = 0
		Me.Text9.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text9.MultiLine = False
		Me.Text9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text9.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text9.TabStop = True
		Me.Text9.Visible = True
		Me.Text9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text9.Name = "Text9"
		Me.Text8.AutoSize = False
		Me.Text8.Size = New System.Drawing.Size(97, 19)
		Me.Text8.Location = New System.Drawing.Point(240, 608)
		Me.Text8.TabIndex = 12
		Me.Text8.Text = "Text8"
		Me.Text8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text8.AcceptsReturn = True
		Me.Text8.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text8.BackColor = System.Drawing.SystemColors.Window
		Me.Text8.CausesValidation = True
		Me.Text8.Enabled = True
		Me.Text8.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text8.HideSelection = True
		Me.Text8.ReadOnly = False
		Me.Text8.Maxlength = 0
		Me.Text8.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text8.MultiLine = False
		Me.Text8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text8.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text8.TabStop = True
		Me.Text8.Visible = True
		Me.Text8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text8.Name = "Text8"
		Me.Adodc3.Size = New System.Drawing.Size(161, 22)
		Me.Adodc3.Location = New System.Drawing.Point(192, 672)
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
		Me.Adodc3.RecordSource = "select * from sales;"
		Me.Adodc3.Text = "Adodc3"
		Me.Adodc3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
		Me.Adodc3.Name = "Adodc3"
		Me.Command2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command2.Text = "PURCHASE"
		Me.Command2.Size = New System.Drawing.Size(169, 41)
		Me.Command2.Location = New System.Drawing.Point(600, 536)
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
		Me.Text7.Size = New System.Drawing.Size(169, 49)
		Me.Text7.Location = New System.Drawing.Point(600, 488)
		Me.Text7.TabIndex = 11
		Me.Text7.Text = "0"
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
		Me.Text7.Visible = True
		Me.Text7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text7.Name = "Text7"
		Me.Text6.AutoSize = False
		Me.Text6.Size = New System.Drawing.Size(113, 25)
		Me.Text6.Location = New System.Drawing.Point(224, 616)
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
		Me.Text5.Size = New System.Drawing.Size(121, 25)
		Me.Text5.Location = New System.Drawing.Point(408, 592)
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
		Me.Text4.AutoSize = False
		Me.Text4.Size = New System.Drawing.Size(121, 19)
		Me.Text4.Location = New System.Drawing.Point(304, 608)
		Me.Text4.TabIndex = 8
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
		Me.Text3.AutoSize = False
		Me.Text3.Size = New System.Drawing.Size(121, 19)
		Me.Text3.Location = New System.Drawing.Point(176, 600)
		Me.Text3.TabIndex = 7
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
		Me.Adodc2.Size = New System.Drawing.Size(177, 25)
		Me.Adodc2.Location = New System.Drawing.Point(24, 608)
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
		Me.Adodc2.RecordSource = "select * from products;"
		Me.Adodc2.Text = "Adodc2"
		Me.Adodc2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
		Me.Adodc2.Name = "Adodc2"
		DataGrid2.OcxState = CType(resources.GetObject("DataGrid2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid2.Size = New System.Drawing.Size(569, 57)
		Me.DataGrid2.Location = New System.Drawing.Point(16, 88)
		Me.DataGrid2.TabIndex = 6
		Me.DataGrid2.Visible = False
		Me.DataGrid2.Name = "DataGrid2"
		Me.Adodc1.Size = New System.Drawing.Size(177, 33)
		Me.Adodc1.Location = New System.Drawing.Point(160, 600)
		Me.Adodc1.Visible = 0
		Me.Adodc1.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.Adodc1.ConnectionTimeout = 15
		Me.Adodc1.CommandTimeout = 30
		Me.Adodc1.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		Me.Adodc1.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.Adodc1.CommandType = ADODB.CommandTypeEnum.adCmdTable
		Me.Adodc1.CacheSize = 50
		Me.Adodc1.MaxRecords = 0
		Me.Adodc1.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.Adodc1.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.Adodc1.BackColor = System.Drawing.SystemColors.Window
		Me.Adodc1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Adodc1.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.Adodc1.Enabled = True
		Me.Adodc1.UserName = ""
		Me.Adodc1.RecordSource = "ontime"
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		DataGrid1.OcxState = CType(resources.GetObject("DataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid1.Size = New System.Drawing.Size(569, 425)
		Me.DataGrid1.Location = New System.Drawing.Point(16, 152)
		Me.DataGrid1.TabIndex = 5
		Me.DataGrid1.Name = "DataGrid1"
		Me.Text2.AutoSize = False
		Me.Text2.Size = New System.Drawing.Size(89, 41)
		Me.Text2.Location = New System.Drawing.Point(592, 48)
		Me.Text2.TabIndex = 1
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
		Me.Text2.Visible = True
		Me.Text2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text2.Name = "Text2"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "ENTER"
		Me.Command1.Size = New System.Drawing.Size(97, 25)
		Me.Command1.Location = New System.Drawing.Point(680, 48)
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
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(569, 41)
		Me.Text1.Location = New System.Drawing.Point(16, 48)
		Me.Text1.TabIndex = 0
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
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		Me.Label4.Text = "TOTAL QUANTITY"
		Me.Label4.Size = New System.Drawing.Size(169, 49)
		Me.Label4.Location = New System.Drawing.Point(600, 216)
		Me.Label4.TabIndex = 18
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.Text = "TOTAL AMOUNT"
		Me.Label3.Size = New System.Drawing.Size(169, 49)
		Me.Label3.Location = New System.Drawing.Point(600, 328)
		Me.Label3.TabIndex = 17
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.Text = "FINAL AMOUNT"
		Me.Label2.Size = New System.Drawing.Size(169, 49)
		Me.Label2.Location = New System.Drawing.Point(600, 440)
		Me.Label2.TabIndex = 15
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
		Me.Label1.Text = "                                                                                                                     SALES WINDOW"
		Me.Label1.Size = New System.Drawing.Size(785, 33)
		Me.Label1.Location = New System.Drawing.Point(0, 0)
		Me.Label1.TabIndex = 14
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
		Me.Controls.Add(Command4)
		Me.Controls.Add(Command3)
		Me.Controls.Add(Text11)
		Me.Controls.Add(Text10)
		Me.Controls.Add(Text9)
		Me.Controls.Add(Text8)
		Me.Controls.Add(Adodc3)
		Me.Controls.Add(Command2)
		Me.Controls.Add(Text7)
		Me.Controls.Add(Text6)
		Me.Controls.Add(Text5)
		Me.Controls.Add(Text4)
		Me.Controls.Add(Text3)
		Me.Controls.Add(Adodc2)
		Me.Controls.Add(DataGrid2)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(DataGrid1)
		Me.Controls.Add(Text2)
		Me.Controls.Add(Command1)
		Me.Controls.Add(Text1)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
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