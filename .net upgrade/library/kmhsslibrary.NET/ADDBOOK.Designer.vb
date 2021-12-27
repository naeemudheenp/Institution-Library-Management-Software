<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form2
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
	Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents Text8 As System.Windows.Forms.TextBox
	Public WithEvents Text7 As System.Windows.Forms.TextBox
	Public WithEvents Text6 As System.Windows.Forms.TextBox
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Text5 As System.Windows.Forms.TextBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Text4 As System.Windows.Forms.TextBox
	Public WithEvents Text3 As System.Windows.Forms.TextBox
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form2))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.DataGrid1 = New AxMSDataGridLib.AxDataGrid
		Me.Adodc1 = New VB6.ADODC
		Me.Text8 = New System.Windows.Forms.TextBox
		Me.Text7 = New System.Windows.Forms.TextBox
		Me.Text6 = New System.Windows.Forms.TextBox
		Me.Command2 = New System.Windows.Forms.Button
		Me.Text5 = New System.Windows.Forms.TextBox
		Me.Command1 = New System.Windows.Forms.Button
		Me.Text4 = New System.Windows.Forms.TextBox
		Me.Text3 = New System.Windows.Forms.TextBox
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "ADD BOOK"
		Me.ClientSize = New System.Drawing.Size(936, 351)
		Me.Location = New System.Drawing.Point(8, 31)
		Me.Icon = CType(resources.GetObject("Form2.Icon"), System.Drawing.Icon)
		Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		Me.BackgroundImage = CType(resources.GetObject("Form2.BackgroundImage"), System.Drawing.Image)
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
		Me.Name = "Form2"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		DataGrid1.OcxState = CType(resources.GetObject("DataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid1.Size = New System.Drawing.Size(105, 49)
		Me.DataGrid1.Location = New System.Drawing.Point(304, 8)
		Me.DataGrid1.TabIndex = 19
		Me.DataGrid1.Visible = False
		Me.DataGrid1.Name = "DataGrid1"
		Me.Adodc1.Size = New System.Drawing.Size(105, 57)
		Me.Adodc1.Location = New System.Drawing.Point(800, 416)
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
		Me.Adodc1.RecordSource = "booklist"
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		Me.Text8.AutoSize = False
		Me.Text8.Size = New System.Drawing.Size(161, 25)
		Me.Text8.Location = New System.Drawing.Point(168, 232)
		Me.Text8.TabIndex = 5
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
		Me.Text7.AutoSize = False
		Me.Text7.Size = New System.Drawing.Size(161, 27)
		Me.Text7.Location = New System.Drawing.Point(168, 168)
		Me.Text7.TabIndex = 3
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
		Me.Text6.Size = New System.Drawing.Size(161, 27)
		Me.Text6.Location = New System.Drawing.Point(168, 200)
		Me.Text6.TabIndex = 4
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
		Me.Text6.Visible = True
		Me.Text6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text6.Name = "Text6"
		Me.Command2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command2.Text = "MAIN MENU"
		Me.Command2.Size = New System.Drawing.Size(97, 25)
		Me.Command2.Location = New System.Drawing.Point(168, 320)
		Me.Command2.TabIndex = 11
		Me.Command2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command2.BackColor = System.Drawing.SystemColors.Control
		Me.Command2.CausesValidation = True
		Me.Command2.Enabled = True
		Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command2.TabStop = True
		Me.Command2.Name = "Command2"
		Me.Text5.AutoSize = False
		Me.Text5.Size = New System.Drawing.Size(161, 25)
		Me.Text5.Location = New System.Drawing.Point(168, 72)
		Me.Text5.TabIndex = 0
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
		Me.Text5.Visible = True
		Me.Text5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text5.Name = "Text5"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "ADD"
		Me.Command1.Size = New System.Drawing.Size(65, 25)
		Me.Command1.Location = New System.Drawing.Point(264, 320)
		Me.Command1.TabIndex = 10
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Text4.AutoSize = False
		Me.Text4.Size = New System.Drawing.Size(161, 25)
		Me.Text4.Location = New System.Drawing.Point(168, 296)
		Me.Text4.TabIndex = 9
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
		Me.Text3.AutoSize = False
		Me.Text3.Size = New System.Drawing.Size(161, 27)
		Me.Text3.Location = New System.Drawing.Point(168, 264)
		Me.Text3.TabIndex = 6
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
		Me.Text2.AutoSize = False
		Me.Text2.Size = New System.Drawing.Size(161, 25)
		Me.Text2.Location = New System.Drawing.Point(168, 136)
		Me.Text2.TabIndex = 2
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
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(161, 25)
		Me.Text1.Location = New System.Drawing.Point(168, 104)
		Me.Text1.TabIndex = 1
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
		Me.Label9.BackColor = System.Drawing.Color.Transparent
		Me.Label9.Text = "RACK"
		Me.Label9.Size = New System.Drawing.Size(137, 25)
		Me.Label9.Location = New System.Drawing.Point(0, 232)
		Me.Label9.TabIndex = 18
		Me.Label9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label9.Enabled = True
		Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = False
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label9.Name = "Label9"
		Me.Label8.BackColor = System.Drawing.Color.Transparent
		Me.Label8.Text = "PUBLICATION"
		Me.Label8.Size = New System.Drawing.Size(137, 25)
		Me.Label8.Location = New System.Drawing.Point(0, 168)
		Me.Label8.TabIndex = 17
		Me.Label8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label8.Enabled = True
		Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label8.UseMnemonic = True
		Me.Label8.Visible = True
		Me.Label8.AutoSize = False
		Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label8.Name = "Label8"
		Me.Label7.BackColor = System.Drawing.Color.Transparent
		Me.Label7.Text = "SHELF NUMBER"
		Me.Label7.Size = New System.Drawing.Size(137, 25)
		Me.Label7.Location = New System.Drawing.Point(0, 200)
		Me.Label7.TabIndex = 16
		Me.Label7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.Enabled = True
		Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.Label6.BackColor = System.Drawing.Color.Transparent
		Me.Label6.Text = "BOOKID"
		Me.Label6.Size = New System.Drawing.Size(137, 25)
		Me.Label6.Location = New System.Drawing.Point(0, 72)
		Me.Label6.TabIndex = 15
		Me.Label6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label6.Enabled = True
		Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.Label5.BackColor = System.Drawing.Color.Transparent
		Me.Label5.Text = "    QUANTITY"
		Me.Label5.Size = New System.Drawing.Size(145, 25)
		Me.Label5.Location = New System.Drawing.Point(-8, 296)
		Me.Label5.TabIndex = 14
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label4.Text = "                        ADD BOOKS"
		Me.Label4.Size = New System.Drawing.Size(209, 33)
		Me.Label4.Location = New System.Drawing.Point(0, 0)
		Me.Label4.TabIndex = 13
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
		Me.Label3.BackColor = System.Drawing.Color.Transparent
		Me.Label3.Text = "   PRICE"
		Me.Label3.Size = New System.Drawing.Size(145, 25)
		Me.Label3.Location = New System.Drawing.Point(-8, 264)
		Me.Label3.TabIndex = 12
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.BackColor = System.Drawing.Color.Transparent
		Me.Label2.Text = "   AUTHOR"
		Me.Label2.Size = New System.Drawing.Size(145, 25)
		Me.Label2.Location = New System.Drawing.Point(-8, 136)
		Me.Label2.TabIndex = 8
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Text = "                        BOOK NAME "
		Me.Label1.Size = New System.Drawing.Size(209, 25)
		Me.Label1.Location = New System.Drawing.Point(-72, 104)
		Me.Label1.TabIndex = 7
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(DataGrid1)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(Text8)
		Me.Controls.Add(Text7)
		Me.Controls.Add(Text6)
		Me.Controls.Add(Command2)
		Me.Controls.Add(Text5)
		Me.Controls.Add(Command1)
		Me.Controls.Add(Text4)
		Me.Controls.Add(Text3)
		Me.Controls.Add(Text2)
		Me.Controls.Add(Text1)
		Me.Controls.Add(Label9)
		Me.Controls.Add(Label8)
		Me.Controls.Add(Label7)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
#Region "Upgrade Support"
	Public Sub VB6_AddADODataBinding()
		DataGrid1.DataSource = CType(Adodc1, MSDATASRC.DataSource)
	End Sub
	Public Sub VB6_RemoveADODataBinding()
		DataGrid1.DataSource = Nothing
	End Sub
#End Region 
End Class