<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form4
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Text4 As System.Windows.Forms.TextBox
	Public WithEvents Adodc2 As VB6.ADODC
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Text3 As System.Windows.Forms.TextBox
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form4))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Text4 = New System.Windows.Forms.TextBox
		Me.Adodc2 = New VB6.ADODC
		Me.Adodc1 = New VB6.ADODC
		Me.Command2 = New System.Windows.Forms.Button
		Me.Text3 = New System.Windows.Forms.TextBox
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.Command1 = New System.Windows.Forms.Button
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "ISSUE"
		Me.ClientSize = New System.Drawing.Size(536, 418)
		Me.Location = New System.Drawing.Point(8, 31)
		Me.Icon = CType(resources.GetObject("Form4.Icon"), System.Drawing.Icon)
		Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		Me.BackgroundImage = CType(resources.GetObject("Form4.BackgroundImage"), System.Drawing.Image)
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
		Me.Name = "Form4"
		Me.Text4.AutoSize = False
		Me.Text4.Size = New System.Drawing.Size(81, 33)
		Me.Text4.Location = New System.Drawing.Point(280, 264)
		Me.Text4.TabIndex = 9
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
		Me.Adodc2.Size = New System.Drawing.Size(97, 49)
		Me.Adodc2.Location = New System.Drawing.Point(16, 248)
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
		Me.Adodc1.Size = New System.Drawing.Size(97, 49)
		Me.Adodc1.Location = New System.Drawing.Point(144, 296)
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
		Me.Adodc1.RecordSource = "issue"
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		Me.Command2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command2.Text = "MAIN MENU"
		Me.Command2.Size = New System.Drawing.Size(97, 33)
		Me.Command2.Location = New System.Drawing.Point(296, 192)
		Me.Command2.TabIndex = 8
		Me.Command2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command2.BackColor = System.Drawing.SystemColors.Control
		Me.Command2.CausesValidation = True
		Me.Command2.Enabled = True
		Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command2.TabStop = True
		Me.Command2.Name = "Command2"
		Me.Text3.AutoSize = False
		Me.Text3.Enabled = False
		Me.Text3.Size = New System.Drawing.Size(273, 33)
		Me.Text3.Location = New System.Drawing.Point(120, 160)
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
		Me.Text3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text3.Name = "Text3"
		Me.Text2.AutoSize = False
		Me.Text2.Size = New System.Drawing.Size(273, 33)
		Me.Text2.Location = New System.Drawing.Point(120, 80)
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
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(273, 33)
		Me.Text1.Location = New System.Drawing.Point(120, 120)
		Me.Text1.TabIndex = 2
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
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "ISSUE"
		Me.Command1.Size = New System.Drawing.Size(81, 33)
		Me.Command1.Location = New System.Drawing.Point(216, 192)
		Me.Command1.Image = CType(resources.GetObject("Command1.Image"), System.Drawing.Image)
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
		Me.Label4.BackColor = System.Drawing.Color.Transparent
		Me.Label4.Text = "DATE"
		Me.Label4.Size = New System.Drawing.Size(81, 33)
		Me.Label4.Location = New System.Drawing.Point(16, 160)
		Me.Label4.TabIndex = 7
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Label3.Text = "STUDENT NAME"
		Me.Label3.Size = New System.Drawing.Size(81, 33)
		Me.Label3.Location = New System.Drawing.Point(16, 120)
		Me.Label3.TabIndex = 6
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
		Me.Label2.Text = "BOOK CODE"
		Me.Label2.Size = New System.Drawing.Size(81, 33)
		Me.Label2.Location = New System.Drawing.Point(16, 80)
		Me.Label2.TabIndex = 5
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
		Me.Label1.Text = "                         ISSUE"
		Me.Label1.Size = New System.Drawing.Size(193, 17)
		Me.Label1.Location = New System.Drawing.Point(112, 0)
		Me.Label1.TabIndex = 0
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
		Me.Controls.Add(Text4)
		Me.Controls.Add(Adodc2)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(Command2)
		Me.Controls.Add(Text3)
		Me.Controls.Add(Text2)
		Me.Controls.Add(Text1)
		Me.Controls.Add(Command1)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class