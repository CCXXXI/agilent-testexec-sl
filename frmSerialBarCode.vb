Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmSerialBarCode
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
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
    Public WithEvents cmdStopScanContinous As System.Windows.Forms.Button
	Public WithEvents cmdStartScanContinuous As System.Windows.Forms.Button
	Public WithEvents cmdScanOnce As System.Windows.Forms.Button
	Public WithEvents comBarCode1 As AxMSCommLib.AxMSComm
	Public WithEvents tmrScantimer As System.Windows.Forms.Timer
	Public WithEvents tmrIOTimer As System.Windows.Forms.Timer
	Public WithEvents Line1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Friend WithEvents rtbAbout As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbStatus As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbSingleScan As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbContinuousScan As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSerialBarCode))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdStopScanContinous = New System.Windows.Forms.Button
        Me.cmdStartScanContinuous = New System.Windows.Forms.Button
        Me.cmdScanOnce = New System.Windows.Forms.Button
        Me.comBarCode1 = New AxMSCommLib.AxMSComm
        Me.tmrScantimer = New System.Windows.Forms.Timer(Me.components)
        Me.tmrIOTimer = New System.Windows.Forms.Timer(Me.components)
        Me.Line1 = New System.Windows.Forms.Label
        Me.rtbAbout = New System.Windows.Forms.RichTextBox
        Me.rtbStatus = New System.Windows.Forms.RichTextBox
        Me.rtbSingleScan = New System.Windows.Forms.RichTextBox
        Me.rtbContinuousScan = New System.Windows.Forms.RichTextBox
        CType(Me.comBarCode1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdStopScanContinous
        '
        Me.cmdStopScanContinous.BackColor = System.Drawing.SystemColors.Control
        Me.cmdStopScanContinous.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdStopScanContinous.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStopScanContinous.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStopScanContinous.Location = New System.Drawing.Point(19, 438)
        Me.cmdStopScanContinous.Name = "cmdStopScanContinous"
        Me.cmdStopScanContinous.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdStopScanContinous.Size = New System.Drawing.Size(113, 37)
        Me.cmdStopScanContinous.TabIndex = 2
        Me.cmdStopScanContinous.Text = "Stop Scanning Continously"
        '
        'cmdStartScanContinuous
        '
        Me.cmdStartScanContinuous.BackColor = System.Drawing.SystemColors.Control
        Me.cmdStartScanContinuous.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdStartScanContinuous.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStartScanContinuous.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStartScanContinuous.Location = New System.Drawing.Point(18, 350)
        Me.cmdStartScanContinuous.Name = "cmdStartScanContinuous"
        Me.cmdStartScanContinuous.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdStartScanContinuous.Size = New System.Drawing.Size(113, 37)
        Me.cmdStartScanContinuous.TabIndex = 1
        Me.cmdStartScanContinuous.Text = "Start Scanning Continously"
        '
        'cmdScanOnce
        '
        Me.cmdScanOnce.BackColor = System.Drawing.SystemColors.Control
        Me.cmdScanOnce.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdScanOnce.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdScanOnce.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdScanOnce.Location = New System.Drawing.Point(22, 256)
        Me.cmdScanOnce.Name = "cmdScanOnce"
        Me.cmdScanOnce.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdScanOnce.Size = New System.Drawing.Size(113, 37)
        Me.cmdScanOnce.TabIndex = 0
        Me.cmdScanOnce.Text = "Scan Once"
        '
        'comBarCode1
        '
        Me.comBarCode1.Enabled = True
        Me.comBarCode1.Location = New System.Drawing.Point(99, 556)
        Me.comBarCode1.Name = "comBarCode1"
        Me.comBarCode1.OcxState = CType(resources.GetObject("comBarCode1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.comBarCode1.Size = New System.Drawing.Size(38, 38)
        Me.comBarCode1.TabIndex = 3
        '
        'tmrScantimer
        '
        Me.tmrScantimer.Interval = 1
        '
        'tmrIOTimer
        '
        Me.tmrIOTimer.Interval = 1
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(21, 178)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(479, 1)
        Me.Line1.TabIndex = 4
        '
        'rtbAbout
        '
        Me.rtbAbout.BackColor = System.Drawing.Color.LightGray
        Me.rtbAbout.Location = New System.Drawing.Point(48, 24)
        Me.rtbAbout.Name = "rtbAbout"
        Me.rtbAbout.Size = New System.Drawing.Size(432, 136)
        Me.rtbAbout.TabIndex = 5
        Me.rtbAbout.Text = ""
        '
        'rtbStatus
        '
        Me.rtbStatus.Location = New System.Drawing.Point(232, 200)
        Me.rtbStatus.Name = "rtbStatus"
        Me.rtbStatus.Size = New System.Drawing.Size(240, 32)
        Me.rtbStatus.TabIndex = 6
        Me.rtbStatus.Text = ""
        '
        'rtbSingleScan
        '
        Me.rtbSingleScan.Location = New System.Drawing.Point(232, 256)
        Me.rtbSingleScan.Name = "rtbSingleScan"
        Me.rtbSingleScan.Size = New System.Drawing.Size(240, 40)
        Me.rtbSingleScan.TabIndex = 7
        Me.rtbSingleScan.Text = ""
        '
        'rtbContinuousScan
        '
        Me.rtbContinuousScan.Location = New System.Drawing.Point(232, 344)
        Me.rtbContinuousScan.Name = "rtbContinuousScan"
        Me.rtbContinuousScan.Size = New System.Drawing.Size(240, 136)
        Me.rtbContinuousScan.TabIndex = 8
        Me.rtbContinuousScan.Text = ""
        '
        'frmSerialBarCode
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(530, 598)
        Me.Controls.Add(Me.rtbContinuousScan)
        Me.Controls.Add(Me.rtbSingleScan)
        Me.Controls.Add(Me.rtbStatus)
        Me.Controls.Add(Me.rtbAbout)
        Me.Controls.Add(Me.cmdStopScanContinous)
        Me.Controls.Add(Me.cmdStartScanContinuous)
        Me.Controls.Add(Me.cmdScanOnce)
        Me.Controls.Add(Me.comBarCode1)
        Me.Controls.Add(Me.Line1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(115, 30)
        Me.Name = "frmSerialBarCode"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        CType(Me.comBarCode1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmSerialBarCode
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmSerialBarCode
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmSerialBarCode()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	'*****************************************************************************
	'frmSerialBarCode.frm
	'frmSerialBarCode.frx
	'
	'The frmSerialBarCode.frm form exists to support the use
	'of serial bar code readers that are programmable, like the
	'Symbolics LS 1220.  It is provided here as an example, but
	'is not "turned on" when TypicalOpUI is first used.  See documentation
	'in frmSerialBarCode and in modConfiguration.
	'*****************************************************************************
	
	'create an instance of the bar code control
	Private WithEvents Reader As clsHPTXSLSerialBarcode
	Dim LocalSource As String
	Private Sub frmSerialBarCode_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim S1 As String
		Dim S2 As String
		Dim S3 As String
		Dim S4 As String
		Dim S5 As String
		Dim S6 As String
		Dim S7 As String
		Dim S8 As String
		Dim S9 As String
		Dim s10 As String
		Dim s12 As String
		
		
		S1 = "This form is provided as a container for the controls and code necessary to interface with a "
		S2 = "a """"black box"""" bar code reader that can be both serially programmed, and read serially.  "
		S3 = "The tested example was a symbolics LS1220 reader.  "
		S4 = "Under normal operation, the bar code reader would be triggered via software from some event detected in the "
		S5 = "TxSL user interface, like a button push, or a digital input that indicates the UUT is in the fixture. "
		S6 = "Under normal operation, this form would have no visible presence in the test system, the controls "
		S7 = "placed on this form would not be visible.  "
		S8 = "To allow for easy standalone debugging, the user interface were provided provided.  To use this form standalone "
		S9 = "create a project, add this form, and include the class clsHPTXSLSerialBarCode to your project, "
		s10 = "plus the error handling module modError.bas."
		
		rtbAbout.Text = S1 & S2 & vbCrLf & S3 & S4 & S5 & S6 & vbCrLf & S7 & vbCrLf & S8 & S9
		
		On Error GoTo LocalErrorHandler
		LocalSource = ":" & Me.Name & ":Form_Load" 'Used to track where the error happened
		
		Reader = New clsHPTXSLSerialBarcode
		Reader.ScanTimer = tmrScantimer
		Reader.IOTimer = tmrIOTimer
		Reader.CommObject = comBarCode1
		
		If Reader.CreateIOSession Then
			rtbStatus.Text = "IO Session Created Successfully"
			If Reader.InitializeScanner Then
				rtbStatus.Text = "Scanner Initialized"
			Else
				Err.Raise(frmErrorDialog.TxSLOpuiError.ScannerNotInitialized, Err.Source & LocalSource, "Could Not Initialize Bar Code Reader")
			End If
		Else
			Err.Raise(frmErrorDialog.TxSLOpuiError.NoIoSession, Err.Source & LocalSource, "Could not open IO session for bar code reader")
		End If
		
		Reader.EndScan() 'turn the laser off, just in case it was left on
		
		Exit Sub
LocalErrorHandler: 
		frmErrorDialog.ErrorHandler(Err.Number, Err.Source, Err.Description)
	End Sub
	
	Private Sub Reader_CodeScanned(ByRef BarCode As String) Handles Reader.CodeScanned
		On Error GoTo LocalErrorHandler
		LocalSource = ":" & Me.Name & ":Reader_CodeScanned" 'Used to track where the error happened
		
		If BarCode = "" Then
			rtbContinuousScan.Text = rtbContinuousScan.Text & "No bar code read"
			Err.Raise(frmErrorDialog.TxSLOpuiError.EmptyBarCode, Err.Source & LocalSource, "The bar code is empty")
		Else
			rtbContinuousScan.Text = rtbContinuousScan.Text & BarCode
		End If
		
		Exit Sub
LocalErrorHandler: 
		frmErrorDialog.ErrorHandler(Err.Number, Err.Source, Err.Description)
		
	End Sub
	
	Public Sub cmdScanOnce_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdScanOnce.Click
		cmdStartScanContinuous.Enabled = False
		cmdStopScanContinous.Enabled = False
		rtbStatus.Text = "Single Scan"
		rtbSingleScan.Text = ""
		Dim BarCode As String
		BarCode = Reader.ScanCode(2) 'The 2 is the time out value
		
		If BarCode = "" Then
			rtbSingleScan.Text = "Timeout, no scan received"
		Else
			rtbSingleScan.Text = BarCode
		End If
		cmdStartScanContinuous.Enabled = True
	End Sub
	
	Public Sub cmdStartScanContinuous_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStartScanContinuous.Click
		rtbStatus.Text = "Continuous Scan"
		rtbContinuousScan.Text = ""
		cmdScanOnce.Enabled = False
		cmdStopScanContinous.Enabled = True
		Reader.ContinuousScan()
	End Sub
	
	Public Sub cmdStopScanContinous_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStopScanContinous.Click
		Reader.EndScan()
		cmdScanOnce.Enabled = True
		cmdStartScanContinuous.Enabled = True
		cmdStopScanContinous.Enabled = False
		rtbStatus.Text = "Scanner Initialized"
	End Sub
	
	Private Sub comBarCode1_OnComm(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comBarCode1.OnComm
		'On the OnComm event from the bar code, call the Reader.CommCallBack function
		'If a good code, this results in the clsHPTxSLSerialBarCode event being raised
		
		Reader.CommCallback()
	End Sub
	
	Private Sub frmSerialBarCode_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		Reader.EndScan() 'Turn off the laser
		Reader.EndIOSession()
	End Sub
	
	Private Sub tmrIOTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrIOTimer.Tick
		Reader.IOTimeoutCallback()
	End Sub
	
	Private Sub tmrScantimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrScantimer.Tick
		Reader.ScanTimeoutCallback()
	End Sub
	
	Public Function StripTrailingCrLF(ByRef BarCode As String) As String
		'Used to strip a trailing CrLf from the bar code, prior to sending it to the
		'handle bar code function.
		Dim Position As Short
		Position = InStr(BarCode, vbCrLf)
		BarCode = VB.Left(BarCode, Position - 1)
		StripTrailingCrLF = BarCode
	End Function
End Class