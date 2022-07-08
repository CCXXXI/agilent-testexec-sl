Option Strict Off
Option Explicit On
Friend Class frmStripPrinter
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
	Public WithEvents cmdTest As System.Windows.Forms.Button
    Public WithEvents comStripPrinter As AxMSCommLib.AxMSComm
    'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Friend WithEvents rtbAbout As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbTestString As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmStripPrinter))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdTest = New System.Windows.Forms.Button
        Me.comStripPrinter = New AxMSCommLib.AxMSComm
        Me.rtbAbout = New System.Windows.Forms.RichTextBox
        Me.rtbTestString = New System.Windows.Forms.RichTextBox
        CType(Me.comStripPrinter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdTest
        '
        Me.cmdTest.BackColor = System.Drawing.SystemColors.Control
        Me.cmdTest.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdTest.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdTest.Location = New System.Drawing.Point(72, 344)
        Me.cmdTest.Name = "cmdTest"
        Me.cmdTest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdTest.Size = New System.Drawing.Size(87, 46)
        Me.cmdTest.TabIndex = 2
        Me.cmdTest.Text = "Test"
        '
        'comStripPrinter
        '
        Me.comStripPrinter.Enabled = True
        Me.comStripPrinter.Location = New System.Drawing.Point(2, 431)
        Me.comStripPrinter.Name = "comStripPrinter"
        Me.comStripPrinter.OcxState = CType(resources.GetObject("comStripPrinter.OcxState"), System.Windows.Forms.AxHost.State)
        Me.comStripPrinter.Size = New System.Drawing.Size(38, 38)
        Me.comStripPrinter.TabIndex = 3
        '
        'rtbAbout
        '
        Me.rtbAbout.BackColor = System.Drawing.Color.LightGray
        Me.rtbAbout.Location = New System.Drawing.Point(24, 24)
        Me.rtbAbout.Name = "rtbAbout"
        Me.rtbAbout.Size = New System.Drawing.Size(552, 136)
        Me.rtbAbout.TabIndex = 4
        Me.rtbAbout.Text = "A container for the controls and code to handle a strip printer.  See the underly" & _
        "ing code for comments."
        '
        'rtbTestString
        '
        Me.rtbTestString.Location = New System.Drawing.Point(24, 192)
        Me.rtbTestString.Name = "rtbTestString"
        Me.rtbTestString.Size = New System.Drawing.Size(552, 136)
        Me.rtbTestString.TabIndex = 5
        Me.rtbTestString.Text = ""
        '
        'frmStripPrinter
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(602, 476)
        Me.Controls.Add(Me.rtbTestString)
        Me.Controls.Add(Me.rtbAbout)
        Me.Controls.Add(Me.cmdTest)
        Me.Controls.Add(Me.comStripPrinter)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(57, 23)
        Me.Name = "frmStripPrinter"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Strip Printer"
        CType(Me.comStripPrinter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmStripPrinter
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmStripPrinter
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmStripPrinter()
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
	'frmStripPrinter.frm
	'frmStripPrinter.frx
	'
	'The frmStripPrinter.frm is provided to handle communication
	'to a strip printer.  The TypicalOpUI application will use
	'this form to do minor formatting of the report stream, and
	'print them to the strip printer.  Note that the strip printer
	'is NOT enabled in the default configuration, and should be
	'enabled in the modConfiguration.bas file.
	'*****************************************************************************
	Private Const ExtraLines As Short = 8 'This defines the number of extra
	'Lines to print, following a print
	
	Private Sub cmdTest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTest.Click
		'Test operation
		AfterTestplanStop((rtbTestString.Text))
	End Sub
	
	Private Sub frmStripPrinter_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim S1 As String
		Dim S2 As String
		Dim S3 As String
		Dim S4 As String
		Dim S5 As String
		Dim S6 As String
		Dim S7 As String
		Dim S8 As String
		
		S1 = "This form is provided as a container for the controls and code to use a strip printer.  "
		S2 = "The strip printer that is targetted is the HP1199B, although given the simplicity of the code, "
		S3 = "you could easily support other printers.  "
		S4 = "The enhanced typical application is coded to react to the AfterTestplanStop event, "
		S5 = "and (depending on the configuration variables for printing) to send the contents of the report window to the strip printer.  "
		S6 = "You can modify report printing to your own scheme, and still use this utility.  "
		S7 = "The command button and input box on this form are provided only for easy diagnostics of this form, allowing operation "
		S8 = "to be tested standalone.  The command button and input box are not used in normal operation.  "
		
		rtbAbout.Text = S1 & S2 & S3 & vbCrLf & S4 & S5 & S6 & vbCrLf & S7 & S8
		
		InitializeCommPort()
		
	End Sub
	
	Public Sub InitializeCommPort()
		'Initializes the comm port, and will overwrite properties set in the property window
		'Printer settings should be set appropriately to match
		
		comStripPrinter.CommPort = 1 'Sets the port number.  This will override the default
		comStripPrinter.DTREnable = True
		comStripPrinter.RTSEnable = True
		comStripPrinter.EOFEnable = False 'The default for the comm port
		comStripPrinter.Handshaking = MSCommLib.HandshakeConstants.comNone 'The default for the comm port
		comStripPrinter.InputMode = MSCommLib.InputModeConstants.comInputModeText 'The default for the comm port
		comStripPrinter.NullDiscard = True 'The default is for the comm port if false
		comStripPrinter.Settings = "9600,N,8,1" 'The defaults for the comm port
		comStripPrinter.PortOpen = True 'Opens the comm to printer connection
		
	End Sub
	
	Public Sub PrintInput(ByRef InputString_Renamed As String)
		'Takes incoming strings and put them out to the printer.
		'The state machine is provided to take the report string set of CR/LFs and
		'convert them for good display on the strip printer.
		Dim chnum, InputStringLength As Integer
		Dim ch As String
		Dim foundcr As Boolean
		Dim code As Short
		foundcr = False
		InputStringLength = Len(InputString_Renamed)
		' state machine.  we keep track of if the previous character was a carraige
		' return (foundcr) and if we receive a carriage return linefeed pair, we send
		' a cr-lf to the printer.  If we receive a lf not proceeded by a cr, we send
		' a cr-lf.  If we receive a cr not followed by a lf, we send a cr-lf
		' cr = chr(13)   lf = chr(10)
		For chnum = 1 To InputStringLength
			ch = Mid(InputString_Renamed, chnum, 1)
			code = Asc(ch)
			If code = 10 Then
				comStripPrinter.Output = Chr(13) & Chr(10)
				foundcr = False
			ElseIf code = 13 And foundcr = True Then 
				comStripPrinter.Output = Chr(13) & Chr(10)
			ElseIf code = 13 And foundcr = False Then 
				foundcr = True
			ElseIf code <> 13 And foundcr = True Then 
				comStripPrinter.Output = Chr(13) & Chr(10)
				comStripPrinter.Output = ch
				foundcr = False
			Else
				comStripPrinter.Output = ch
			End If
		Next 
		If foundcr = True Then
			comStripPrinter.Output = Chr(13) & Chr(10)
		End If
	End Sub
	
	Public Sub AfterTestplanStop(ByRef Report As String)
		'This is named AfterTestplanStop to mirror the most likely used application
		'of the printer, that is--to print out strings following an AfterTestplanStop
		'event from TxSl.
		
		PrintInput(Report)
		
		'Print out an extra cr/lf's, to allow the paper to be ripped off.
		Dim I As Short
		For I = 1 To ExtraLines
			PrintInput(vbCrLf)
		Next 
	End Sub
	
	Private Sub frmStripPrinter_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		comStripPrinter.PortOpen = False
	End Sub
End Class