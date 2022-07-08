Option Strict Off
Option Explicit On
Friend Class frmPreDefinedTxSLUserMessages
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
    Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdNegative As System.Windows.Forms.Button
	Public WithEvents cmdPositive As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Friend WithEvents rtbAdjust As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbMessage As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdNegative = New System.Windows.Forms.Button
        Me.cmdPositive = New System.Windows.Forms.Button
        Me.rtbMessage = New System.Windows.Forms.RichTextBox
        Me.rtbAdjust = New System.Windows.Forms.RichTextBox
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(368, 168)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(105, 33)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdNegative
        '
        Me.cmdNegative.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNegative.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNegative.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNegative.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNegative.Location = New System.Drawing.Point(200, 168)
        Me.cmdNegative.Name = "cmdNegative"
        Me.cmdNegative.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNegative.Size = New System.Drawing.Size(105, 33)
        Me.cmdNegative.TabIndex = 1
        Me.cmdNegative.Text = "No"
        '
        'cmdPositive
        '
        Me.cmdPositive.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPositive.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPositive.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPositive.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPositive.Location = New System.Drawing.Point(24, 168)
        Me.cmdPositive.Name = "cmdPositive"
        Me.cmdPositive.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPositive.Size = New System.Drawing.Size(105, 33)
        Me.cmdPositive.TabIndex = 0
        Me.cmdPositive.Text = "Yes"
        '
        'rtbMessage
        '
        Me.rtbMessage.Location = New System.Drawing.Point(16, 16)
        Me.rtbMessage.Name = "rtbMessage"
        Me.rtbMessage.Size = New System.Drawing.Size(472, 96)
        Me.rtbMessage.TabIndex = 3
        Me.rtbMessage.Text = ""
        '
        'rtbAdjust
        '
        Me.rtbAdjust.Location = New System.Drawing.Point(184, 128)
        Me.rtbAdjust.Name = "rtbAdjust"
        Me.rtbAdjust.Size = New System.Drawing.Size(144, 24)
        Me.rtbAdjust.TabIndex = 4
        Me.rtbAdjust.Text = ""
        '
        'frmPreDefinedTxSLUserMessages
        '
        Me.AcceptButton = Me.cmdPositive
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(505, 217)
        Me.Controls.Add(Me.rtbAdjust)
        Me.Controls.Add(Me.rtbMessage)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdNegative)
        Me.Controls.Add(Me.cmdPositive)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(201, 257)
        Me.Name = "frmPreDefinedTxSLUserMessages"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "User Message"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmPreDefinedTxSLUserMessages
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmPreDefinedTxSLUserMessages
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmPreDefinedTxSLUserMessages()
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
	'frmPreDefinedTxSLUserMessages.frm
	'frmPreDefinedTxSLUserMessages.frx
	'
	'The frmPreDefinedTxSLUserMessages.frm form handles predefined user messages that
	'may be sent from TxSL testplans.  In general, these
	'are dialogs, coming from tests, that are asking the operator
	'some question, that can be answered by
	'OK, OK/Cancel, Yes/No or Yes/No/Cancel.
	
	'As Txsl has evolved, we have added user messages beyound the scope of simple
	'dialogs.  Hence, this file has evolved to be two purposes
	'1--It documents the reserved numbers for ALL user messages.
	'2--It holds the implementation code for responding to user messages
	'that set up dialogs
	'Note that the implementation code for NON-dialog related user messages
	'has been left to other forms.
	
	'*****************************************************************************
	
	Private LocalSource As String
	
	'These next two booleans are provided to allow
	'state tracking of what dialog we are displaying
	'For example, if we are doing Ok (only), we will reject
	'any button clicks (with a beep) of anything but an ok
	Private NegativeEnabled As Boolean
	Private CancelEnabled As Boolean
	
	'A local storage of the messageid
	Private LocalMessageID As txslMessageID
	
	Public Enum txslMessageID
		'An enumeration of predefined values, for use in user
		'defined messages, either sent by the operator interface via the
		'SendUserMessage method
		'or sent to the operator interface via the UserMessageEvent
		'Numbering Guidelines
		'HP has reserved the numbers from 10,000 to 99,999.
		'Within this block, groups of 10,000 are segmented into message categories.
		'The following message categories are defined.
		'WRITE Digital Output   10000
		'READ Digital Input     20000
		'Digital Input EVENT    30000
		'Input RESPONSE         40000
		'TxSL special needs     50000
		'  TxSL DataLogging Customization Starting at 50001
		'  TxSL Throughput Multiplier Messages Starting at 50101
		'Future                 60000
		'Future                 70000
		'Telecomm Systems       80000
		'Automotive Systems     90000
		'
		'Within each category of 10,000, the "tens" digit is used to represent
		'an individual message.
		'
		'With each individual message (a group of 10), the "ones" digit is used to represent
		'sub messages appropriate for that message.
		'Specifically, for Digital Input Event messages, the following submessages
		'are recommended.
		'0 = the Enable message
		'1 = the Found messsage
		'2 = the Lost Message
		'
		'For InputResponse Messages, the following submessages are recommended:
		'0= Prompt Available
		'1= Prompt
		'2= Adjust (Adjust Message only)
		'9= Done
		'With all that, here are the predefined numbers
		'Used for WRITING digital outputs
		IndicateOK = 10010
		IndicateWarning = 10020
		IndicateRequestAssistance = 10030
		IndicateTestplanResultPassFail = 10040
		IndicateMachineNotReady = 10050
		IndicateBoardAvailable = 10060
		'Used for READING digital inputs
		ReadRunButton = 20010
		ReadStopButton = 20020
		ReadAbortButton = 20030
		ReadOK = 20040
		ReadCancel = 20050
		ReadYes = 20060
		ReadNo = 20070
		ReadSafetyInput = 20080
		ReadUUTDetected = 20090
		ReadUpStreamBoardAvailable = 20100
		ReadDownStreamMachineNotReady = 20110
		'Used for enable and reading EVENTS
		'Enable is used to enable the monitoring of the input
		'The Found message indicates that the input made a false to true transition
		'The Lost message indicates that the input made a True to False transition
		RunEnable = 30010
		StopEnable = 30020
		AbortEnable = 30030
		OKEnable = 30040
		CancelEnable = 30050
		YesEnable = 30060
		NoEnable = 30070
		SafetyInputEnable = 30080
		UUTDetectedEnable = 30090
		UpStreamBoardAvailableEnable = 30100
		DownStreamMachineNotReadyEnable = 30110
		FoundOK = 30041
		FoundCancel = 30051
		FoundYes = 30061
		FoundNo = 30071
		FoundSafetyInput = 30081
		FoundUUTDetected = 30091
		FoundUpStreamBoardAvailable = 30101
		FoundDownStreamMachineNotReady = 30111
		LostOK = 30042
		LostCancel = 30052
		LostYes = 30062
		LostNo = 30072
		LostSafetyInput = 30082
		LostUUTDetected = 30092
		LostUpStreamBoardAvailable = 30102
		LostDownStreamMachineNotReady = 30112
		'Used for response inputs
		OKPromptAvailable = 40010
		OKPrompt = 40011
		OKDone = 40019
		OKCancelPromptAvailable = 40020
		OKCancelPrompt = 40021
		OkCancelDone = 40029
		YesNoPromptAvailable = 40030
		YesNoPrompt = 40031
		YesNoDone = 40039
		YesNoCancelPromptAvailable = 40040
		YesNoCancelPrompt = 40041
		YesNoCancelDone = 40049
		AdjustPromptAvailable = 40050
		AdjustPrompt = 40051
		AdjustUpdateValue = 40052
		AdjustDone = 40059
		'The 50,000 block is used for user defined messages not covered in
		'the above categories, but are specific to TxSL customization.
		'TxSL DataLog Customization Messages
		AfterDataLogCreateDone = 50001
		AfterDataLogFileWriteDone = 50002
		DataLogFileWriteEnabled = 50010
		UserDefinedFileName = 50011
		UserDefinedFileNamePrefix = 50012
		UserDefinedFileNameSuffix = 50013
		'Txsl Throughput Multiplier Messages
		ReadUutSerialNumbersFromTestplan = 50101
		ReadUutXoutFlagsFromTestplan = 50102
	End Enum
	
	Public Function UserDefinedMessage(ByVal MessageID As Integer, ByVal Message As String) As Boolean
		'A function to handle messages that were sent to the operator interface that
		'request us to post dialogs to the user.
		'This function in particular checks for HP Predefined messages and
		'returns true when the message is handled.
		'It will return false if this is an not a predefined TxsL message relating to dialogs, or
		
		Select Case MessageID
			
			'Customize:  Add message handling here as needed, if you add more dialogs
			'Other messages may be better handled in a context closer to where they make
			'sense
			
			'Note that sending a user defined query (from any source) results in a message being
			'broadcast to the entire system, including back to yourself.
			
			'The following defines the responses to standard dialog actions built into TxSL.
			Case Is = txslMessageID.OKPromptAvailable, txslMessageID.OKCancelPromptAvailable, txslMessageID.YesNoPromptAvailable, txslMessageID.YesNoCancelPromptAvailable, txslMessageID.AdjustPromptAvailable
                frmMain.TestExecSL1.SendUserDefinedResponse(MessageID, "Available")
                UserDefinedMessage = True

            Case Is = txslMessageID.OKPrompt, txslMessageID.OKCancelPrompt, txslMessageID.YesNoPrompt, txslMessageID.YesNoCancelPrompt, txslMessageID.AdjustPrompt
                PutUp(MessageID, Message)
                UserDefinedMessage = True

            Case Is = txslMessageID.OKDone, txslMessageID.OkCancelDone, txslMessageID.YesNoDone, txslMessageID.YesNoCancelDone, txslMessageID.AdjustDone
                TakeDown()
                UserDefinedMessage = True

            Case txslMessageID.AdjustUpdateValue
                UpdateAdjustValue((Message))
                UserDefinedMessage = True

            ' This is an example of how to rename a log file.  Uncomment out the next 3 lines.
            ' Case txslMessageID.AfterDataLogFileWriteDone
            '   System.IO.File.Move(Message, "c:\tmptxsllogfile.txt")
            '   UserDefinedMessage = True
            Case Else
                'This is the case of a non-predefined message.
                'Return false, and allow other message handlers in the system to have a crack at it
                UserDefinedMessage = False

        End Select
    End Function


    Public Sub PutUp(ByRef MessageID As txslMessageID, ByRef Message As String)
        'A sub that puts up the form, customized to reflect the type
        'of dialog that was requested.

        'This saves the messageid away, for use later in responses.
        LocalMessageID = MessageID

        Select Case MessageID
            Case txslMessageID.OKPrompt
                'Set the state:
                NegativeEnabled = False
                CancelEnabled = False

                cmdPositive.Text = LangLookup(modLocalization.txslLangIndex.gnOK)
                cmdPositive.Visible = True
                cmdNegative.Visible = False
                cmdCancel.Visible = False
                rtbMessage.Text = Message
                rtbAdjust.Visible = False

            Case txslMessageID.OKCancelPrompt
                'Set the state:
                NegativeEnabled = False
                CancelEnabled = True

                cmdPositive.Text = LangLookup(modLocalization.txslLangIndex.gnOK)
                cmdNegative.Visible = False
                cmdCancel.Visible = True
                cmdCancel.Text = LangLookup(modLocalization.txslLangIndex.gnCancel)
                rtbMessage.Text = Message
                rtbAdjust.Visible = False

            Case txslMessageID.YesNoPrompt
                'Set the state:
                NegativeEnabled = True
                CancelEnabled = False


                cmdPositive.Text = LangLookup(modLocalization.txslLangIndex.gnYes)
                cmdPositive.Visible = True
                cmdNegative.Visible = True
                cmdNegative.Text = LangLookup(modLocalization.txslLangIndex.gnNo)
                cmdCancel.Visible = False
                rtbMessage.Text = Message
                rtbAdjust.Visible = False

            Case txslMessageID.YesNoCancelPrompt
                'Set the state:
                NegativeEnabled = True
                CancelEnabled = True

                cmdPositive.Text = LangLookup(modLocalization.txslLangIndex.gnYes)
                cmdPositive.Text = "Yes"
                cmdPositive.Visible = True
                cmdNegative.Visible = True
                cmdNegative.Text = LangLookup(modLocalization.txslLangIndex.gnNo)
                cmdCancel.Visible = True
                rtbMessage.Text = Message
                rtbAdjust.Visible = False

            Case txslMessageID.AdjustPrompt
                'Set the state:
                NegativeEnabled = True
                CancelEnabled = True

                cmdPositive.Text = LangLookup(modLocalization.txslLangIndex.gnYes)
                cmdPositive.Visible = True
                cmdNegative.Visible = True
                cmdNegative.Text = LangLookup(modLocalization.txslLangIndex.gnNo)
                cmdCancel.Visible = True
                rtbMessage.Text = Message
                rtbAdjust.Visible = True
        End Select

        Show()

    End Sub

    Public Sub TakeDown()
        Hide()
    End Sub

    Public Sub UpdateAdjustValue(ByRef Message As String)
        rtbAdjust.Text = Message
    End Sub

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click 'TxSLErrorTrapSite
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdCancel_Click"
        If CancelEnabled Then
            Select Case LocalMessageID
                Case Is = txslMessageID.OKPrompt, txslMessageID.YesNoPrompt
                    MsgBox(LocalSource & "Should not get here")
                Case Is = txslMessageID.OKCancelPrompt, txslMessageID.YesNoCancelPrompt, txslMessageID.AdjustPrompt
                    frmMain.TestExecSL1.SendUserDefinedResponse(LocalMessageID, "Cancel")
				Case Else
					MsgBox(LocalSource & "Unknown Message ID")
			End Select
			Hide()
		Else 'beep and don't take down the screen if can't accept this input
			Beep()
		End If
		
		Exit Sub
LocalErrorHandler: 
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
    End Sub

    Private Sub cmdNegative_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNegative.Click 'TxSLErrorTrapSite
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdNegative_Click"
        If NegativeEnabled Then
            Select Case LocalMessageID
                Case Is = txslMessageID.OKPrompt, txslMessageID.OKCancelPrompt
                    MsgBox("Should not get here")
                Case Is = txslMessageID.YesNoPrompt, txslMessageID.YesNoCancelPrompt, txslMessageID.AdjustPrompt
                    frmMain.TestExecSL1.SendUserDefinedResponse(LocalMessageID, "No")
                Case Else
                    MsgBox("Unknown Message ID")
            End Select
            Hide()
        Else
            Beep() ''beep and don't take down the screen if can't accept this input
        End If

        Exit Sub
LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
    End Sub

    Private Sub cmdPositive_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPositive.Click 'TxSLErrorTrapSite
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdPositive_Click"
        'This choice is always enabled.
        Select Case LocalMessageID
            Case Is = txslMessageID.OKPrompt, txslMessageID.OKCancelPrompt
                frmMain.TestExecSL1.SendUserDefinedResponse(LocalMessageID, "OK")
            Case Is = txslMessageID.YesNoPrompt, txslMessageID.YesNoCancelPrompt, txslMessageID.AdjustPrompt
                frmMain.TestExecSL1.SendUserDefinedResponse(LocalMessageID, "Yes")
            Case Else
                MsgBox("Unknown Message ID")
        End Select
        Hide()
        Exit Sub
LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
    End Sub
	
	Private Sub frmPreDefinedTxSLUserMessages_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'This sub (when keypreview on the form is also enabled)
		'allows us to intercept keyboard events and handle
		'them appropriate.  It's major use is for the keypad
		'State tracking is provided in the click events, to
		'reject keypresses when not appropriate.
		
		Select Case KeyCode
			Case System.Windows.Forms.Keys.A
				If (Shift = (VB6.ShiftConstants.ShiftMask + VB6.ShiftConstants.CtrlMask + VB6.ShiftConstants.AltMask)) Then
					cmdPositive_Click(cmdPositive, New System.EventArgs())
				End If
			Case System.Windows.Forms.Keys.S
				If (Shift = (VB6.ShiftConstants.ShiftMask + VB6.ShiftConstants.CtrlMask + VB6.ShiftConstants.AltMask)) Then
					cmdNegative_Click(cmdNegative, New System.EventArgs())
				End If
			Case System.Windows.Forms.Keys.D
				'not used
			Case System.Windows.Forms.Keys.F
				If (Shift = (VB6.ShiftConstants.ShiftMask + VB6.ShiftConstants.CtrlMask + VB6.ShiftConstants.AltMask)) Then
					cmdCancel_Click(cmdCancel, New System.EventArgs())
				End If
			Case Else
				'do nothing
		End Select
	End Sub
End Class