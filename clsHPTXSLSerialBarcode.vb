Option Strict Off
Option Explicit On
Friend Class clsHPTXSLSerialBarcode
	'****************************************************************************
	'clsHPTXSLSerialBarcode.Cls
	'
	'This file defines a class for use in controlling serial bar code readers.
	'It is specifically designed around the Symbolics LS-1220.
	'It is included as an example
	'but is not used in "out of the box" operation of typical interface.
	'It is used in conjunction with the frmSerialBarCode.frm.
	'*****************************************************************************
	
	'A set of properties for the class
	Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Integer)
	Private bcScanRunning As Boolean
	Private bcComPort As Short
	Private bcBaudRate As Short
	Private bcParity As String
	Private bcDataBits As Short
	Private bcStopBits As Short
	'The next two variables are used to define handshaking for
	'the Serial interface.
	Private bcHardwareHandShaking As Boolean
	Private bcSoftwareHandShaking As Boolean
	
	'The bcAckNakOnProgram boolean is a configuration flag that
	'defines whether the reader has an Ack/Nak handshake upon programmming.
	'bcAckNakOnProgram is defined to conform to the Symbol Technologies behavior
	'of having an Ack/Nak handshake upon every program.
	Private bcAckNakOnProgram As Boolean
	
	'The bcCodeTerminator string is the string that is appended by the
	'reader to the end of a bar code string.  It is used to indicate
	'that a full and complete code was read
	Private bcCodeTerminator As String
	
	'Other reader properties
	Private bcLaserOnCmd As String
	Private bcLaserOffCmd As String
	Private bcInitCmd As String
	Private bcAckString As String
	Private bcNAckString As String
	Private bcInputBuff As String
	Private bcScanTimedOut As Boolean
	Private bcIOTimedOut As Boolean
	Private bcIOTimeout As Short
	Private bcLastCode As String
	Private bcIOTimer As System.Windows.Forms.Timer
	Private bcScanTimer As System.Windows.Forms.Timer
	Private bcComm As AxMSCommLib.AxMSComm
	Public Event CodeScanned(ByRef code As String)
	
	
	'**************************************************************************
	' Opens the com port to talk to the barcode reader.  Do this after you
	' have specified the host i/o parameters.  Returns true if successful
	Public Function CreateIOSession() As Boolean
		bcComm.CommPort = bcComPort
		bcComm.DTREnable = False
		bcComm.EOFEnable = False
		bcComm.InputMode = MSCommLib.InputModeConstants.comInputModeText
		bcComm.NullDiscard = False
		
		'Set the handshaking modes for the com port.
		If bcHardwareHandShaking And bcSoftwareHandShaking Then
			bcComm.Handshaking = MSCommLib.HandshakeConstants.comRTSXOnXOff
		ElseIf bcHardwareHandShaking Then 
			bcComm.Handshaking = MSCommLib.HandshakeConstants.comRTS
		ElseIf bcSoftwareHandShaking Then 
			bcComm.Handshaking = MSCommLib.HandshakeConstants.comXOnXoff
		Else
			bcComm.Handshaking = MSCommLib.HandshakeConstants.comNone
		End If
		
		bcComm.RThreshold = 0 'Disables the event on the receive buffer
		bcComm.SThreshold = 0 'Disables the event on the transmit buffer
		
		Dim SettingString As String
		SettingString = CStr(bcBaudRate) & "," & bcParity & "," & CStr(bcDataBits) & "," & CStr(bcStopBits)
		bcComm.Settings = SettingString
		
		bcComm.PortOpen = True
		If bcComm.PortOpen Then
			CreateIOSession = True
		End If
		
	End Function
	'**************************************************************************
	Public Sub EndIOSession()
		bcComm.PortOpen = False
	End Sub
	'**************************************************************************
	' Sends the init string to the barcode reader.  Do this after you have
	' set the init string and opened an I/O session.  Returns true if successful.
	Public Function InitializeScanner() As Boolean
		bcComm.InBufferCount = 0 'Clears the input buffer
		bcComm.Output = bcInitCmd 'Initialize the barcode
		If bcAckNakOnProgram Then
			InitializeScanner = WaitForAck
		Else
			InitializeScanner = True
		End If
		
		'Customize:
		Sleep((2000)) 'Hardcode a sleep, to allow the reader to recover
		
	End Function
	'**************************************************************************
	Public Function ContinuousScan() As Boolean
		If bcScanRunning = False Then
			bcScanRunning = True
			bcComm.Output = bcLaserOnCmd
			If bcAckNakOnProgram = True Then ContinuousScan = WaitForAck
			bcComm.RThreshold = 1 'this enables the OnComm event
		Else
			ContinuousScan = False
		End If
	End Function
	'**************************************************************************
	' stops continuous scan
	Public Sub EndScan()
		If bcScanRunning = True Then
			bcScanRunning = False
			bcComm.RThreshold = 0 'this disables the OnComm event
			bcComm.Output = bcLaserOffCmd
			If bcAckNakOnProgram = True Then WaitForAck()
		End If
	End Sub
	'**************************************************************************
	' last successfully scanned code
	Public Function GetLastCode() As String
		GetLastCode = bcLastCode
	End Function
	
	'**************************************************************************
	' Turns on the laser, waits waitSeconds for a scan, returns the contents of the receive
	' buffer and turns off the laser.
	'An empty string  "" is returned if
	'1) the reader did not respond before the time out.
	'2) If it was already running
	'3) If the comm port was not open
	Public Function ScanCode(ByRef waitSeconds As Short) As String
		If bcScanRunning = True Or Not bcComm.PortOpen Then
			ScanCode = ""
			Beep()
			Exit Function
		Else
			bcScanRunning = True
		End If
		bcComm.RThreshold = 0
		bcComm.InBufferCount = 0 'Clear the input buffer
		bcInputBuff = ""
		bcComm.Output = bcLaserOnCmd
		If bcAckNakOnProgram = False Then
			ReceiveScanCode((waitSeconds))
		ElseIf bcAckNakOnProgram = True Then 
			If WaitForAck = True Then ' wait for an ack before receiving code
				ReceiveScanCode((waitSeconds))
			End If
			
		End If
		ScanCode = bcInputBuff
		bcInputBuff = ""
		bcComm.Output = bcLaserOffCmd
		If bcAckNakOnProgram = True Then
			WaitForAck()
		End If
		bcScanRunning = False
	End Function
	'**************************************************************************
	' Waits for an acknowledgement to come in the input buffer.  returns true if an ack
	' was received within the io timeout limit.  Returns false if it timed out, a nack
	' was received, or an unrecognized code was received.  If an ack or a nack is
	' received, it will be removed from the input buffer
	Private Function WaitForAck() As Boolean
		Dim bGotCode As Boolean
		Dim nacklen, acklen, maxlen As Short
		acklen = Len(bcAckString)
		nacklen = Len(bcNAckString)
		If (acklen > nacklen) Then maxlen = acklen Else maxlen = nacklen
		bcIOTimedOut = False ' set to true by the timer
		bcInputBuff = ""
		bcIOTimer.Interval = bcIOTimeout * 1000 'convert the number of seconds
		'to the number of milliseconds
		
		bcIOTimer.Enabled = True
		Do 
			System.Windows.Forms.Application.DoEvents()
			If bcComm.InBufferCount > 0 Then 'as we get characters, append them to the buffer. The input
				'method will clean out the buffer.
				bcInputBuff = bcInputBuff + bcComm.Input
			End If
			If acklen <= Len(bcInputBuff) Then
				bGotCode = IsAck(bcInputBuff) ' if we got an ack, leave loop
			End If
			If nacklen <= Len(bcInputBuff) Then
				bGotCode = IsNAck(bcInputBuff) ' if we got a nack, leave loop
			End If
			If maxlen <= Len(bcInputBuff) Then
				bGotCode = True ' we have an overrun, leave loop
			End If
		Loop Until bcIOTimedOut = True Or bGotCode = True
		bcIOTimer.Enabled = False
		
		'Figure out what to return, and generate errors if appropriate.
		If IsAck(bcInputBuff) = True Then
			' the command was sent successfully
			' remove the ack from the beginning of the buffer
			bcInputBuff = Mid(bcInputBuff, Len(bcAckString), Len(bcInputBuff) - Len(bcAckString))
			WaitForAck = True
		ElseIf IsNAck(bcInputBuff) = True Then  ' the command failed
			' remove the nack from the beginning of the buffer
			bcInputBuff = Mid(bcInputBuff, Len(bcNAckString), Len(bcInputBuff) - Len(bcNAckString))
			WaitForAck = False
		ElseIf Len(bcInputBuff) > 0 Then 
			' the data we received was neither an ack nor an nack, or we timed out
			'MsgBox "Neither ack or nack received, WaitforAck failed.  The string that was received was " + bcInputBuff
			'WaitForAck = False
		ElseIf bcIOTimedOut = True Then 
			'MsgBox "WaitforAck Timed Out"
			'WaitForAck = False
		End If
		
	End Function
	'**************************************************************************
	' Checks if the first byte(s) in the string is an ack
	Private Function IsAck(ByRef iostring As String) As Boolean
		If InStr(iostring, bcAckString) = 1 Then IsAck = True Else IsAck = False
	End Function
	'**************************************************************************
	' Checks if the first byte(s) in string is a nack
	Private Function IsNAck(ByRef iostring As String) As Boolean
		If InStr(iostring, bcNAckString) = 1 Then IsNAck = True Else IsNAck = False
	End Function
	'**************************************************************************
	' Returns true if there is a carriage return-linefeed pair at the end of the string
	Private Function IsCompleteCode(ByRef iostring As String) As Boolean
		If Len(iostring) > 2 Then 'check to see if string is terminated properly
			If InStr(iostring, bcCodeTerminator) = Len(iostring) - 1 Then
				IsCompleteCode = True
			Else
				IsCompleteCode = False
			End If
		End If
	End Function
	'**************************************************************************
	' Grabs a scan code from the port
	Private Sub ReceiveScanCode(ByRef waitSeconds As Short)
		Dim bGotCode As Boolean
		bGotCode = False
		bcScanTimedOut = False ' set to true by the timer
		bcScanTimer.Interval = waitSeconds * 1000 'convert the wait to the number of milliseconds
		bcScanTimer.Enabled = True
		Do 
			System.Windows.Forms.Application.DoEvents() 'enable events during the loop, in particular, the timer.
			If bcComm.InBufferCount > 0 Then
				bcInputBuff = bcInputBuff + bcComm.Input
			End If
			bGotCode = IsCompleteCode(bcInputBuff)
		Loop Until bcScanTimedOut = True Or bGotCode = True
		bcScanTimer.Enabled = False
	End Sub
	'**************************************************************************
	' These defaults work with a Symbol Technologies LS1220
	Public Sub Use_LS1220_Defaults()
		
		'This section is specific to the Symbol Technologiest reader.
		'Consult their documentation on how to program the symbol reader
		'A confusing element to Symbol usage is the need to calculate the
		'an LRC character, that is transmitted at the end of every command sent to the
		'reader.
		'This LRC character is composed of the exclusive OR of the preceding characters, with
		'exception of the initial STX character. This exclusive or operation is why
		'the XOR operator is needed.
		'The code below shows how that LRC character is calculated.
		
		'Create the Laser on Command
		bcLaserOnCmd = Chr(&H2s) & Chr(&H53s) & Chr(&H54s) & Chr(&H49s) & Chr(&H45s) & Chr(&H3s)
		
		'    bcLaserOnCmd = Chr(&H2) + Chr(&H53) + Chr(&H54) + Chr(&H49) + Chr(&H45) _
		''                   STX         S               T           I       E
		'            + Chr(&H3)
		'               ETX
		Dim length, iter As Short
		Dim X As Byte
		' calculate the xor of the command for the last byte
		length = Len(bcLaserOnCmd)
		X = 0
		For iter = 2 To length
			X = X Xor Asc(Mid(bcLaserOnCmd, iter, 1))
		Next 
		bcLaserOnCmd = bcLaserOnCmd & Chr(X)
		
		'Note that the Laser Off command DOESN'T need the LRC character, and consists
		'only of an excape character.
		bcLaserOffCmd = Chr(&H1Bs)
		
		'The init command puts the bar code reader into a defined state (which
		'happens to be the default state)
		bcInitCmd = Chr(&H2s) & Chr(&H53s) & Chr(&H54s) & Chr(&H49s) & Chr(&H48s) & Chr(&H31s) & Chr(&H30s) & Chr(&H30s) & Chr(&H36s) & Chr(&H31s) & Chr(&H31s) & Chr(&H30s) & Chr(&H31s) & Chr(&H31s) & Chr(&H32s) & Chr(&H30s) & Chr(&H32s) & Chr(&H31s) & Chr(&H33s) & Chr(&H30s) & Chr(&H37s) & Chr(&H31s) & Chr(&H34s) & Chr(&H30s) & Chr(&H30s) & Chr(&H31s) & Chr(&H35s) & Chr(&H30s) & Chr(&H30s) & Chr(&H31s) & Chr(&H36s) & Chr(&H30s) & Chr(&H34s) & Chr(&H3s)
		
		'Documentation for what all this means.
		'Sends out the beginning state state for the bar code reader,
		'bcInitCmd = Chr(&H2) + Chr(&H53) + Chr(&H54) + Chr(&H49) + Chr(&H48) _
		''                 STX       S       T           I               H
		'Tells the bar code to accept down loaded parameters
		
		'                + Chr(&H31) + Chr(&H30) + Chr(&H30) + Chr(&H36) _
		''                   1               0           0           6
		'Sets the baud rate to 9600
		
		'                + Chr(&H31) + Chr(&H31) + Chr(&H30) + Chr(&H31) _
		''                   1               1           0           1
		'Sets the parity to even
		
		'                + Chr(&H31) + Chr(&H32) + Chr(&H30) + Chr(&H32) _
		''                   1               2           0           2
		'Sets the stop bits to 2
		
		'                + Chr(&H31) + Chr(&H33) + Chr(&H30) + Chr(&H37) _
		''                   1               3           0           7
		'Sets the data bits to 7
		
		'                + Chr(&H31) + Chr(&H34) + Chr(&H30) + Chr(&H30) _
		''                   1               4           0           0
		'Sets hardware handshaking to none
		
		'                + Chr(&H31) + Chr(&H35) + Chr(&H30) + Chr(&H30) _
		''                   1               5           0           0
		'Sets software handshaking to none
		
		'                + Chr(&H31) + Chr(&H36) + Chr(&H30) + Chr(&H34) _
		''                   1               6           0           4
		'Sets triggering options to host triggering
		
		'                + Chr(&H3)
		'                   ETX
		'The terminator symbol for a command.
		
		'Now calculate the xor of the command, to be used in the last byte of the command
		length = Len(bcInitCmd)
		X = 0
		For iter = 2 To length
			X = X Xor Asc(Mid(bcInitCmd, iter, 1))
		Next 
		bcInitCmd = bcInitCmd & Chr(X)
		
		'The following code sets up the default properties for a Symbol instance
		'of the BarCode object.
		
		bcComPort = 1
		'The following are the defaults for the Symbol Technologies LS1220
		bcBaudRate = 9600
		bcParity = "E"
		bcStopBits = 2
		bcDataBits = 7
		bcHardwareHandShaking = False
		bcSoftwareHandShaking = False
		'The LS1220 protocol uses an Ack/Nak upon all programming, and this is
		'NOT a settable property for the LS1220 reader.  The value set here is simply
		'a state variable, that is used to control how we look for messages.
		bcAckNakOnProgram = True
		bcAckString = Chr(&H6s)
		bcNAckString = Chr(&H15s)
		bcCodeTerminator = Chr(&HDs) & Chr(&HAs) 'defines a carriage return, line feed
		bcIOTimeout = 2
	End Sub
	
	'**************************************************************************
	' sets the com port with which the host will talk
	'**************************************************************************
	Public Property ComPort() As Short
		Get
			ComPort = bcComPort
		End Get
		Set(ByVal Value As Short)
			bcComPort = Value
		End Set
	End Property
	'**************************************************************************
	' sets the baudrate the host will send and receive at after init
	'**************************************************************************
	Public Property Baud() As Short
		Get
			Baud = bcBaudRate
		End Get
		Set(ByVal Value As Short)
			bcBaudRate = Value
		End Set
	End Property
	'**************************************************************************
	' sets the parity the host will encode and decode after init
	'**************************************************************************
	Public Property parity() As String
		Get
			parity = bcParity
		End Get
		Set(ByVal Value As String)
			bcParity = Value
		End Set
	End Property
	'**************************************************************************
	' sets the number of data bits the host will send and receive after init
	'**************************************************************************
	Public Property NumDataBits() As Short
		Get
			NumDataBits = bcDataBits
		End Get
		Set(ByVal Value As Short)
			bcDataBits = Value
		End Set
	End Property
	'**************************************************************************
	' sets the number of stop bits the host will send and expect after init
	'**************************************************************************
	Public Property NumStopBits() As Short
		Get
			NumStopBits = bcStopBits
		End Get
		Set(ByVal Value As Short)
			bcStopBits = Value
		End Set
	End Property
	'**************************************************************************
	' the host will expect software handshaking. (the host always expects a nack or
	' an ack after sending the init string
	'**************************************************************************
	' gets the state of hardware handshaking
	Public Property HWHandShaking() As Boolean
		Get
			HWHandShaking = bcHardwareHandShaking
		End Get
		Set(ByVal Value As Boolean)
			bcHardwareHandShaking = Value
		End Set
	End Property
	'**************************************************************************
	' sets software handshaking via NAck and Ack
	'**************************************************************************
	Public Property SWHandShaking() As Boolean
		Get
			SWHandShaking = bcSoftwareHandShaking
		End Get
		Set(ByVal Value As Boolean)
			bcSoftwareHandShaking = Value
		End Set
	End Property
	
	'**************************************************************************
	' sets handshaking on sending a command
	'**************************************************************************
	Public Property AckNakOnProgram() As Boolean
		Get
			AckNakOnProgram = bcAckNakOnProgram
		End Get
		Set(ByVal Value As Boolean)
			bcAckNakOnProgram = Value
		End Set
	End Property
	
	'**************************************************************************
	' sets the string that defines the end of a fully read bar code
	'**************************************************************************
	Public Property CodeTerminator() As String
		Get
			CodeTerminator = bcCodeTerminator
		End Get
		Set(ByVal Value As String)
			bcCodeTerminator = Value
		End Set
	End Property
	
	'**************************************************************************
	' sets the timeout for barcode I/O
	'**************************************************************************
	Public Property IOWait() As Short
		Get
			IOWait = bcIOTimeout
		End Get
		Set(ByVal Value As Short)
			bcIOTimeout = Value
		End Set
	End Property
	'**************************************************************************
	' this string will be sent to the barcode reader to initialize it
	'**************************************************************************
	Public Property InitCmd() As String
		Get
			InitCmd = bcInitCmd
		End Get
		Set(ByVal Value As String)
			bcInitCmd = Value
		End Set
	End Property
	'**************************************************************************
	' this string will be sent to the barcode reader to turn it on to scan
	'**************************************************************************
	Public Property LaserOnCmd() As String
		Get
			LaserOnCmd = bcLaserOnCmd
		End Get
		Set(ByVal Value As String)
			bcLaserOnCmd = Value
		End Set
	End Property
	'**************************************************************************
	' this string will be sent to the barcode reader to turn it off after
	' a successfull scan or after a timeout
	'**************************************************************************
	Public Property LaserOffCmd() As String
		Get
			LaserOffCmd = bcLaserOffCmd
		End Get
		Set(ByVal Value As String)
			bcLaserOffCmd = Value
		End Set
	End Property
	'**************************************************************************
	' this is the string the host will expect from the barcode reader after
	' a command is successfully received by the barcode reader if software
	' handshaking is turned on
	'**************************************************************************
	Public Property SuccessAckStr() As String
		Get
			SuccessAckStr = bcAckString
		End Get
		Set(ByVal Value As String)
			bcAckString = Value
		End Set
	End Property
	'**************************************************************************
	' this is the string the host will receive from the barcode reader upon
	' a failure to interpret a command if software handshaking is turned on
	'**************************************************************************
	Public Property FailAckStr() As String
		Get
			FailAckStr = bcNAckString
		End Get
		Set(ByVal Value As String)
			bcNAckString = Value
		End Set
	End Property
	'**************************************************************************
	'**************************************************************************
	Public Property IOTimer() As Object
		Get
			IOTimer = bcIOTimer
		End Get
		Set(ByVal Value As Object)
			bcIOTimer = Value
		End Set
	End Property
	'**************************************************************************
	'**************************************************************************
	Public Property ScanTimer() As Object
		Get
			ScanTimer = bcScanTimer
		End Get
		Set(ByVal Value As Object)
			bcScanTimer = Value
		End Set
	End Property
	'**************************************************************************
	'**************************************************************************
	Public Property CommObject() As Object
		Get
			'Not needed, as we took out the IOMOD underlayer
			CommObject = bcComm
		End Get
		Set(ByVal Value As Object)
			bcComm = Value
			
		End Set
	End Property
	'**************************************************************************
	' called by bccomm control when data is available
	Public Sub CommCallback()
		bcInputBuff = bcInputBuff + bcComm.Input
		If IsCompleteCode(bcInputBuff) Then
			bcLastCode = bcInputBuff
			bcInputBuff = ""
			RaiseEvent CodeScanned(bcLastCode)
		End If
	End Sub
	'**************************************************************************
	' sets a variable read by a loop to decide when to time out
	Public Sub IOTimeoutCallback()
		bcIOTimedOut = True
	End Sub
	'**************************************************************************
	' sets a variable read by a loop to decide when to time out
	Public Sub ScanTimeoutCallback()
		bcScanTimedOut = True
	End Sub
	'**************************************************************************
	Private Sub Class_Initialize_Renamed()
		' used so only one scan happens at a time
		bcScanRunning = False
		'Customize:  Change this to whatever defaults you wish to use.
		Use_LS1220_Defaults()
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	'**************************************************************************
	Private Sub Class_Terminate_Renamed()
		bcScanTimedOut = True
		bcIOTimedOut = True
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class