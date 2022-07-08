Option Strict Off
Option Explicit On
Module modConfiguration
	'*****************************************************************************
	'modConfiguration.bas
	'
	'The file modConfiguration.bas  is used to set configuration variables
	'for the application.  Variables that can be set here include:
	'--the language of the operator interface
	'--the default language of the operator interface
	'--whether diagnostics are printed to the immediate window
	'--whether the test time is output to the report.
	'--whether a run command should be attempted following a successful
	'bar code read
	'--under what conditions printing should occur.
	'
	'See the comments below on how each can be set.
	'
	'Note that the configuration of serial devices (like a strip
	'printer, or a serial bar code reader) is done in their
	'respective forms.
	
	'Select configuration variables are declared in other modules
	'For example, the Localization modules are defined in the localization
	'module, but are defined (set) here.
	'*****************************************************************************
	
	
	Public gbRunAutomaticallyAfterBarCode As Boolean
	Public gbWedgeKeyboardBarCode As Boolean
    Public gbSerialBoxBarCode As Boolean
    Public gfrmSerialBoxBarCode As frmSerialBarCode
    Public gbTxSLSharedIO As Boolean
    Public gfrmTxSLSharedIO As frmTxSLSharedIO
	
    Public gbStripPrinterAvailable As Boolean
    Public gfrmStripPrinter As frmStripPrinter
	Public gbAutoPrintPass As Boolean
	Public gbAutoPrintNonPass As Boolean
	
	Public gbTxSLDebugEvents As Boolean
	Public gbTxSLTestplanTimeEnabled As Boolean
	Public gdtTestplanStartTime As Date
	Public gdtTestplanStopTime As Date
	
	Private LocalSource As String 'A string that would be used to identify the context for error handling
	Private Const ModuleName As String = "modConfiguration"
	'MaxSelLength defines the maximum length of the rich text boxes
	'It is used to keep the focus as the bottom of the control, as the
	'boxes are filled.
	Public Const MaxSelLength As Integer = 1000000
	
	
	'Note that guLanguage and guDefaultLanguage are declared
	'in the localization.bas module
	
	Public Sub Configure()
		LocalSource = ":" & ModuleName & ":Configure"
		On Error GoTo LocalErrorHandler
		
		'Customize:  This whole sub is used to specify configuration variables.  Change
		'as desired
		
		'The gbTxSLDebugEvents variable, when set true, will result in a
		'string containing event information to be output to the immediate window upon each TxSL event.
		'The default value is false
		gbTxSLDebugEvents = True
		
		'The gbTxSLTestplanTimeEnabled variable, when set true, will result in the total time of the testplan
		'being output to the report window. This output is useful for measuring the impact
		'of changes to the testplan or user interface, upon execution time.
		'The default value is false.
		gbTxSLTestplanTimeEnabled = False
		
		'guLanguage is the Language in which the operator interface will be displayed
		'If the program finds a zero length string using the guLangauge variable, it will substitute
		'the string found when using guDefaultLanguage.  This allows for partial implementation of
		'the operator interface in the language specified by guLanguage, and would suggest a full
		'implementation in the language specified by guDefaultLangauge
		modLocalization.guLanguage = modLocalization.TxSLOPUIInterfaceLanguage.English
		modLocalization.guDefaultLanguage = modLocalization.TxSLOPUIInterfaceLanguage.English
		
		'This configuration variable indicates whether a wedge type, Keyboard emulating bar
		'code reader is present.  These readers do not take any configuration, but this configuration variable
		'will be used to control how the application sets focus.  For example, if the gbWedgeKeyBoardBarCode variable
		'is true, then, after successful login, the focus will be set to the bar code field.  If it is
		'false, it will be set to the Load Testplan button.  Setting this value to false does not
		'disable the wedge + keyboard bar codes; setting it to false will make it necessary
		'to tab to the correct fields if not using a barcode reader.
		gbWedgeKeyboardBarCode = True
		
		'Setting this true will cause a successful barcode read to be automatically followed
		'by a testplan run method call.  The default value is False.
		'Note that the bar code could have been read from either a keyboard wedge or
		'a serial bar code, with both putting results into the bar code masked edit
		'box on the main form.
		gbRunAutomaticallyAfterBarCode = False
		
		'Setting this true indicates that there is a serial "box" bar code reader in the system.
		'This means that a bar code reader that can be programmed serially is present in the system
		'Today, this would imply a Symbolic LS 1200A, but other models are possible.
		'The code relies on certain functionality being present.  See references to gbSerialBoxBarCode
		'and look at the form frmSerialBarCode
		'The default value is False.
		gbSerialBoxBarCode = False
		
		'Setting this true indicates that the operator interface is writing to some
		'IO that is under the control of the TxSL process.  These are usually digital lines.
		'Code is written to send messages what "Indicate OK", "IndicateRequestHelp" and "IndicateWarning"
		'This code is implemented on the form frmTxSLSharedIO.  See that code for details.
		'The default value is False.
		gbTxSLSharedIO = False
		
		'Setting these parameters controls if and when this interface will be updated
		'by a TxSL generated timer.  This TxSL timer is used to guarantee that the operator interface
		'is being updated at "safe" times.  The default values are
		'AdviseUpdateEventEnabled = False
		'AdviseUpdateInterval = 500
		'The parameters could be set in code, as in the following two lines.
		'frmMain.TestExecSL1.AdviseUpdateEventEnabled = True
		'frmMain.TestExecSL1.AdviseUpdateInterval = 500
		
		'This sets whether test level events are enabled
		'Remember that test levels events may impose more overhead on system operation
		'than you wish to tolerate.  The default value is False.
		'The parameter could be set in code, as in the following line.
        'frmMain.TestExecSL1.TestEventsEnabled = True
		'However, it is more likely that you will set properties from the property page of the TxSL control.
		'Setting properties via the property page is the recommended approach.
		
		'The timer control called tmrFailure on frmMain.frm controls how long
		'the operator interface will wait on commands like Run, Stop, Abort and Exit
		'before returning control to the operator for further action.  It is there to
		'allow the operator to regain some control, should a call like Run not return
		'Setting it's default value should be done via the tmrFailure property page on frmMain.
		'We recommend a long default value, like 60000 (60 seconds)
		
		'gbStripPrinterAvailable is a global variable that indicates whether a
		'serial printer is present in the system.
		'This variable only indicates that the strip printer is available, it must have been separately
		'installed and configured.  See the form frmStripPrinter.
		'When gbStripPrinterAvailable is used in combination with the
		'gbAutoPrintPass and gbAutoPrintNonPass variables, control of printing can
		'be established.
		'The default value for gbStripPrinterAvailable is false
		gbStripPrinterAvailable = False
		
		'This sets under which conditions a "repair ticket" will be printed.
		'Setting gbAutoPrintPasses = True results in tickets being printed, if the UUT
		'Passed.
		'The default value for gbAutoPrintPasses is False.
		'Setting gbAutoPrintNonPass = True results in tickets being printed if the
		'UUT did NOT pass.  The most typical example of a non-pass is the UUT failing,
		'but other nonpass conditions (like a testplan that was stopped or aborted)
		'will also generate a repair ticket.
		'The default value for gbAutoPrintNonPass is False.
		'The default value for gbAutoPrintPass if False
		
		gbAutoPrintPass = False
		gbAutoPrintNonPass = False
		
		Exit Sub
LocalErrorHandler: 
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
	End Sub
End Module