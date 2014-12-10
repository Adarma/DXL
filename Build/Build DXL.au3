#cs ----------------------------------------------------------------------------

Author:
	Adam Cadamally

Description:
	Run DXL code and pipe output to STDOUT.

Licence:

	Copyright (c) 2013 Adam Cadamally

	Permission is hereby granted, free of charge, to any person obtaining a copy
	of this software and associated documentation files (the "Software"), to deal
	in the Software without restriction, including without limitation the rights
	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	copies of the Software, and to permit persons to whom the Software is
	furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in
	all copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
	THE SOFTWARE.

#ce ----------------------------------------------------------------------------

; MAIN PROGRAM

Opt("WinSearchChildren", 1)   ;0=no, 1=search children also
Opt("MustDeclareVars", 1)     ;0=no, 1=yes

#include <File.au3>			; For: _PathFull(), _PathSplit

Local $oErrorHandler = ObjEvent("AutoIt.Error","ComErrorHandler")    ; Initialize a COM error handler
; This is my custom defined error handler
Func ComErrorHandler($oMyError)
	Return
	WriteUnicode("We intercepted a COM Error !"    & @LF  & @LF & _
             "err.description is: " & @TAB & $oMyError.description  & @LF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @LF & _
             "err.number is: "       & @TAB & hex($oMyError.number,8)  & @LF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @LF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @LF & _
             "err.source is: "       & @TAB & $oMyError.source       & @LF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @LF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext    & @LF & _
            )
Endfunc

; Debugging Arguments
Local $ScriptFile = _PathFull("../Test/Debug Test.dxl", @ScriptDir & "\")
Local $DxlMode = "Run"

; Real Arguments
If @Compiled Then
	; Check if Arguments were passed correctly
	If $CmdLine[0] < 2 Then
		WriteUnicode("Wrong Parameters" & @LF)
		Exit
	EndIf

	; Get Full File Path from Arguments
	$ScriptFile = $CmdLine[1]
	$DxlMode = $CmdLine[2]
Else
	WriteUnicode("Check Unicode: öÊãt ♠♣♥♦ ♪♫ ♀♂" & @LF)
EndIf

Local $LogFile = GetDoorsLogFile()

Local $szDrive, $szDir, $szFName, $szExt
_PathSplit($LogFile, $szDrive, $szDir, $szFName, $szExt)
Local $LogFolder = $szDrive & $szDir
Local $LastLogFile = $LogFolder & $szFName & " Last.log"

If $DxlMode = "ShowErrors" Then

	; Pipe the last errors from code invoked via Sublime Text
	Local $LastLogFileHandle = FileOpen($LastLogFile, $FO_READ)
	Local $LastLogFileText = FileRead($LastLogFileHandle)
	If StringLen($LastLogFileText) > 0 Then
		WriteUnicode(@LF & "Last Error Log:" & @LF)
		WriteUnicode($LastLogFileText & @LF)
	EndIf
	FileClose($LastLogFileHandle)

	; Pipe the full error log to get errors in code not invoked via Sublime Text
	Local $LogFileHandle = FileOpen($LogFile, $FO_READ)
	Local $LogFileText = FileRead($LogFileHandle)
	If StringLen($LogFileText) > 0 Then
		WriteUnicode(@LF & "Full Error Log:" & @LF)
		WriteUnicode($LogFileText & @LF)
	EndIf
	FileClose($LogFileHandle)

Else

	; We require 'DOORSLOGFILE' to be defined in order to pipe Errors and Warnings
	If Not $LogFile Then
		WriteUnicode("'DOORSLOGFILE' is not defined for Warning and Error logging" & @LF & @LF)
		WriteUnicode("Unfortunatly it is not possible to toggle warning and error logging during a DOORS session." & @LF)
		WriteUnicode("Changing the setting requires DOORS to be restarted." & @LF)
		WriteUnicode("To send the warnings and errors to Sublime Text, logging is required." & @LF & @LF)
		WriteUnicode("Logging may be enabled by setting a Value in the Registry Key:" & @LF)
		WriteUnicode("HKCU\Software\Telelogic\DOORS\#.#\Config" & @LF)
		WriteUnicode("The ValueName should be 'LOGFILE' of Type 'REG_SZ' with the Data set to a full file path." & @LF)
		WriteUnicode("DOORS must be restarted for it to take effect." & @LF & @LF)
		WriteUnicode("Unfortunatly this will redirect all warnings and errors to the file, not just dxl started via Sublime Text." & @LF)
		WriteUnicode("You may display the log with F6, Build: DXL: Errors." & @LF & @LF)
		WriteUnicode("Example Commandline to set the registry value:" & @LF)
		WriteUnicode('reg add "HKCU\Software\Telelogic\DOORS\9.5\Config" /f /v "LOGFILE" /t REG_SZ /d "C:\DxlErrors.log"' & @LF & @LF)
		WriteUnicode("Example Commandline to delete the registry value:" & @LF)
		WriteUnicode('reg delete "HKCU\Software\Telelogic\DOORS\9.5\Config" /f /v "LOGFILE"' & @LF)
		Exit
	EndIf

	; TODO: Use getenv("DOORSLOGFILE") via COM or DXL Interaction in case multiversions of DOORS
;~ 	$ObjDoors.runStr('oleSetResult(getenv("DOORSLOGFILE"))')
;~ 	$LogFile = $ObjDoors.Result

	; Use the last active DXL Interaction Window if one is open, to allow database targeting when 2+ DOORS instances are running

	; Remember if the DXL Interaction window is already open
	Local $DxlInteractionWindow = "DXL Interaction - DOORS"
	Local $DxlOpen = IsVisible($DxlInteractionWindow)

	; The 'current' Module when executing via COM
	Local $ModuleFullName = ""
	Local $ModuleType = ""
	Local $ModuleHwnd = 0
	Local $IsBaseline = "false"
	Local $Major = 0
	Local $Minor = 0
	Local $Suffix = ""

	; We require DOORS to be running
	Local $DoorsRunning = $DxlOpen
	If Not $DxlOpen Then
		; Get Active DOORS Window details (ByRef)
		$DoorsRunning = GetActiveDoorsWindowDetails($ModuleFullName, $ModuleType, $ModuleHwnd, $IsBaseline, $Major, $Minor, $Suffix) ; MAKE SURE IT IS CHILD OF COM DATABASE
	EndIf

	If Not $DoorsRunning Then
		WriteUnicode("DOORS is not running!" & @LF)
		Exit
	EndIf

	; Get the running Doors COM Object using it's class name [ObjGet() fails]
	Local $ObjDoors = ObjCreate("DOORS.application")

	If @error Then
		WriteUnicode("Failed to connect to DOORS." & @LF & "Error Code: " & Hex(@error, 8) & @LF)
		Exit
	EndIf

	If Not IsObj($ObjDoors) Then
		WriteUnicode("Failed to connect to DOORS.")
		Exit
	EndIf

	If $DxlOpen Then
		WriteUnicode("Code Execution: 'DXL Interaction' window" & @LF)
	Else
		WriteUnicode("Code Execution: COM" & @LF)
	EndIf

	; File to 'print' to
	Local $OutFile = $LogFolder & "DxlPrint.log"

	; Find the Relitive Base Paths for DOORS
	ClearFile($OutFile)
	Local $BasePathsString = ""
	ShellExecute("Run DXL.exe", '"" "' & $ScriptFile & '" "' & $ModuleFullName & '" ' & $IsBaseline & ' ' & $Major & ' ' & $Minor & ' "' & $Suffix & '" "' & $OutFile & '" ' & "RelitiveBasePaths", @ScriptDir)
	Sleep(100)
	While _FileInUse($OutFile, 1) = 1
		Sleep(50)
	WEnd
	Local $OutFileHandle = FileOpen($OutFile, $FO_READ)
	If $OutFileHandle = -1 Then
		WriteUnicode("Relitive Paths: " & "Failed to get Base Paths" & @LF)
	Else
		Local $BasePathsString = StringReplace(FileRead($OutFileHandle), "/", "\")
		WriteUnicode("Relitive Paths: " & $BasePathsString & @LF)
	EndIf
	FileClose($OutFileHandle)
	Local $BasePathsArray = StringSplit($BasePathsString, ';', 1)

	; Extract //<Requires> directives
	Local $Requires = LintRequires($ObjDoors, $ScriptFile)
	; ConsoleWrite("Parse Requires: " & $Requires & @CRLF)

	; Make the DXL include statement
	Local $IncludeString = $Requires & "#include <" & $ScriptFile & ">;" & @LF

	; Lint the Code and Time It
	Local $ParseTime = TimerInit()
	Local $DXLOutputText = LintCode($ObjDoors, $IncludeString)
	$ParseTime = TimerDiff($ParseTime)
	$ParseTime = StringLeft($ParseTime, StringInStr($ParseTime, ".") -1)
	WriteUnicode("Parse Duration: " & $ParseTime & " milliseconds" & @LF)

	If $DXLOutputText <> "" Then
		WriteUnicode("Parse Errors:" & @LF)
		Local $OutputLines = stringsplit($DXLOutputText, @CRLF, 0)
		Local $LineIndex
		For $LineIndex = 1 To $OutputLines[0]
			If $OutputLines[$LineIndex] = "" Then
				; Don't pipe last empty line
				If $LineIndex < $OutputLines[0] Then
					WriteUnicode($OutputLines[$LineIndex] & @LF)
				EndIf
			Else
				WriteUnicode(AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex]) & @LF)
			EndIf
		Next
	Else
		; Find the Active Sublime Text Window so it can be reactivated after a dxl error
		Local $ActiveWindow = GetActiveSublimeTextWindow()

		; Initialize
		Local $TraceAllLines = (StringRight($DxlMode, 7) = "Verbose")

		Local $LogFileHandle = FileOpen($LogFile, $FO_READ)
		Local $OldLogFileText = FileRead($LogFileHandle)
		FileClose($LogFileHandle)

		; Run the DXL - Invoked by a separate process so this one can pipe the output back
		ClearFile($OutFile)
		ShellExecute("Run DXL.exe", '"' & $Requires & '" "' & $ScriptFile & '" "' & $ModuleFullName & '" ' & $IsBaseline & ' ' & $Major & ' ' & $Minor & ' "' & $Suffix & '" "' & $OutFile & '" ' & $DxlMode, @ScriptDir)
		Sleep($ParseTime + 500)

		; Error Window Titles
		Local $CppErrorWindow = "Microsoft Visual C++ Runtime Library"
		Local $DiagnosticLogWindow = "Diagnostic Log - DOORS"
		Local $RuntimeErrorWindow = "DOORS report"

		; Set File to pipe the output from
		Local $PipeFilePath = $OutFile
		If StringLeft($DxlMode, 16) = "TraceAllocations" Then
			$PipeFilePath = $LogFolder & "DxlAllocations.log"
			WriteUnicode("Allocations:" & @LF)
		EndIf
		If StringLeft($DxlMode, 14) = "TraceExecution" Then
			$PipeFilePath = $LogFolder & "DxlCallTrace.log"
			WriteUnicode("Execution:" & @LF)
		EndIf
		If StringLeft($DxlMode, 11) = "TraceDelays" Then
			$PipeFilePath = $LogFolder & "DxlCallTrace.log"
			WriteUnicode("Delays:" & @LF)
		EndIf
		If StringLeft($DxlMode, 14) = "TraceVariables" Then
			$PipeFilePath = $LogFolder & "DxlVariables.log"
			WriteUnicode("Variables:" & @LF)
		EndIf

		; While running, pipe the output
		Local $CppError = False
		Local $DiagnosticLog = False
		Local $RuntimeError = False
		Local $OldOutputText = ""
		Local $StartTime = TimerInit()
		While _FileInUse($OutFile, 1) = 1
			Sleep(100)

			If Not FileExists($OutFile) and TimerDiff($StartTime) > 1000 Then
				ExitLoop
			EndIf

			; Possible C++ Error Window
			$CppError = WinExists($CppErrorWindow)

			; Possible Runtime Error Window
			$RuntimeError = WinExists($RuntimeErrorWindow)

			; Possible Diagnostic Log Window
			$DiagnosticLog = WinExists($DiagnosticLogWindow) And BitAnd(WinGetState($DiagnosticLogWindow), 2)


			If $CppError Then
				; Close C++ Error message box
				ControlClick($CppErrorWindow, "", "[CLASS:Button; INSTANCE:1]")
			EndIf

			If $RuntimeError Then
				; Close Runtime Error message box
				ControlClick($RuntimeErrorWindow, "", "[CLASS:Button; INSTANCE:1]")
			EndIf

			If $DiagnosticLog Then
				; Close Runtime Error message box
				ControlClick($DiagnosticLogWindow, "", "[CLASS:Button; INSTANCE:1]")
			EndIf

			; Pipe the new output
			Local $OutFileHandle = FileOpen($PipeFilePath, $FO_READ)
			If $OutFileHandle = -1 Then
				WriteUnicode("Unable to get DOORS Output" & @LF)
				ExitLoop
			Else
				Local $OutputText = FileRead($OutFileHandle)
				Local $NewText = StringTrimLeft($OutputText, StringLen($OldOutputText))
				If $NewText <> "" Then
					Local $OutputLines = stringsplit($NewText, @CRLF, 1)
					Local $LineIndex
					For $LineIndex = 1 To $OutputLines[0]
						If $OutputLines[$LineIndex] = "" Then
							; Don't pipe empty line from Trace
							If StringLeft($DxlMode, 5) <> "Trace" and $LineIndex < $OutputLines[0] Then
								WriteUnicode($OutputLines[$LineIndex] & @LF)
							EndIf
						Else
							If StringLeft($DxlMode, 5) = "Trace" Then
								; Don't pipe <Line:...> from Trace
								If StringLeft($OutputLines[$LineIndex], 1) = "<" Then
									If StringLeft($OutputLines[$LineIndex], 6) <> "<Line:" Then
										If $TraceAllLines Or StringLeft($OutputLines[$LineIndex], StringLen($ScriptFile) + 2) = "<" & $ScriptFile & ":" Then
											WriteUnicode(AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex]) & @LF)
										EndIf
									EndIf
								EndIf
							Else
								WriteUnicode(AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex]) & @LF)
							EndIf
						EndIf
					Next
				EndIf
			EndIf
			$OldOutputText = $OutputText
			FileClose($OutFileHandle)

		WEnd

		; Make sure Debugging features are turned off when code interupted (crash, halt, show etc)
		$ObjDoors.runStr('setDebugging_(false);stopDXLTracing_();')

		; Close DXL Interaction Window if it got opened
		If Not $DxlOpen Then
			Local $DxlInteractionWindow = "DXL Interaction - DOORS"
			If WinExists($DxlInteractionWindow) Then
				If IsVisible($DxlInteractionWindow) Then
					WinClose($DxlInteractionWindow)
				EndIf
			EndIf
		EndIf

		; Pipe the remaining output
		Local $OutFileHandle = FileOpen($PipeFilePath, $FO_READ)
		If $OutFileHandle = -1 Then
			WriteUnicode("Unable to get DOORS Output" & @LF)
		Else
			Local $OutputText = FileRead($OutFileHandle)
			Local $NewText = StringTrimLeft($OutputText, StringLen($OldOutputText))
			If $NewText <> "" Then
				Local $OutputLines = stringsplit($NewText, @CRLF, 1)
				Local $LineIndex
				For $LineIndex = 1 To $OutputLines[0]
					If $OutputLines[$LineIndex] = "" Then
						; Don't pipe empty line from Trace
						If StringLeft($DxlMode, 5) <> "Trace" Then
							WriteUnicode($OutputLines[$LineIndex] & @LF)
						EndIf
					Else
						If StringLeft($DxlMode, 5) = "Trace" Then
							; Don't pipe <Line:...> from Trace
							If StringLeft($OutputLines[$LineIndex], 1) = "<" Then
								If StringLeft($OutputLines[$LineIndex], 6) <> "<Line:" Then
									If $TraceAllLines Or StringLeft($OutputLines[$LineIndex], StringLen($ScriptFile) + 2) = "<" & $ScriptFile & ":" Then
										WriteUnicode(AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex]) & @LF)
									EndIf
								EndIf
							EndIf
						Else
							WriteUnicode(AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex]) & @LF)
						EndIf
					EndIf
				Next
			EndIf
			$OldOutputText = $OutputText
		EndIf
		FileClose($OutFileHandle)

		; Report Closed Error Popups
		If $RuntimeError Or WinExists($RuntimeErrorWindow) Then
			ControlClick($RuntimeErrorWindow, "", "[CLASS:Button; INSTANCE:1]")
			WriteUnicode(@LF)
			WriteUnicode("++++++++++++++++++" & @LF)
			WriteUnicode("+ Runtime Error! +" & @LF)
			WriteUnicode("++++++++++++++++++" & @LF)
		EndIf
		If $DiagnosticLog Or (WinExists($DiagnosticLogWindow) And BitAnd(WinGetState($DiagnosticLogWindow), 2)) Then
			ControlClick($DiagnosticLogWindow, "", "[CLASS:Button; INSTANCE:1]")
			WriteUnicode(@LF)
			WriteUnicode("******************" & @LF)
			WriteUnicode("* Diagnostic Log *" & @LF)
			WriteUnicode("******************" & @LF)
		EndIf
		If $CppError Or WinExists($CppErrorWindow) Then
			ControlClick($CppErrorWindow, "", "[CLASS:Button; INSTANCE:1]")
			WriteUnicode(@LF)
			WriteUnicode("##################" & @LF)
			WriteUnicode("# MS V C++ Error #" & @LF)
			WriteUnicode("##################" & @LF)
		EndIf

		; Pipe Errors and Warnings
		Local $LogFileHandle = FileOpen($LogFile, $FO_READ)
		If $LogFileHandle = -1 Then
			WriteUnicode("Unable to get DOORS Errors" & @LF)
		Else
			Local $LogFileText = FileRead($LogFileHandle)
			Local $NewText = StringTrimLeft($LogFileText, StringLen($OldLogFileText))
			If $NewText <> "" Then
				; Save Error log containing just the last errors
				Local $NewLogFileHandle = FileOpen($LastLogFile, $FO_OVERWRITE)
				WriteUnicode(@LF & "Error Log:" & @LF)
				Local $OutputLines = stringsplit($NewText, @CRLF, 1)
				Local $LineIndex
				For $LineIndex = 1 To $OutputLines[0]
					Local $AbsoluteLine = AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex])
					WriteUnicode($AbsoluteLine & @LF)
					FileWrite($NewLogFileHandle, $AbsoluteLine)
				Next
				FileClose($NewLogFileHandle)
			EndIf
		EndIf
		FileClose($LogFileHandle)

		; Reactivate selected module because errors will activate explorer
		; The code would then be run in DOORS Explorer the next time
		If StringLen($LogFileText) > StringLen($OldLogFileText) Then
			If $ModuleHwnd And $ActiveWindow Then
				WinSetOnTop($ActiveWindow, "", 1)
				WinActivate($ModuleHwnd)
				WinSetOnTop($ActiveWindow, "", 0)	; Better to restore previous state here
				WinActivate($ActiveWindow)
			EndIf
		EndIf

	EndIf

	$ObjDoors = 0
	WriteUnicode(@LF)

EndIf


; ******************************************************************************************************************* ;


Func LintRequires($ObjDoors, $sFilePath)

	Local $Requires = ''

	; Open the file for reading and store the handle to a variable.
	Local $FileHandle = FileOpen($sFilePath, 0)
	If $FileHandle <> -1 Then

		Local $LineNo = 0
		While 1
			Local $Require = GetNextRequire($FileHandle, $LineNo)
			If $Require == '' Then ExitLoop

			Local $DXLOutputText = LintCode($ObjDoors, $Requires & $Require)
			If $DXLOutputText == '' Then
				$Requires &= $Require & ";"
				ConsoleWrite("Insert Require: " & $Require & @CRLF)
			Else

			   ; Check if file was found
			   Local $Match = StringRegExp($DXLOutputText, "^-E- DXL: <Line:[0-9]+> (.*)" , 1)

			   If @Error == 0 Then
				  ConsoleWrite("-W- DXL: <" & $sFilePath & ":" & $LineNo & "> " & $Match[0] & @CRLF)
			   Else
				  Local $IncludeFile = StringMid($Require, 11, StringLen($Require) - 11)
				  ConsoleWrite("-W- DXL: <" & $sFilePath & ":" & $LineNo & "> could not run include file (" & $IncludeFile & ") (Syntax errors in file)" & @CRLF)
			   EndIf
			EndIf
		Wend

		; Close the handle returned by FileOpen.
		FileClose($FileHandle)

	EndIf

	Return $Requires

EndFunc

Func GetNextRequire($FileHandle, ByRef $LineNo)

    Local $Regexp = '^//<([^>]+)>\s*(.*)$'
	Local $Require = ''

	While 1
		; Read the fist line of the file using the handle returned by FileOpen.
		Local $FileLine = FileReadLine($FileHandle)

		If @error = -1 Then ExitLoop
		$LineNo += 1

		Local $MatchArray = StringRegExp($FileLine, $Regexp, 1)
		If @error Then ExitLoop

		 If $MatchArray[0] == "Requires" Then
		   Local $IncludeMatchArray = StringRegExp($MatchArray[1], '^(#include ["<][^">]+[">])\s*(.*)$', 1)
			If @error Then
			   ConsoleWrite("-W- DXL: <Line:" & $LineNo & "> Invalid '//<Requires>' syntax: Expected '#include ' (" & $MatchArray[1] & ")" & @CRLF)
			Else
			   $Require = $IncludeMatchArray[0]
			   ExitLoop
			EndIf
		 EndIf
	Wend

	Return $Require

EndFunc

Func LintCode($ObjDoors, $Include)

	Local $EscapedInclude = StringReplace($Include, "\", "\\")
	$EscapedInclude = StringReplace($EscapedInclude, '"', '\"')
	Local $TestCode = 'oleSetResult(checkDXL("' & $EscapedInclude & '"))'

	; Test the code
	$ObjDoors.Result = ""
	$ObjDoors.runStr($TestCode)

	Local $DXLOutputText = $ObjDoors.Result
	Return $DXLOutputText

EndFunc

Func GetDoorsLogFile()
	Local $baseKey = "HKEY_CURRENT_USER\SOFTWARE\Telelogic\DOORS"
	Local $index = 1
	While True
		Local $doorsVersion = RegEnumKey($baseKey, $index)
		If @error = 0 Then
			Local $LogFilePath = RegRead($baseKey  & "\"& $doorsVersion & "\Config", "LOGFILE")
			If @error = 0 Then
				Return $LogFilePath
			EndIf
			$index += 1
		Else
			ExitLoop
		EndIf
	WEnd
	Return
EndFunc

Func GetActiveSublimeTextWindow()
	; Find the Sublime Text Window
	Local $SublimeWindows = WinList("[CLASS:PX_WINDOW_CLASS]")
	Local $ActiveSublimeWindow = 0
	; Loop Sublime Text Windows in the order they were last activated
	For $i = 1 To $SublimeWindows[0][0]
		; Only visble windows that have a title
		If $SublimeWindows[$i][0] <> "" And IsVisible($SublimeWindows[$i][1]) Then
			$ActiveSublimeWindow = $SublimeWindows[$i][1]
			ExitLoop
		EndIf
	Next
	Return $ActiveSublimeWindow
EndFunc

Func GetActiveDoorsWindowDetails(ByRef $ModuleFullName, Byref $ModuleType, ByRef $ModuleHwnd, ByRef $IsBaseline, ByRef $Major, ByRef $Minor, ByRef $Suffix)
	; Find Last Active DOORS Window - Explorer or Module
	Local $DoorsRunning = false
	$ModuleFullName = ""
	$ModuleType = ""
	$ModuleHwnd = 0
	; Loop DOORS Windows in the order they were last activated
	Local $DoorsWindows = WinList("[CLASS:DOORSWindow]")
	For $i = 1 To $DoorsWindows[0][0]
		; Only visble windows that have a title
		If $DoorsWindows[$i][0] <> "" And IsVisible($DoorsWindows[$i][1]) Then

			; Detect Doors Explorer from the title
			Local $DoorsExplorer = StringRegExp($DoorsWindows[$i][0], "^[^'].+: (/|\.{3}).* - DOORS$", 0)

			If ($DoorsExplorer) Then
				$DoorsRunning = True
				ExitLoop
			EndIf

			; Detect Open Formal Modules from their titles
			Local $DoorsModule = StringRegExp($DoorsWindows[$i][0], "^'(.+)' (current|baseline) ([0-9]+)\.([0-9]+)( \([^)]+\))? in (.+) \((Formal|Link) module\) - DOORS$", 1)

			If (UBound($DoorsModule) == 7) Then
				$DoorsRunning = True
				Local $ModuleName = $DoorsModule[0]
				Local $Baseline = $DoorsModule[1]
				If ($Baseline == "baseline") Then
				   $IsBaseline = "true"
				Else
				   $IsBaseline = "false"
				EndIf
				$Major = $DoorsModule[2]
				$Minor = $DoorsModule[3]
				$Suffix = $DoorsModule[4]
				Local $FolderName = $DoorsModule[5]
				$ModuleType = $DoorsModule[6]
				$ModuleFullName = $FolderName & "/" & $ModuleName
				$ModuleHwnd = $DoorsWindows[$i][1]
				ExitLoop
			EndIf
		EndIf
	Next
	Return $DoorsRunning
EndFunc

Func ClearFile($sFilename)
	Local $FileHandle = FileOpen($sFilename, $FO_OVERWRITE)
	If $FileHandle <> -1 Then
		FileClose($FileHandle)
	EndIf
EndFunc

Func IsVisible($WindowHandle)
	Return BitAND(WinGetState($WindowHandle), 2)
EndFunc

Func AbsolutePath($BasePathsArray, $RelativePath)
	For $PathIndex = 1 To $BasePathsArray[0]
		Local $AbsolutePath = _PathFull($RelativePath, $BasePathsArray[$PathIndex])
		If FileExists($AbsolutePath) Then
			Return $AbsolutePath
		EndIf
	Next
	Return $RelativePath
EndFunc

Func AbsoluteLine($BasePathsArray, $Line)
	Local $LineRegExp = "^((?:-?R?-[EWF]- DXL: |\s)?<(?!Line:))(.*)(:(?:[0-9]+)> ?(?:.*))"
	Local $Matches = StringRegExp($Line, $LineRegExp, 1, 1)
	If @error = 0 Then
		Local $AbsolutePath = AbsolutePath($BasePathsArray, $Matches[1])
		Return $Matches[0] &  $AbsolutePath & $Matches[2]
	EndIf
	Return $Line
EndFunc

Func Unicode2Ansi($sString = "")
    ; Convert UTF8 to ANSI
    Local Const $SF_ANSI = 1
    Local Const $SF_UTF8 = 4
    Return BinaryToString(StringToBinary($sString, $SF_UTF8), $SF_ANSI)
EndFunc

Func WriteUnicode($sString = "")
    ConsoleWrite(Unicode2Ansi($sString))
EndFunc

;===============================================================================
;
; Function Name:    _FileInUse()
; Description:      Checks if file is in use
; Syntax.........: _FileInUse($sFilename, $iAccess = 1)
; Parameter(s):     $sFilename = File name
; Parameter(s):     $iAccess = 0 = GENERIC_READ - other apps can have file open in readonly mode
;                   $iAccess = 1 = GENERIC_READ|GENERIC_WRITE - exclusive access to file,
;                   fails if file open in readonly mode by app
; Return Value(s):  1 - file in use (@error contains system error code)
;                   0 - file not in use
;                   -1 dllcall error (@error contains dllcall error code)
; Author:           Siao
; Modified          rover - added some additional error handling, access mode
; Remarks           _WinAPI_CreateFile() WinAPI.au3
;===============================================================================
Func _FileInUse($sFilename, $iAccess = 0)
    Local $aRet, $hFile, $iError, $iDA
    Local Const $GENERIC_WRITE = 0x40000000
    Local Const $GENERIC_READ = 0x80000000
    Local Const $FILE_ATTRIBUTE_NORMAL = 0x80
    Local Const $OPEN_EXISTING = 3
    $iDA = $GENERIC_READ
    If BitAND($iAccess, 1) <> 0 Then $iDA = BitOR($GENERIC_READ, $GENERIC_WRITE)
    $aRet = DllCall("Kernel32.dll", "hwnd", "CreateFile", _
                                    "str", $sFilename, _ ;lpFileName
                                    "dword", $iDA, _ ;dwDesiredAccess
                                    "dword", 0x00000000, _ ;dwShareMode = DO NOT SHARE
                                    "dword", 0x00000000, _ ;lpSecurityAttributes = NULL
                                    "dword", $OPEN_EXISTING, _ ;dwCreationDisposition = OPEN_EXISTING
                                    "dword", $FILE_ATTRIBUTE_NORMAL, _ ;dwFlagsAndAttributes = FILE_ATTRIBUTE_NORMAL
                                    "hwnd", 0) ;hTemplateFile = NULL
    $iError = @error
    If @error Or IsArray($aRet) = 0 Then Return SetError($iError, 0, -1)
    $hFile = $aRet[0]
    If $hFile = -1 Then ;INVALID_HANDLE_VALUE = -1
        $aRet = DllCall("Kernel32.dll", "int", "GetLastError")
        ;ERROR_SHARING_VIOLATION = 32 0x20
        ;The process cannot access the file because it is being used by another process.
        If @error Or IsArray($aRet) = 0 Then Return SetError($iError, 0, 1)
        Return SetError($aRet[0], 0, 1)
    Else
        ;close file handle
        DllCall("Kernel32.dll", "int", "CloseHandle", "hwnd", $hFile)
        Return SetError(@error, 0, 0)
    EndIf
EndFunc
