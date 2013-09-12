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

#include <File.au3>			; For: _TempFile()

Local $oErrorHandler = ObjEvent("AutoIt.Error","ComErrorHandler")    ; Initialize a COM error handler
; This is my custom defined error handler
Func ComErrorHandler($oMyError)
	Return
	ConsoleWrite("We intercepted a COM Error !"    & @LF  & @LF & _
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
		ConsoleWrite("Wrong Parameters" & @LF)
		Exit
	EndIf

	; Get Full File Path from Arguments
	$ScriptFile = $CmdLine[1]
	$DxlMode = $CmdLine[2]
EndIf
	
If $DxlMode = "ShowErrors" Then
	
	; We don't require DOORS to be running
	Local $LogFile = GetDoorsLogFile()
	Local $LastLogFile = StringTrimRight($LogFile, 4) & " Last.log"
	
	; Pipe the last errors from code invoked via Sublime Text
	Local $LastLogFileHandle = FileOpen($LastLogFile, 0)
	Local $LastLogFileText = FileRead($LastLogFileHandle)
	If StringLen($LastLogFileText) > 0 Then
		ConsoleWrite(@LF & "Last Error Log:" & @LF)
		ConsoleWrite($LastLogFileText & @LF)
	EndIf
	FileClose($LastLogFileHandle)
	
	; Pipe the full error log to get errors in code not invoked via Sublime Text
	Local $LogFileHandle = FileOpen($LogFile, 0)
	Local $LogFileText = FileRead($LogFileHandle)
	If StringLen($LogFileText) > 0 Then
		ConsoleWrite(@LF & "Full Error Log:" & @LF)
		ConsoleWrite($LogFileText & @LF)
	EndIf
	FileClose($LogFileHandle)
	
Else
	
	; We require 'DOORSLOGFILE' to be defined in order to pipe Errors and Warnings
	Local $LogFile = GetDoorsLogFile()
	If Not $LogFile Then
		ConsoleWrite("'DOORSLOGFILE' is not defined for Warning and Error logging" & @LF)
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
	
	; We require DOORS to be running
	Local $DoorsRunning = $DxlOpen
	If Not $DxlOpen Then
		; Get Active DOORS Window details (ByRef)
		$DoorsRunning = GetActiveDoorsWindowDetails($ModuleFullName, $ModuleType, $ModuleHwnd) ; MAKE SURE IT IS CHILD OF COM DATABASE
	EndIf
	
	If Not $DoorsRunning Then
		ConsoleWrite("DOORS is not running!" & @LF)
		Exit
	EndIf
	
	; Get the running Doors COM Object using it's class name [ObjGet() fails]
	Local $ObjDoors = ObjCreate("DOORS.application")

	If @error Then
		ConsoleWrite("Failed to connect to DOORS." & @LF & "Error Code: " & Hex(@error, 8) & @LF)
		Exit
	EndIf

	If Not IsObj($ObjDoors) Then
		ConsoleWrite("Failed to connect to DOORS.")
		Exit
	EndIf
	
	If $DxlOpen Then
		ConsoleWrite("Executing Code via: 'DXL Interaction' window" & @LF)
	Else
		ConsoleWrite("Executing Code via: COM" & @LF)
	EndIf
	
	; Make the DXL include statement
	Local $IncludeString = "#include <" & $ScriptFile & ">;" & @LF
	
	; Make the checkDXL string
	Local $EscapedInclude = StringReplace($IncludeString, "\", "\\")
	Local $TestCode = 'oleSetResult(checkDXL("' & $EscapedInclude & '"))'
	
	; Test the code
	Local $ParseTime = TimerInit()
	$ObjDoors.Result = ""
	$ObjDoors.runStr($TestCode)
	$ParseTime = TimerDiff($ParseTime)
	$ParseTime = StringLeft($ParseTime, StringInStr($ParseTime, ".") -1)
	ConsoleWrite("DXL Code Parsed in: " & $ParseTime & " milliseconds" & @LF & @LF)

	Local $DXLOutputText = $ObjDoors.Result
	If $DXLOutputText <> "" Then
		ConsoleWrite("Parse Errors:" & @LF & $DXLOutputText & @LF)
	Else
	
		; Find the Active Sublime Text Window so it can be reactivated after a dxl error
		Local $ActiveWindow = GetActiveSublimeTextWindow()
		
		; File to 'print' to
		Local $OutFile = _TempFile()
		
		; Find the Relitive Base Paths for DOORS
		Local $BasePathsString = ""
		ShellExecute("Run DXL.exe", '"' & $ScriptFile & '" "' & $ModuleFullName & '" "' & $OutFile & '" ' & "RelitiveBasePaths", @ScriptDir)
		Sleep(100)
		While _FileInUse($OutFile, 1) = 1
			Sleep(50)
		WEnd
		Local $OutFileHandle = FileOpen($OutFile, 0)
		If $OutFileHandle = -1 Then
			ConsoleWrite("Relitive Paths: " & "Failed to get Base Paths" & @LF)
		Else
			Local $BasePathsString = StringReplace(FileRead($OutFileHandle), "/", "\")
			ConsoleWrite("Relitive Paths: " & $BasePathsString & @LF)
		EndIf
		FileClose($OutFileHandle)
		Local $BasePathsArray = StringSplit($BasePathsString, ';', 1)
		
		; Initialize
		Local $TraceAllLines = (StringRight($DxlMode, 7) = "Verbose")
		
		Local $LastLogFile = StringTrimRight($LogFile, 4) & " Last.log"

		Local $LogFileHandle = FileOpen($LogFile, 0)
		Local $OldLogFileText = FileRead($LogFileHandle)
		FileClose($LogFileHandle)

		; Run the DXL - Invoked by a separate process so this one can pipe the output back
		ShellExecute("Run DXL.exe", '"' & $ScriptFile & '" "' & $ModuleFullName & '" "' & $OutFile & '" ' & $DxlMode, @ScriptDir)
		Sleep($ParseTime + 500)

		; Error Window Titles
		Local $CppErrorWindow = "Microsoft Visual C++ Runtime Library"
		Local $DiagnosticLogWindow = "Diagnostic Log - DOORS"
		Local $RuntimeErrorWindow = "DOORS report"

		; Set File to pipe the output from
		Local $PipeFilePath = $OutFile
		If StringLeft($DxlMode, 16) = "TraceAllocations" Then
			$PipeFilePath = "C:\\DxlAllocations.log"
			ConsoleWrite("Allocations:" & @LF)
		EndIf
		If StringLeft($DxlMode, 14) = "TraceExecution" Then
			$PipeFilePath = "C:\\DxlCallTrace.log"
			ConsoleWrite("Execution:" & @LF)
		EndIf
		If StringLeft($DxlMode, 11) = "TraceDelays" Then
			$PipeFilePath = "C:\\DxlCallTrace.log"
			ConsoleWrite("Delays:" & @LF)
		EndIf
		If StringLeft($DxlMode, 14) = "TraceVariables" Then
			$PipeFilePath = "C:\\DxlVariables.log"
			ConsoleWrite("Variables:" & @LF)
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
			Local $OutFileHandle = FileOpen($PipeFilePath, 0)
			If $OutFileHandle = -1 Then
				ConsoleWrite("Unable to get DOORS Output" & @LF)
				ExitLoop
			Else
				Local $OutputText = FileRead($OutFileHandle)
				Local $NewText = StringTrimLeft($OutputText, StringLen($OldOutputText))
				If $NewText <> "" Then
;~ 					ConsoleWrite("{" & $NewText & "}" & @LF)
					Local $OutputLines = stringsplit($NewText, @CRLF, 1)
					Local $LineIndex
					For $LineIndex = 1 To $OutputLines[0]
						If $OutputLines[$LineIndex] = "" Then
							; Don't pipe empty line from Trace
							If StringLeft($DxlMode, 5) <> "Trace" and $LineIndex < $OutputLines[0] Then
								ConsoleWrite($OutputLines[$LineIndex] & @LF)
							EndIf
						Else
							If StringLeft($DxlMode, 5) = "Trace" Then
								; Don't pipe <Line:...> from Trace
								If StringLeft($OutputLines[$LineIndex], 1) = "<" Then
									If StringLeft($OutputLines[$LineIndex], 6) <> "<Line:" Then
										If $TraceAllLines Or StringLeft($OutputLines[$LineIndex], StringLen($ScriptFile) + 2) = "<" & $ScriptFile & ":" Then
											ConsoleWrite(AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex]) & @LF)
										EndIf
									EndIf
								EndIf
							Else
								ConsoleWrite(AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex]) & @LF)
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
				WinClose($DxlInteractionWindow)
			EndIf
		EndIf
		
		; Pipe the remaining output
		Local $OutFileHandle = FileOpen($PipeFilePath, 0)
		If $OutFileHandle = -1 Then
			ConsoleWrite("Unable to get DOORS Output" & @LF)
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
							ConsoleWrite($OutputLines[$LineIndex] & @LF)
						EndIf
					Else
						If StringLeft($DxlMode, 5) = "Trace" Then
							; Don't pipe <Line:...> from Trace
							If StringLeft($OutputLines[$LineIndex], 1) = "<" Then
								If StringLeft($OutputLines[$LineIndex], 6) <> "<Line:" Then
									If $TraceAllLines Or StringLeft($OutputLines[$LineIndex], StringLen($ScriptFile) + 2) = "<" & $ScriptFile & ":" Then
										ConsoleWrite(AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex]) & @LF)
									EndIf
								EndIf
							EndIf
						Else
							ConsoleWrite(AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex]) & @LF)
						EndIf
					EndIf
				Next
			EndIf
			$OldOutputText = $OutputText
		EndIf
		FileClose($OutFileHandle)

		; Delete output file
		FileDelete($OutFile)
		
		; Report Closed Error Popups
		If $RuntimeError Or WinExists($RuntimeErrorWindow) Then
			ControlClick($RuntimeErrorWindow, "", "[CLASS:Button; INSTANCE:1]")
			ConsoleWrite(@LF)
			ConsoleWrite("++++++++++++++++++" & @LF)
			ConsoleWrite("+ Runtime Error! +" & @LF)
			ConsoleWrite("++++++++++++++++++" & @LF)
		EndIf
		If $DiagnosticLog Or (WinExists($DiagnosticLogWindow) And BitAnd(WinGetState($DiagnosticLogWindow), 2)) Then
			ControlClick($DiagnosticLogWindow, "", "[CLASS:Button; INSTANCE:1]")
			ConsoleWrite(@LF)
			ConsoleWrite("******************" & @LF)
			ConsoleWrite("* Diagnostic Log *" & @LF)
			ConsoleWrite("******************" & @LF)
		EndIf
		If $CppError Or WinExists($CppErrorWindow) Then
			ControlClick($CppErrorWindow, "", "[CLASS:Button; INSTANCE:1]")
			ConsoleWrite(@LF)
			ConsoleWrite("##################" & @LF)
			ConsoleWrite("# MS V C++ Error #" & @LF)
			ConsoleWrite("##################" & @LF)
		EndIf
		
		; Pipe Errors and Warnings
		Local $LogFileHandle = FileOpen($LogFile, 0)
		If $LogFileHandle = -1 Then
			ConsoleWrite("Unable to get DOORS Errors" & @LF)
		Else
			Local $LogFileText = FileRead($LogFileHandle)
			Local $NewText = StringTrimLeft($LogFileText, StringLen($OldLogFileText))
			If $NewText <> "" Then
				; Save Error log containing just the last errors
				Local $NewLogFileHandle = FileOpen($LastLogFile, 2)
				ConsoleWrite(@LF & "Error Log:" & @LF)
				Local $OutputLines = stringsplit($NewText, @CRLF, 1)
				Local $LineIndex
				For $LineIndex = 1 To $OutputLines[0]
					Local $AbsoluteLine = AbsoluteLine($BasePathsArray, $OutputLines[$LineIndex])
					ConsoleWrite($AbsoluteLine & @LF)
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
	ConsoleWrite(@LF)
	
EndIf


; ******************************************************************************************************************* ;


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

Func GetActiveDoorsWindowDetails(ByRef $ModuleFullName, Byref $ModuleType, ByRef $ModuleHwnd)
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
			Local $DoorsExplorer = StringRegExp($DoorsWindows[$i][0], "^[^'].+: (/|\.{3}).+ - DOORS$", 0)
			
			If ($DoorsExplorer) Then
				$DoorsRunning = True
				ExitLoop
			EndIf
			
			; Detect Open Formal Modules from their titles
			Local $DoorsModule = StringRegExp($DoorsWindows[$i][0], "^'(.+)' current .+ in (.+) \((Formal|Link) module\) - DOORS$", 1)
			
			If (UBound($DoorsModule) == 3) Then
				$DoorsRunning = True
				Local $ModuleName = $DoorsModule[0]
				Local $FolderName = $DoorsModule[1]
				$ModuleType = $DoorsModule[2]
				$ModuleFullName = $FolderName & "/" & $ModuleName
				$ModuleHwnd = $DoorsWindows[$i][1]
				ExitLoop
			EndIf
		EndIf
	Next
	Return $DoorsRunning
EndFunc

Func ClearFile($sFilename)
    If FileExists($sFilename) Then
		FileDelete($sFilename)
		Local $FileHandle = FileOpen($sFilename, 2)
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
