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

#include <File.au3>

Local $oErrorHandler = ObjEvent("AutoIt.Error","ComErrorHandler")    ; Initialize a COM error handler
; This is my custom defined error handler
Func ComErrorHandler($oMyError)
	Return
	ConsoleWrite("We intercepted a COM Error !"    & @CRLF  & @CRLF & _
             "err.description is: " & @TAB & $oMyError.description  & @CRLF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
             "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext    & @CRLF & _
            )
Endfunc


; Check if any Arguments were passed
If $CmdLine[0] < 2 Then
	ConsoleWrite("Wrong Parameters" & @CRLF)
Else
	; Get Full File Path from Arguments
	Local $ScriptFile = $CmdLine[1]
	Local $DxlMode = $CmdLine[2]
	
	; Find the Sublime Text Window
	Local $Windows = WinList("[CLASS:PX_WINDOW_CLASS]")
	Local $ActiveWindow = 0
	
	; They are in the order they were last activated
	For $i = 1 To $Windows[0][0]
		; Only visble windows that have a title
		If $Windows[$i][0] <> "" And IsVisible($Windows[$i][1]) Then
			$ActiveWindow = $Windows[$i][1]
			ExitLoop
		EndIf
	Next
	
	; Find Last Active Formal Module
	Local $DoorsWindows = WinList("[CLASS:DOORSWindow]")
	Local $DoorsRunning = false
	Local $ModuleName = ""
	Local $FolderName = ""
	Local $ModuleFullName = ""
	Local $ModuleType = ""
	Local $ModuleHwnd = 0
	
	; They are in the order they were last activated
	For $i = 1 To $DoorsWindows[0][0]
		; Only visble windows that have a title
		If $DoorsWindows[$i][0] <> "" And IsVisible($DoorsWindows[$i][1]) Then
			
			; Detect Doors Explorer from the title
			Local $DoorsExplorer = StringRegExp($DoorsWindows[$i][0], "^[^'].+: (/|\.{3}).+ - DOORS$", 0)
			
			If ($DoorsExplorer) Then
				$DoorsRunning = true
				ExitLoop
			EndIf
			
			; Detect Open Formal Modules from their titles
			Local $DoorsModule = StringRegExp($DoorsWindows[$i][0], "^'(.+)' current .+ in (.+) \((Formal|Link) module\) - DOORS$", 1)
			
			If (UBound($DoorsModule) == 3) Then
				$DoorsRunning = true
				$ModuleName = $DoorsModule[0]
				$FolderName = $DoorsModule[1]
				$ModuleType = $DoorsModule[2]
				$ModuleFullName = $FolderName & "/" & $ModuleName
				$ModuleHwnd = $DoorsWindows[$i][1]
				ExitLoop
			EndIf
		EndIf
	Next
	
	
	If Not $DoorsRunning Then
		ConsoleWrite("DOORS is not running." & @CRLF)
	Else
		; Get Object using it's class name
		; Doors must be running for this to be successfull
		Local $ObjDoors = ObjCreate("DOORS.application")

		If @error Then
			ConsoleWrite("Error connecting to Doors" & @CRLF & "Error Code: " & Hex(@error, 8) & @CRLF)

			If Not IsObj($ObjDoors) Then
				ConsoleWrite("Failed to obtain DOORS windows. Error: " & @error)
				Exit
			EndIf
		 Else
			
			; Make the DXL include statement
			Local $IncludeString = "#include <" & $ScriptFile & ">;" & @CRLF
	
			Local $EscapedInclude = StringReplace($IncludeString, "\", "\\")
			Local $TestCode = 'oleSetResult(checkDXL("' & $EscapedInclude & '"))'
			
			; Test the code
			Local $ParseTime = TimerInit()
			$ObjDoors.Result = ""
			$ObjDoors.runStr($TestCode)
			$ParseTime = TimerDiff($ParseTime)
			$ParseTime = StringLeft($ParseTime, StringInStr($ParseTime, ".") -1)
			ConsoleWrite("DXL Code Parsed in: " & $ParseTime & " milliseconds" & @CRLF)
			
			Local $DXLOutputText = $ObjDoors.Result
			If $DXLOutputText <> "" Then
				ConsoleWrite("Parse Errors:" & @CRLF & $DXLOutputText & @CRLF)
			Else
			
				$ObjDoors.runStr('oleSetResult(getenv("DOORSLOGFILE"))')
				Local $LogFile = $ObjDoors.Result
				If Not $LogFile Then
					ConsoleWrite("'DOORSLOGFILE' is not defined for Warning and Error logging" & @CRLF)
				EndIf

				Local $LogFileHandle = FileOpen($LogFile, 0)
				Local $OldLogFileNext = FileRead($LogFileHandle)
				FileClose($LogFileHandle)
				
				Local $OutFile = _TempFile()
				
				$ObjDoors.runStr('oleSetResult(getDatabaseName())')
				Local $DatabaseName = $ObjDoors.Result
				ConsoleWrite("DOORS Database: " & $DatabaseName & @CRLF)
				
				If $ModuleFullName <> "" Then
				   ConsoleWrite("Running code in "& $ModuleType & " Module:" & @CRLF & @TAB & $ModuleFullName & @CRLF & @CRLF)
				Else
				   ConsoleWrite("Running code in Doors Explorer..." & @CRLF & @CRLF)
				EndIf
				
				; Remember if the DXL Interaction window is already open
				Local $DxlInteractionWindow = "DXL Interaction - DOORS"
				Local $DxlOpen = WinExists($DxlInteractionWindow) And BitAnd(WinGetState($DxlInteractionWindow), 2)
				Local $OldCode = ""
				If $DxlOpen Then
					$OldCode = ControlGetText($DxlInteractionWindow, "", "[CLASS:RICHEDIT50W; INSTANCE:2]")
				EndIf
				ShellExecute("Run DXL.exe", '"' & $IncludeString & '" "' & $ModuleFullName & '" "' & $OutFile & '" ' & $DxlOpen & ' ' & $DxlMode, @ScriptDir)
				Sleep($ParseTime + 500)

				; Error Window Titles
				Local $CppErrorWindow = "Microsoft Visual C++ Runtime Library"
				Local $DiagnosticLogWindow = "Diagnostic Log - DOORS"
				Local $RuntimeErrorWindow = "DOORS report"

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
					
					; Possible Diagnostic Log Window
					$DiagnosticLog = WinExists($DiagnosticLogWindow) And BitAnd(WinGetState($DiagnosticLogWindow), 2)
					
					; Possible Runtime Error Window
					$RuntimeError = WinExists($RuntimeErrorWindow)
					
					If $CppError Then
						; Close C++ Error message box
						ControlClick($CppErrorWindow, "", "[CLASS:Button; INSTANCE:1]")
					EndIf
					
					If $DiagnosticLog Then
						; Close Runtime Error message box
						ControlClick($DiagnosticLogWindow, "", "[CLASS:Button; INSTANCE:1]")
					EndIf
					
					If $RuntimeError Then
						; Close Runtime Error message box
						ControlClick($RuntimeErrorWindow, "", "[CLASS:Button; INSTANCE:1]")
					EndIf
					
					; Pipe the new output
					Local $OutFileHandle = FileOpen($OutFile, 256)
					Local $OutputText = FileRead($OutFileHandle)
					ConsoleWrite(StringTrimLeft($OutputText, StringLen($OldOutputText)))
					$OldOutputText = $OutputText
					FileClose($OutFileHandle)
					
				WEnd
				
				; Possible DXL Interaction Window
				If $DxlOpen Then
					ControlSetText($DxlInteractionWindow, "", "[CLASS:RICHEDIT50W; INSTANCE:2]", $OldCode)
				Else
					Local $DxlInteractionWindow = "DXL Interaction - DOORS"
					If WinExists($DxlInteractionWindow) Then
						WinClose($DxlInteractionWindow)
					EndIf
				EndIf
				
				; Pipe the remaining output
				Local $OutFileHandle = FileOpen($OutFile, 256)
				If $OutFileHandle = -1 Then
					ConsoleWrite("Unable to get DOORS Output" & @LF)
				Else
					Local $OutputText = FileRead($OutFileHandle)
					ConsoleWrite(StringTrimLeft($OutputText, StringLen($OldOutputText)))
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
				Local $LogFileHandle = FileOpen($LogFile, 256)
				Local $LogFileNext = FileRead($LogFileHandle)
				If StringLen($LogFileNext) > StringLen($OldLogFileNext) Then
					ConsoleWrite(@LF & "Error Log:" & @LF)
				EndIf
				ConsoleWrite(StringTrimLeft($LogFileNext, StringLen($OldLogFileNext)))
				FileClose($LogFileHandle)
				
				; Save Error log containing just the last errors
				Local $NewLogFileHandle = FileOpen(StringTrimRight($LogFile, 4) & " Last.log", 2)
				FileWrite($NewLogFileHandle, StringTrimLeft($LogFileNext, StringLen($OldLogFileNext)))
				FileClose($NewLogFileHandle)
				
				; Reactivate selected module because errors will activate explorer 
				; The code would then be run in DOORS Explorer the next time
				If StringLen($LogFileNext) > StringLen($OldLogFileNext) Then
					If $ModuleHwnd And $ActiveWindow Then
						WinSetOnTop($ActiveWindow, "", 1)
						WinActivate($ModuleHwnd)
						WinSetOnTop($ActiveWindow, "", 0)	; Better to restore previous state here
						WinActivate($ActiveWindow)
					EndIf
				EndIf
				
			EndIf
		   
		EndIf
		
		$ObjDoors = 0
		ConsoleWrite(@LF & @LF)
		
	EndIf
EndIf


; ******************************************************************************************************************* ;


Func ClearFile($sFilename)
    If FileExists($sFilename) Then
		FileDelete($sFilename)
		Local $FileHandle = FileOpen($sFilename, 2)
		FileClose($FileHandle)
	EndIf
EndFunc

Func IsVisible($handle)
	If BitAND(WinGetState($handle), 2) Then
		Return 1
	Else
		Return 0
	EndIf
EndFunc   ;==>IsVisible

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
