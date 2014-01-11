#cs ----------------------------------------------------------------------------

Author:
	Adam Cadamally

Description:
	Lint DOORS Code

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

Opt("MustDeclareVars", 1)		;0=no, 1=yes

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
If $CmdLine[0] < 1 Then
	ConsoleWrite("Wrong Parameters" & @CRLF)
Else	
	; Get Full File Path from Arguments
	Local $Switch = $CmdLine[1]
	
	; Get Doors Process Properties
	Local $DoorsProcess = _ProcessListProperties("doors.exe")
	Local $DoorsRunning = $DoorsProcess[0][0] > 0
	
	If $DoorsRunning Then

		ConsoleWrite("DOORS Process Running"& @CRLF)

		; Get Doors CPU usage
		Local $DoorsCPU = $DoorsProcess[1][6]
		Local $DoorsIdle = $DoorsCPU < 5
		
		ConsoleWrite("DOORS CPU: " & $DoorsCPU & @CRLF)
		If $DoorsIdle Then
			
			ConsoleWrite("DOORS is Idle"& @CRLF)
			
			If $Switch = "-v" Then
				ConsoleWrite("DOORS is Idle")
			Else
				Local $ScriptFile = $Switch
				
				; Connect to DOORS via COM
				Local $ObjDoors = ObjCreate("DOORS.application")
				If @error Then
					ConsoleWrite("Error connecting to Doors" & @CRLF & "Error Code: " & Hex(@error, 8) & @CRLF)

					If Not IsObj($ObjDoors) Then
						ConsoleWrite("Failed to obtain DOORS windows. Error: " & @error)
						Exit
					EndIf
				 Else
					; Get Full File Path from Arguments
					Local $ScriptFile = $CmdLine[1]
					
					; Make the DXL include statement
					Local $IncludeString = "#include <" & $ScriptFile & ">"
					
					Local $EscapedInclude = StringReplace($IncludeString, "\", "\\")
					Local $TestCode = 'oleSetResult(checkDXL("' & $EscapedInclude & '"))'
					
					; Test the code
					$ObjDoors.Result = ""
					$ObjDoors.runStr($TestCode)
					
					Local $DXLOutputText = $ObjDoors.Result
					If $DXLOutputText <> "" Then
						ConsoleWrite($DXLOutputText)
					EndIf
				EndIf
			EndIf
		Else
			ConsoleWrite("DOORS is Busy: " & @CRLF)
		EndIf
	Else
		ConsoleWrite("DOORS is not Running"& @CRLF)
	EndIf
EndIf

;===============================================================================
; Function Name:    _ProcessListProperties()
; Description:      Get various properties of a process, or all processes
; Call With:        _ProcessListProperties( [$Process [, $sComputer]] )
; Parameter(s):     (optional) $Process - PID or name of a process, default is all
;                   (optional) $sComputer - remote computer to get list from, default is local
; Requirement(s):   AutoIt v3.2.4.9+
; Return Value(s):  On Success - Returns a 2D array of processes, as in ProcessList()
;             with additional columns added:
;             [0][0] - Number of processes listed (can be 0 if no matches found)
;             [1][0] - 1st process name
;             [1][1] - 1st process PID
;             [1][2] - 1st process Parent PID
;             [1][3] - 1st process owner
;             [1][4] - 1st process priority (0 = low, 31 = high)
;             [1][5] - 1st process executable path
;             [1][6] - 1st process CPU usage
;             [1][7] - 1st process memory usage
;             ...
;             [n][0] thru [n][7] - last process properties
; On Failure:       Returns array with [0][0] = 0 and sets @Error to non-zero (see code below)
; Author(s):        PsaltyDS at <a href='http://www.autoitscript.com/forum' class='bbc_url' title=''>http://www.autoitscript.com/forum</a>
; Notes:            If a numeric PID or string process name is provided and no match is found,
;             then [0][0] = 0 and @error = 0 (not treated as an error, same as ProcessList)
;           This function requires admin permissions to the target computer.
;           All properties come from the Win32_Process class in WMI.
;===============================================================================
Func _ProcessListProperties($Process = "", $sComputer = ".")
    Local $sUserName, $sMsg, $sUserDomain, $avProcs
    If $Process = "" Then
        $avProcs = ProcessList()
    Else
        $avProcs = ProcessList($Process)
    EndIf

    ; Return for no matches
    If $avProcs[0][0] = 0 Then Return $avProcs

    ; ReDim array for additional property columns
    ReDim $avProcs[$avProcs[0][0] + 1][8]

    ; Connect to WMI and get process objects
    Local $oWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & $sComputer & "\root\cimv2")
    If IsObj($oWMI) Then
        ; Get collection of all processes from Win32_Process
        Local $colProcs = $oWMI.ExecQuery ("select * from win32_process")
        If IsObj($colProcs) Then
            ; For each process...
            For $oProc In $colProcs
                ; Find it in the array
                For $n = 1 To $avProcs[0][0]
                    If $avProcs[$n][1] = $oProc.ProcessId Then

                        ; [n][2] = Parent PID
                        $avProcs[$n][2] = $oProc.ParentProcessId
                        ; [n][3] = Owner
                        If $oProc.GetOwner ($sUserName, $sUserDomain) = 0 Then $avProcs[$n][3] = $sUserDomain & "\" & $sUserName
                        ; [n][4] = Priority
                        $avProcs[$n][4] = $oProc.Priority
                        ; [n][5] = Executable path
                        $avProcs[$n][5] = $oProc.ExecutablePath

                        ExitLoop
                    EndIf
                Next
            Next
        Else
            SetError(2) ; Error getting process collection from WMI
        EndIf

        ; Get collection of all processes from Win32_PerfFormattedData_PerfProc_Process
        ; Have to use an SWbemRefresher to pull the collection, or all Perf data will be zeros
        Local $oRefresher = ObjCreate("WbemScripting.SWbemRefresher")
        $colProcs = $oRefresher.AddEnum ($oWMI, "Win32_PerfFormattedData_PerfProc_Process" ).objectSet
        $oRefresher.Refresh
        
        ; Time delay before calling refresher
        Local $iTime = TimerInit()
        Do
            Sleep(10)
        Until TimerDiff($iTime) > 100
        $oRefresher.Refresh
        
        ; Get PerfProc data
        For $oProc In $colProcs
            ; Find it in the array
            For $n = 1 To $avProcs[0][0]
                If $avProcs[$n][1] = $oProc.IDProcess Then
                    $avProcs[$n][6] = $oProc.PercentProcessorTime
                    $avProcs[$n][7] = $oProc.WorkingSet
                    ExitLoop
                EndIf
            Next
        Next
    Else
        SetError(1) ; Error connecting to WMI
    EndIf

    ; Return array
    Return $avProcs
EndFunc   ;==>_ProcessListProperties