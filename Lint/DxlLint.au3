#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.2.12.1
	Author:         Adam Cadamally

	Script Function:
	Lint DOORS Code

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
		
	; Find Last Active Formal Module
	Local $DoorsWindows = WinList("[CLASS:DOORSWindow]")
	Local $DoorsRunning = $DoorsWindows[0][0] > 0
	
	; Existance Check
	If $DoorsRunning Then
		If $Switch = "-v" Then
			ConsoleWrite("DOORS is Running")
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
	EndIf
EndIf