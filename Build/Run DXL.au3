#cs ----------------------------------------------------------------------------

Author:
	Adam Cadamally

Description:
	Run DOORS Code

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

Opt("WinSearchChildren", 1)		;0=no, 1=search children also
Opt("MustDeclareVars", 1)		;0=no, 1=yes
;~ Opt("RunErrorsFatal", 0)		;1=fatal, 0=silent set @error is in the script


Local $oErrorHandler = ObjEvent("AutoIt.Error","ComErrorHandler")    ; Initialize a COM error handler
; This is my custom defined error handler
Func ComErrorHandler($oMyError)
	Return
	Msgbox(0,"AutoItCOM Error","We intercepted a COM Error !"    & @LF  & @LF & _
             "err.description is: " & @TAB & $oMyError.description  & @LF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @LF & _
             "err.number is: "       & @TAB & hex($oMyError.number,8)  & @LF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @LF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @LF & _
             "err.source is: "       & @TAB & $oMyError.source       & @LF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @LF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext _
            )
Endfunc


; Check if any Arguments were passed
If $CmdLine[0] < 4 Then
	MsgBox(0, "Wrong Parameters", "Wrong Parameters: " & $CmdLine[0])
	Exit
EndIf

; Get Arguments
Local $ScriptFile = $CmdLine[1]
Local $ModuleFullName = $CmdLine[2]
Local $OutFile = $CmdLine[3]
Local $DxlMode = $CmdLine[4]

Local $TraceAllLines = (StringRight($DxlMode, 7) = "Verbose")

Local $DxlInteractionWindow = "DXL Interaction - DOORS"
Local $DxlOpen = WinExists($DxlInteractionWindow) And BitAnd(WinGetState($DxlInteractionWindow), 2)

Local $IncludeString = "#include <" & $ScriptFile & ">;" & @LF
Local $EscapedInclude = StringReplace($IncludeString, "\", "\\")

Local $EscapedOutFile = StringReplace($OutFile, "\", "\\")
Local $PrintRepurposeCode =  "// Repurpose Print" & @LF & _
	"Stream SublimeText2_PrintStream = write(""" & $EscapedOutFile & """, CP_UTF8)" & @LF & _
	"void SublimeText2_DxlPrint(string s) { print(s) }" & @LF & _
	"void print(string s)" & @LF & _
	"{" & @LF & _
	@TAB & "if(" & StringLower($DxlOpen) & ") { SublimeText2_DxlPrint(s) }" & @LF & _
	@TAB & "if(!null(s)) {" & @LF & _
	@TAB & @TAB & "SublimeText2_PrintStream << s" & @LF & _
	@TAB & @TAB & "flush(SublimeText2_PrintStream)" & @LF & _
	@TAB & "}" & @LF & _
	"}" & @LF & _
	'void print(bool b)		{ print(b "\n") }' & @LF & _
	'void print(char c)		{ print(c "\n") }' & @LF & _
	'void print(Date d)		{ print(d "\n") }' & @LF & _
	'void print(int i)		{ print(i "\n") }' & @LF & _
	'void print(real r)		{ print(r "\n") }' & @LF

Local $SetModuleCode = ""
If $ModuleFullName <> "" Then
	; Set the current module
	$SetModuleCode = "// Set Module" & @LF & _
		'{' & @LF & _
		@TAB & 'Item oItem = item("' & $ModuleFullName & '"); ' & @LF & _
		@TAB & 'if(!null(oItem)) {' & @LF & _
		@TAB & @TAB & 'Module oModule = module(oItem);' & @LF & _
		@TAB & @TAB & 'if(!null(oModule)) {' & @LF & _
		@TAB & @TAB & @TAB & '(current ModuleRef__) = oModule;' & @LF & _
		@TAB & @TAB & '}' & @LF & _
		@TAB & '}' & @LF & _
		'}' & @LF
EndIf

Local $ContextCode = "// Print DXL Context" & @LF & _
	'{' & @LF & _
	@TAB & 'print("DOORS Database: " (getDatabaseName()) "\n");' & @LF & _
	@TAB & 'Module oModule = current();' & @LF & _
	@TAB & 'if(null(oModule)) {' & @LF & _
	@TAB & @TAB & 'print("Current Module:\n");' & @LF & _
	@TAB & '} else {' & @LF & _
	@TAB & @TAB & 'print("Current Module: [" (type(oModule)) "] " (fullName(oModule)) "\n");' & @LF & _
	@TAB & '}' & @LF & _
	@TAB & 'User oUser = find();' & @LF & _
	@TAB & 'string sUserName = oUser.name;' & @LF & _
	@TAB & 'print("Active Account: " sUserName "\n\n");' & @LF & _
	'}' & @LF

Local $DebugInclude = '#include <' & @ScriptDir & '\Debug.inc>;'

Local $EscapedTraceFile = StringReplace($ScriptFile, "\", "\\")
If $TraceAllLines Then
	$EscapedTraceFile = ""
EndIf

Local $Code = $IncludeString
Switch $DxlMode
	Case "CheckAllocationLeak"
		; Check Final Allocations
		; TODO: Overload functions: halt, show, block etc
		$Code = $DebugInclude & @LF
		$Code = $Code & $IncludeString & @LF
		$Code = $Code & 'print("Final Allocated Object Count : " Debug_GetAllocatedObjectCount() "\n");' & @LF
	Case "TraceAllocations", "TraceAllocationsVerbose"
		; Log Allocations
		$Code = $DebugInclude & @LF
		$Code = $Code & 'Debug_Logging(false, true, false, "' & $EscapedTraceFile & '", "' & $EscapedTraceFile & '");' & @LF
		$Code = $Code & $IncludeString & @LF
		$Code = $Code & 'Debug_Logging(false, false, false, "' & $EscapedTraceFile & '", "' & $EscapedTraceFile & '");' & @LF
	Case "TraceExecution", "TraceExecutionVerbose"
		; Log Calls
		$Code = $DebugInclude & @LF
		$Code = $Code & 'Debug_Logging(true, false, true, "' & $EscapedTraceFile & '", "' & $EscapedTraceFile & '");' & @LF
		$Code = $Code & $IncludeString & @LF
		$Code = $Code & 'Debug_Logging(false, false, true, "' & $EscapedTraceFile & '", "' & $EscapedTraceFile & '");' & @LF
	Case "TraceDelays", "TraceDelaysVerbose"
		; Log Calls
		$Code = $DebugInclude & @LF
		$Code = $Code & 'Debug_Logging(true, false, false, "' & $EscapedTraceFile & '", "' & $EscapedTraceFile & '");' & @LF
		$Code = $Code & $IncludeString & @LF
		$Code = $Code & 'Debug_Logging(false, false, false, "' & $EscapedTraceFile & '", "' & $EscapedTraceFile & '");' & @LF
	Case "TraceVariables", "TraceVariablesVerbose"
		; Trace DXL
		$Code = 'startDXLTracing_("C:\\DxlVariables.log");' & @LF
		$Code = $Code & $IncludeString & @LF
		$Code = $Code & 'stopDXLTracing_();' & @LF
EndSwitch

Local $PostfixCode = @LF & "close(SublimeText2_PrintStream)" & @LF

;~ 	Local $FullCode = $PrintRepurposeCode & $SetModuleCode & $IncludeString & $PostfixCode
Local $FullCode = $PrintRepurposeCode & $SetModuleCode & $ContextCode & $Code

If $DxlOpen Then
	
	; Save the current code
	Local $OldCode = ControlGetText($DxlInteractionWindow, "", "[CLASS:RICHEDIT50W; INSTANCE:2]")
	
	; Write the code in DOORS window
	ControlSetText($DxlInteractionWindow, "", "[CLASS:RICHEDIT50W; INSTANCE:2]", $FullCode)

	; Click the Run Button
	ControlClick($DxlInteractionWindow, "", "[CLASS:Button; INSTANCE:8]")
	
	; Restore the saved code
	ControlSetText($DxlInteractionWindow, "", "[CLASS:RICHEDIT50W; INSTANCE:2]", $OldCode)
Else
	; Run the Code via COM
	Local $ObjDoors = ObjCreate("DOORS.application")
	$ObjDoors.Result = ""
	$ObjDoors.runStr($FullCode)
	$ObjDoors = 0
EndIf
